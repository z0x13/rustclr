use alloc::{format, string::{String, ToString}, vec};
use obfstr::obfstr as s;
use spin::Mutex;
use windows::core::BSTR;
use windows::Win32::Foundation::VARIANT_BOOL;

use crate::com::_Assembly;
use crate::error::{ClrError, Result};
use crate::variant::{create_safe_args, create_string_array_variant};
use crate::{com, Invocation, RustClrEnv};

struct CompiledEnv {
    bootstrap_type: com::_Type,
    automation: _Assembly,
    clr: RustClrEnv,
}

// SAFETY: COM pointers can be transferred between threads. The CLR's COM interop
// handles apartment marshaling internally. AddRef/Release are thread-safe per COM spec.
// Access is serialized by the Mutex, preventing concurrent use.
unsafe impl Send for CompiledEnv {}

static COMPILED_ENV: Mutex<Option<CompiledEnv>> = Mutex::new(None);

fn get_custom_host_code() -> String {
    String::from(s!(r#"
using System;
using System.Text;
using System.Reflection;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;

public static class HostBootstrap {
    private static bool _initialized = false;
    private static Assembly _smaAssembly = null;
    private static Assembly _compiledAsm = null;
    private static string _lastError = null;
    private static long _nextId = 1;

    private class InstanceInfo {
        public Runspace Runspace;
        public CaptureHost Host;
        public IntPtr OriginalStdOut;
        public IntPtr PipeReadOut;
        public IntPtr PipeWriteOut;
    }
    private static Dictionary<long, InstanceInfo> _instances = new Dictionary<long, InstanceInfo>();

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetStdHandle(int nStdHandle, IntPtr hHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CreatePipe(out IntPtr hReadPipe, out IntPtr hWritePipe, ref SECURITY_ATTRIBUTES lpPipeAttributes, uint nSize);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool ReadFile(IntPtr hFile, byte[] lpBuffer, uint nNumberOfBytesToRead, out uint lpNumberOfBytesRead, IntPtr lpOverlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool PeekNamedPipe(IntPtr hNamedPipe, IntPtr lpBuffer, uint nBufferSize, IntPtr lpBytesRead, out uint lpTotalBytesAvail, IntPtr lpBytesLeftThisMessage);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    [StructLayout(LayoutKind.Sequential)]
    private struct SECURITY_ATTRIBUTES {
        public int nLength;
        public IntPtr lpSecurityDescriptor;
        public bool bInheritHandle;
    }

    private const int STD_OUTPUT_HANDLE = -11;

    public static string GetLastError() { return _lastError ?? ""; }
    public static void SetCompiledAssembly(Assembly asm) { _compiledAsm = asm; }

    public static void Initialize() {
        if (_initialized) return;
        _initialized = true;

        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies()) {
            if (asm.GetName().Name == "System.Management.Automation") {
                _smaAssembly = asm;
                break;
            }
        }

        AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
    }

    private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args) {
        if (args.Name.Contains(".resources")) return null;
        if (args.Name.StartsWith("System.Management.Automation,") && _smaAssembly != null) {
            return _smaAssembly;
        }
        return null;
    }

    public static long CreateRunspace() {
        try {
            Initialize();
            if (_smaAssembly == null) {
                _lastError = "SMA assembly not found";
                return -1;
            }
            if (_compiledAsm == null) {
                _lastError = "Compiled assembly not set";
                return -1;
            }

            var host = new CaptureHost();
            var iss = InitialSessionState.CreateDefault2();
            var runspace = RunspaceFactory.CreateRunspace(host, iss);
            runspace.Open();

            var sa = new SECURITY_ATTRIBUTES();
            sa.nLength = Marshal.SizeOf(sa);
            sa.bInheritHandle = true;
            IntPtr pipeRead, pipeWrite;
            if (!CreatePipe(out pipeRead, out pipeWrite, ref sa, 0)) {
                runspace.Close();
                _lastError = "CreatePipe failed";
                return -1;
            }

            var id = _nextId++;
            _instances[id] = new InstanceInfo {
                Runspace = runspace,
                Host = host,
                OriginalStdOut = GetStdHandle(STD_OUTPUT_HANDLE),
                PipeReadOut = pipeRead,
                PipeWriteOut = pipeWrite
            };

            return id;
        } catch (Exception ex) {
            var inner = ex;
            while (inner.InnerException != null) inner = inner.InnerException;
            _lastError = inner.GetType().FullName + ": " + inner.Message;
            return -1;
        }
    }

    public static object GetRunspace(long id) {
        InstanceInfo info;
        return _instances.TryGetValue(id, out info) ? info.Runspace : null;
    }

    public static void CloseRunspace(long id) {
        InstanceInfo info;
        if (_instances.TryGetValue(id, out info)) {
            try {
                if (info.Runspace != null) {
                    if (info.Runspace.RunspaceStateInfo.State == RunspaceState.Opened) {
                        info.Runspace.Close();
                    }
                    info.Runspace.Dispose();
                }
                if (info.PipeReadOut != IntPtr.Zero) CloseHandle(info.PipeReadOut);
                if (info.PipeWriteOut != IntPtr.Zero) CloseHandle(info.PipeWriteOut);
                // Clear references to help GC
                info.Runspace = null;
                info.Host = null;
            } catch { }
            _instances.Remove(id);

            // Force GC to release runspace resources
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }

    public static void BeginCapture(long id) {
        InstanceInfo info;
        if (_instances.TryGetValue(id, out info) && info.PipeWriteOut != IntPtr.Zero) {
            SetStdHandle(STD_OUTPUT_HANDLE, info.PipeWriteOut);
        }
    }

    public static void EndCapture(long id) {
        InstanceInfo info;
        if (_instances.TryGetValue(id, out info) && info.OriginalStdOut != IntPtr.Zero) {
            SetStdHandle(STD_OUTPUT_HANDLE, info.OriginalStdOut);
        }
    }

    private static string ReadFromPipe(IntPtr pipeRead) {
        var sb = new StringBuilder();
        uint available;
        while (PeekNamedPipe(pipeRead, IntPtr.Zero, 0, IntPtr.Zero, out available, IntPtr.Zero) && available > 0) {
            var buffer = new byte[Math.Min(available, 4096)];
            uint read;
            if (ReadFile(pipeRead, buffer, (uint)buffer.Length, out read, IntPtr.Zero) && read > 0) {
                sb.Append(Encoding.UTF8.GetString(buffer, 0, (int)read));
            } else break;
        }
        return sb.ToString();
    }

    public static void ClearHostOutput(long id) {
        InstanceInfo info;
        if (_instances.TryGetValue(id, out info)) {
            info.Host.ClearOutput();
            ReadFromPipe(info.PipeReadOut);
        }
    }

    public static string GetHostOutput(long id) {
        InstanceInfo info;
        if (!_instances.TryGetValue(id, out info)) return "";
        var sb = new StringBuilder();
        sb.Append(ReadFromPipe(info.PipeReadOut));
        sb.Append(info.Host.GetOutput());
        return sb.ToString();
    }
}

public class CaptureRawUI : PSHostRawUserInterface {
    public override ConsoleColor BackgroundColor { get { return ConsoleColor.Black; } set { } }
    public override ConsoleColor ForegroundColor { get { return ConsoleColor.White; } set { } }
    public override Size BufferSize { get { return new Size(120, 50); } set { } }
    public override Coordinates CursorPosition { get { return new Coordinates(0, 0); } set { } }
    public override int CursorSize { get { return 1; } set { } }
    public override Size MaxPhysicalWindowSize { get { return new Size(120, 50); } }
    public override Size MaxWindowSize { get { return new Size(120, 50); } }
    public override Coordinates WindowPosition { get { return new Coordinates(0, 0); } set { } }
    public override Size WindowSize { get { return new Size(120, 50); } set { } }
    public override string WindowTitle { get { return ""; } set { } }
    public override bool KeyAvailable { get { return false; } }
    public override void FlushInputBuffer() { }
    public override BufferCell[,] GetBufferContents(Rectangle r) { return null; }
    public override KeyInfo ReadKey(ReadKeyOptions o) { return new KeyInfo(); }
    public override void ScrollBufferContents(Rectangle s, Coordinates d, Rectangle c, BufferCell f) { }
    public override void SetBufferContents(Rectangle r, BufferCell f) { }
    public override void SetBufferContents(Coordinates o, BufferCell[,] c) { }
}

public class CaptureUI : PSHostUserInterface {
    private StringBuilder _out;
    private CaptureRawUI _rawUI;
    public CaptureUI(StringBuilder sb) { _out = sb; _rawUI = new CaptureRawUI(); }
    public override PSHostRawUserInterface RawUI { get { return _rawUI; } }
    public override void Write(string value) { _out.Append(value); }
    public override void Write(ConsoleColor f, ConsoleColor b, string value) { _out.Append(value); }
    public override void WriteLine(string value) { _out.AppendLine(value); }
    public override void WriteLine(ConsoleColor f, ConsoleColor b, string value) { _out.AppendLine(value); }
    public override void WriteDebugLine(string m) { _out.AppendLine("DEBUG: " + m); }
    public override void WriteErrorLine(string m) { _out.AppendLine("ERROR: " + m); }
    public override void WriteVerboseLine(string m) { _out.AppendLine("VERBOSE: " + m); }
    public override void WriteWarningLine(string m) { _out.AppendLine("WARNING: " + m); }
    public override void WriteProgress(long id, ProgressRecord r) { }
    public override string ReadLine() { return ""; }
    public override System.Security.SecureString ReadLineAsSecureString() { return new System.Security.SecureString(); }
    public override System.Collections.Generic.Dictionary<string, PSObject> Prompt(string c, string m, System.Collections.ObjectModel.Collection<FieldDescription> d) { return new System.Collections.Generic.Dictionary<string, PSObject>(); }
    public override int PromptForChoice(string c, string m, System.Collections.ObjectModel.Collection<ChoiceDescription> ch, int df) { return df; }
    public override PSCredential PromptForCredential(string c, string m, string u, string t) { return null; }
    public override PSCredential PromptForCredential(string c, string m, string u, string t, PSCredentialTypes at, PSCredentialUIOptions o) { return null; }
}

public class CaptureHost : PSHost {
    private Guid _id;
    private StringBuilder _output;
    private CaptureUI _ui;

    public CaptureHost() {
        _id = Guid.NewGuid();
        _output = new StringBuilder();
        _ui = new CaptureUI(_output);
    }

    public override string Name { get { return "CaptureHost"; } }
    public override Version Version { get { return new Version(1, 0); } }
    public override Guid InstanceId { get { return _id; } }
    public override CultureInfo CurrentCulture { get { return CultureInfo.CurrentCulture; } }
    public override CultureInfo CurrentUICulture { get { return CultureInfo.CurrentUICulture; } }
    public override PSHostUserInterface UI { get { return _ui; } }
    public override void SetShouldExit(int code) { }
    public override void EnterNestedPrompt() { }
    public override void ExitNestedPrompt() { }
    public override void NotifyBeginApplication() { }
    public override void NotifyEndApplication() { }

    public string GetOutput() { return _output.ToString(); }
    public void ClearOutput() { _output.Clear(); }
}
"#))
}

fn compile_env() -> Result<CompiledEnv> {
    let clr = RustClrEnv::new(None)?;
    let mscorlib = clr.app_domain.get_assembly(s!("mscorlib"))?;
    let reflection_assembly = mscorlib.resolve_type(s!("System.Reflection.Assembly"))?;
    let load_partial_name = reflection_assembly.method_signature(s!(
        "System.Reflection.Assembly LoadWithPartialName(System.String)"
    ))?;

    // Load System assembly
    let system_param = create_safe_args(vec![s!("System").into()])?;
    let system_asm = load_partial_name.invoke(None, Some(&system_param))?;

    // Get CSharpCodeProvider type
    let provider_type_result = reflection_assembly.invoke(
        s!("GetType"),
        Some(system_asm.clone()),
        Some(vec![s!("Microsoft.CSharp.CSharpCodeProvider").into()]),
        Invocation::Instance,
    )?;
    let provider_type_ptr = unsafe { provider_type_result.Anonymous.Anonymous.Anonymous.byref };
    if provider_type_ptr.is_null() {
        return Err(ClrError::Msg("CSharpCodeProvider type not found"));
    }
    let provider_type_obj = com::_Type::from_raw(provider_type_ptr)?;

    // Create CSharpCodeProvider instance
    let activator = mscorlib.resolve_type(s!("System.Activator"))?;
    let create_instance =
        activator.method_signature(s!("System.Object CreateInstance(System.Type)"))?;
    let provider_type_variant = (&*provider_type_obj).clone().into();
    let provider_args = create_safe_args(vec![provider_type_variant])?;
    let provider = create_instance.invoke(None, Some(&provider_args))?;

    // Create CompilerParameters
    let params_type_result = reflection_assembly.invoke(
        s!("GetType"),
        Some(system_asm.clone()),
        Some(vec![
            s!("System.CodeDom.Compiler.CompilerParameters").into(),
        ]),
        Invocation::Instance,
    )?;
    let params_type_ptr = unsafe { params_type_result.Anonymous.Anonymous.Anonymous.byref };
    if params_type_ptr.is_null() {
        return Err(ClrError::Msg("CompilerParameters type not found"));
    }
    let params_type = crate::com::_Type::from_raw(params_type_ptr)?;
    let params_type_variant = (&*params_type).clone().into();
    let compiler_params_args = create_safe_args(vec![params_type_variant])?;
    let compiler_params = create_instance.invoke(None, Some(&compiler_params_args))?;

    // Set GenerateInMemory = true
    params_type.invoke(
        s!("set_GenerateInMemory"),
        Some(compiler_params.clone()),
        Some(vec![true.into()]),
        Invocation::Instance,
    )?;

    // Add references
    let get_assemblies = params_type.method_signature(s!(
        "System.Collections.Specialized.StringCollection get_ReferencedAssemblies()"
    ))?;
    let assemblies = get_assemblies.invoke(Some(compiler_params.clone()), None)?;

    let string_collection_result = reflection_assembly.invoke(
        s!("GetType"),
        Some(system_asm.clone()),
        Some(vec![
            s!("System.Collections.Specialized.StringCollection").into(),
        ]),
        Invocation::Instance,
    )?;
    let string_collection_ptr =
        unsafe { string_collection_result.Anonymous.Anonymous.Anonymous.byref };
    if string_collection_ptr.is_null() {
        return Err(ClrError::Msg("StringCollection type not found"));
    }
    let string_collection = crate::com::_Type::from_raw(string_collection_ptr)?;
    let add_method = string_collection.method_signature(s!("Int32 Add(System.String)"))?;

    // Load System.Management.Automation via partial name (version-agnostic)
    let sma_param = create_safe_args(vec![s!("System.Management.Automation").into()])?;
    let sma_asm = load_partial_name.invoke(None, Some(&sma_param))?;
    let sma_ptr = unsafe { sma_asm.Anonymous.Anonymous.Anonymous.byref };
    if sma_ptr.is_null() {
        return Err(ClrError::Msg("System.Management.Automation assembly not found"));
    }
    let automation = _Assembly::from_raw(sma_ptr)?;

    // Get SMA location for reference
    let get_location = reflection_assembly.method_signature(s!("System.String get_Location()"))?;
    let sma_location_result = get_location.invoke(Some(sma_asm), None)?;
    let sma_location = sma_location_result.to_string();

    // Add all references
    for asm in [s!("System.dll"), s!("mscorlib.dll"), s!("System.Core.dll")] {
        let asm_args = create_safe_args(vec![asm.into()])?;
        add_method.invoke(Some(assemblies.clone()), Some(&asm_args))?;
    }
    let sma_location_args = create_safe_args(vec![BSTR::from(sma_location.as_str()).into()])?;
    add_method.invoke(Some(assemblies.clone()), Some(&sma_location_args))?;

    // Compile the C# code
    let source_array = create_string_array_variant(vec![get_custom_host_code()])?;

    let object_type = mscorlib.resolve_type(s!("System.Object"))?;
    let get_type_method = object_type.method_signature(s!("System.Type GetType()"))?;
    let provider_type_obj = get_type_method.invoke(Some(provider.clone()), None)?;
    let provider_type_real = crate::com::_Type::from_raw(unsafe {
        provider_type_obj.Anonymous.Anonymous.Anonymous.byref
    })?;

    let compile_result = provider_type_real.invoke(
        s!("CompileAssemblyFromSource"),
        Some(provider),
        Some(vec![compiler_params, source_array]),
        Invocation::Instance,
    )?;

    // Check for compilation errors
    let compiler_results_result = reflection_assembly.invoke(
        s!("GetType"),
        Some(system_asm),
        Some(vec![
            s!("System.CodeDom.Compiler.CompilerResults").into(),
        ]),
        Invocation::Instance,
    )?;
    let compiler_results_ptr =
        unsafe { compiler_results_result.Anonymous.Anonymous.Anonymous.byref };
    if compiler_results_ptr.is_null() {
        return Err(ClrError::Msg("CompilerResults type not found"));
    }
    let compiler_results_type = crate::com::_Type::from_raw(compiler_results_ptr)?;

    let get_errors = compiler_results_type.method_signature(s!(
        "System.CodeDom.Compiler.CompilerErrorCollection get_Errors()"
    ))?;
    let errors = get_errors.invoke(Some(compile_result.clone()), None)?;

    let icollection = mscorlib.resolve_type(s!("System.Collections.ICollection"))?;
    let get_count = icollection.method_signature(s!("Int32 get_Count()"))?;
    let error_count = unsafe {
        get_count
            .invoke(Some(errors), None)?
            .Anonymous
            .Anonymous
            .Anonymous
            .lVal
    };

    if error_count > 0 {
        return Err(ClrError::Msg("C# host compilation failed"));
    }

    // Get compiled assembly
    let get_compiled_asm = compiler_results_type
        .method_signature(s!("System.Reflection.Assembly get_CompiledAssembly()"))?;
    let compiled_asm_variant = get_compiled_asm.invoke(Some(compile_result), None)?;
    let compiled_asm =
        _Assembly::from_raw(unsafe { compiled_asm_variant.Anonymous.Anonymous.Anonymous.byref })?;

    let bootstrap_type = compiled_asm.resolve_type(s!("HostBootstrap"))?;

    // Store the compiled assembly reference in C# for later use
    let asm_variant = (&*compiled_asm).clone().into();
    bootstrap_type.invoke(
        s!("SetCompiledAssembly"),
        None,
        Some(vec![asm_variant]),
        Invocation::Static,
    )?;

    Ok(CompiledEnv {
        bootstrap_type,
        automation,
        clr,
    })
}

/// Provides a persistent interface for executing PowerShell commands.
///
/// # Example
///
/// ```
/// let pwsh = PowerShell::new()?;
/// let out = pwsh.execute("whoami")?;
/// print!("Output: {}", out);
/// ```
pub struct PowerShell {
    instance_id: i64,
}

impl PowerShell {
    /// Creates a new `PowerShell` with a custom host for output capture.
    pub fn new() -> Result<Self> {
        let mut guard = COMPILED_ENV.lock();
        if guard.is_none() {
            *guard = Some(compile_env()?);
        }
        let env = guard.as_ref().ok_or(ClrError::Msg("PowerShell environment not initialized"))?;

        // Create a new runspace instance (returns ID)
        let create_result =
            env.bootstrap_type
                .invoke(s!("CreateRunspace"), None, None, Invocation::Static)?;
        let instance_id = unsafe { create_result.Anonymous.Anonymous.Anonymous.llVal };

        if instance_id < 0 {
            return Err(ClrError::Msg("CreateRunspace failed"));
        }

        Ok(Self { instance_id })
    }

    /// Executes a PowerShell command and returns its output as a string.
    pub fn execute(&self, command: &str) -> Result<String> {
        let guard = COMPILED_ENV.lock();
        let env = guard.as_ref().ok_or(ClrError::Msg("PowerShell environment not initialized"))?;
        let mscorlib = env.clr.app_domain.get_assembly(s!("mscorlib"))?;

        // Clear previous output
        let _clear_result = env.bootstrap_type.invoke(
            s!("ClearHostOutput"),
            None,
            Some(vec![self.instance_id.into()]),
            Invocation::Static,
        )?;

        // Get runspace for this instance
        let runspace = env.bootstrap_type.invoke(
            s!("GetRunspace"),
            None,
            Some(vec![self.instance_id.into()]),
            Invocation::Static,
        )?;

        // Create pipeline
        let runspace_type = env
            .automation
            .resolve_type(s!("System.Management.Automation.Runspaces.Runspace"))?;
        let create_pipeline = runspace_type.method_signature(s!(
            "System.Management.Automation.Runspaces.Pipeline CreatePipeline()"
        ))?;
        let pipe = create_pipeline.invoke(Some(runspace.clone()), None)?;

        // Add script (simple wrapper, no try-catch needed with InvokeAsync)
        let pipeline_type = env
            .automation
            .resolve_type(s!("System.Management.Automation.Runspaces.Pipeline"))?;
        let get_commands =
            pipeline_type.invoke(s!("get_Commands"), Some(pipe.clone()), None, Invocation::Instance)?;

        let script = format!("& {{ {command} }} | {out}", out = s!("Out-String"));
        let command_collection = env.automation.resolve_type(s!(
            "System.Management.Automation.Runspaces.CommandCollection"
        ))?;
        let add_script =
            command_collection.method_signature(s!("Void AddScript(System.String)"))?;
        let script_args = create_safe_args(vec![BSTR::from(script.as_str()).into()])?;
        let _add_result = add_script.invoke(Some(get_commands), Some(&script_args))?;

        // Begin capturing native command output
        let _begin_capture = env.bootstrap_type.invoke(
            s!("BeginCapture"),
            None,
            Some(vec![self.instance_id.into()]),
            Invocation::Static,
        )?;

        // Use InvokeAsync - doesn't throw on script errors
        let _invoke_result = pipeline_type.invoke(s!("InvokeAsync"), Some(pipe.clone()), None, Invocation::Instance)?;

        // Read output via get_Output().ReadToEnd()
        let output_reader =
            pipeline_type.invoke(s!("get_Output"), Some(pipe.clone()), None, Invocation::Instance)?;

        let ps_reader_type = env.automation.resolve_type(s!(
            "System.Management.Automation.Runspaces.PipelineReader`1[System.Management.Automation.PSObject]"
        ))?;
        let read_to_end = ps_reader_type.method_signature(s!(
            "System.Collections.ObjectModel.Collection`1[System.Management.Automation.PSObject] ReadToEnd()"
        ))?;
        let output_collection = read_to_end.invoke(Some(output_reader), None)?;

        // End capturing - restore original stdout
        let _ = env.bootstrap_type.invoke(
            s!("EndCapture"),
            None,
            Some(vec![self.instance_id.into()]),
            Invocation::Static,
        );

        // Get pipeline output
        let mut result = String::new();
        let output_ptr = unsafe { &output_collection.Anonymous.Anonymous.Anonymous.punkVal };
        if output_ptr.is_some() {
            let icollection = mscorlib.resolve_type(s!("System.Collections.ICollection"))?;
            let get_count = icollection.method_signature(s!("Int32 get_Count()"))?;
            let count_var = get_count.invoke(Some(output_collection.clone()), None)?;
            let count = unsafe { count_var.Anonymous.Anonymous.Anonymous.lVal };

            if count > 0 {
                let ilist = mscorlib.resolve_type(s!("System.Collections.IList"))?;
                let get_item = ilist.method_signature(s!("System.Object get_Item(Int32)"))?;
                let object_type = mscorlib.resolve_type(s!("System.Object"))?;
                let to_string = object_type.method_signature(s!("System.String ToString()"))?;

                for i in 0..count {
                    let item_args = create_safe_args(vec![i.into()])?;
                    let item = get_item.invoke(Some(output_collection.clone()), Some(&item_args))?;
                    if unsafe { &item.Anonymous.Anonymous.Anonymous.punkVal }.is_some() {
                        let item_str = to_string.invoke(Some(item), None)?;
                        let s = item_str.to_string();
                        if !s.is_empty() {
                            if !result.is_empty() {
                                result.push('\n');
                            }
                            result.push_str(&s);
                        }
                    }
                }
            }
        }

        // Check for errors and append error message if any
        let had_errors =
            pipeline_type.invoke(s!("get_HadErrors"), Some(pipe.clone()), None, Invocation::Instance)?;
        let had_errors_bool = unsafe { had_errors.Anonymous.Anonymous.Anonymous.boolVal } != VARIANT_BOOL(0);

        if had_errors_bool {
            let state_info = pipeline_type.invoke(
                s!("get_PipelineStateInfo"),
                Some(pipe.clone()),
                None,
                Invocation::Instance,
            )?;

            let state_ptr = unsafe { &state_info.Anonymous.Anonymous.Anonymous.punkVal };
            if state_ptr.is_some() {
                let state_info_type = env.automation.resolve_type(s!(
                    "System.Management.Automation.Runspaces.PipelineStateInfo"
                ))?;
                let get_reason =
                    state_info_type.method_signature(s!("System.Exception get_Reason()"))?;

                if let Ok(reason) = get_reason.invoke(Some(state_info), None) {
                    let reason_ptr = unsafe { &reason.Anonymous.Anonymous.Anonymous.punkVal };
                    if reason_ptr.is_some() {
                        let exception_type = mscorlib.resolve_type(s!("System.Exception"))?;
                        let get_message =
                            exception_type.method_signature(s!("System.String get_Message()"))?;

                        if let Ok(msg) = get_message.invoke(Some(reason), None) {
                            let s = msg.to_string();
                            if !s.is_empty() {
                                if !result.is_empty() {
                                    result.push('\n');
                                }
                                result.push_str(&s);
                            }
                        }
                    }
                }
            }
        }

        // Get host-captured output (Write-Host, native commands)
        let host_output = env.bootstrap_type.invoke(
            s!("GetHostOutput"),
            None,
            Some(vec![self.instance_id.into()]),
            Invocation::Static,
        )?;
        let host_str = host_output.to_string();

        // Dispose pipeline
        let _ = pipeline_type.invoke(s!("Dispose"), Some(pipe), None, Invocation::Instance);

        // Combine: host output first, then pipeline output
        let mut combined = String::new();
        if !host_str.is_empty() {
            combined.push_str(&host_str);
        }
        if !result.is_empty() {
            if !combined.is_empty() && !combined.ends_with('\n') {
                combined.push('\n');
            }
            combined.push_str(&result);
        }

        Ok(combined)
    }
}

impl Drop for PowerShell {
    fn drop(&mut self) {
        let guard = COMPILED_ENV.lock();
        if let Some(env) = guard.as_ref() {
            let _ = env.bootstrap_type.invoke(
                s!("CloseRunspace"),
                None,
                Some(vec![self.instance_id.into()]),
                Invocation::Static,
            );
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::error::Result;
    use super::PowerShell;

    #[test]
    fn test_powershell() -> Result<()> {
        let pwsh = PowerShell::new()?;
        let output = pwsh.execute("whoami /all")?;
        assert!(
            output.contains("\\")
                || output.contains("User")
                || output.contains("Account")
                || output.contains("Authority"),
            "whoami output does not look valid: {output}"
        );

        Ok(())
    }
}