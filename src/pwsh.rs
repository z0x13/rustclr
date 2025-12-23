use alloc::{format, string::String, vec};
use obfstr::obfstr as s;

use crate::error::Result;
use crate::com::_Assembly;
use crate::string::ComString;
use crate::{Invocation, RustClrEnv};
use crate::variant::{Variant, create_safe_args};

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
    /// The loaded .NET automation assembly.
    automation: _Assembly,

    /// CLR environment used to host the .NET runtime.
    #[allow(dead_code)]
    clr: RustClrEnv,
}

impl PowerShell {
    /// Creates a new `PowerShell`.
    pub fn new() -> Result<Self> {
        // Initialize .NET runtime (v4.0).
        let clr = RustClrEnv::new(None)?;

        // Load `mscorlib` and resolve `System.Reflection.Assembly`.
        let mscorlib = clr.app_domain.get_assembly(s!("mscorlib"))?;
        let reflection_assembly = mscorlib.resolve_type(s!("System.Reflection.Assembly"))?;

        // Resolve and invoke `LoadWithPartialName` method.
        let load_partial_name = reflection_assembly.method_signature(s!(
            "System.Reflection.Assembly LoadWithPartialName(System.String)"
        ))?;
        let param = create_safe_args(vec![s!("System.Management.Automation").to_variant()])?;
        let result = load_partial_name.invoke(None, Some(param))?;

        // Convert result to `_Assembly`.
        let automation =
            _Assembly::from_raw(unsafe { result.Anonymous.Anonymous.Anonymous.byref })?;

        Ok(Self { automation, clr })
    }

    /// Executes a PowerShell command and returns its output as a string.
    pub fn execute(&self, command: &str) -> Result<String> {
        // Invoke `CreateRunspace` method.
        let runspace_factory = self.automation.resolve_type(
            s!("System.Management.Automation.Runspaces.RunspaceFactory")
        )?;
        let create_runspace = runspace_factory.method_signature(s!(
            "System.Management.Automation.Runspaces.Runspace CreateRunspace()"
        ))?;
        let runspace = create_runspace.invoke(None, None)?;

        // Invoke `CreatePipeline` method.
        let assembly_runspace = self.automation.resolve_type(
            s!("System.Management.Automation.Runspaces.Runspace"),
        )?;
        assembly_runspace.invoke(
            s!("Open"),
            Some(runspace),
            None,
            Invocation::Instance,
        )?;
        
        let create_pipeline = assembly_runspace.method_signature(s!(
            "System.Management.Automation.Runspaces.Pipeline CreatePipeline()"
        ))?;
        let pipe = create_pipeline.invoke(Some(runspace), None)?;

        // Invoke `get_Commands` method.
        let pipeline = self.automation.resolve_type(
            s!("System.Management.Automation.Runspaces.Pipeline"),
        )?;
        let get_command = pipeline.invoke(
            s!("get_Commands"),
            Some(pipe),
            None,
            Invocation::Instance,
        )?;

        // Invoke `AddScript` method.
        let cmd = vec![format!("{} | {}", command, s!("Out-String")).to_variant()];
        let command_collection = self.automation.resolve_type(
            s!("System.Management.Automation.Runspaces.CommandCollection"),
        )?;
        let add_script = command_collection.method_signature(s!(
            "Void AddScript(System.String)"
        ))?;
        
        add_script.invoke(Some(get_command), Some(create_safe_args(cmd)?))?;

        // Invoke `InvokeAsync` method.
        pipeline.invoke(
            s!("InvokeAsync"),
            Some(pipe),
            None,
            Invocation::Instance,
        )?;

        // Invoke `get_Output` method.
        let output_reader = pipeline.invoke(
            s!("get_Output"),
            Some(pipe),
            None,
            Invocation::Instance,
        )?;

        let mscorlib = self.clr.app_domain.get_assembly(s!("mscorlib"))?;
        let object_type = mscorlib.resolve_type(s!("System.Object"))?;
        let to_string = object_type.method_signature(s!("System.String ToString()"))?;

        let ps_reader_type = self.automation.resolve_type(s!(
            "System.Management.Automation.Runspaces.PipelineReader`1[System.Management.Automation.PSObject]"
        ))?;
        let read_to_end = ps_reader_type.method_signature(s!(
            "System.Collections.ObjectModel.Collection`1[System.Management.Automation.PSObject] ReadToEnd()"
        ))?;
        let output_collection = read_to_end.invoke(Some(output_reader), None)?;

        let mut result = String::new();

        let output_ptr = unsafe { output_collection.Anonymous.Anonymous.Anonymous.punkVal };
        if !output_ptr.is_null() {
            let icollection = mscorlib.resolve_type(s!("System.Collections.ICollection"))?;
            let get_count = icollection.method_signature(s!("Int32 get_Count()"))?;
            let count_variant = get_count.invoke(Some(output_collection), None)?;
            let count = unsafe { count_variant.Anonymous.Anonymous.Anonymous.lVal };

            if count > 0 {
                let ilist = mscorlib.resolve_type(s!("System.Collections.IList"))?;
                let get_item = ilist.method_signature(s!("System.Object get_Item(Int32)"))?;

                for i in 0..count {
                    let idx_args = create_safe_args(vec![i.to_variant()])?;
                    let item = get_item.invoke(Some(output_collection), Some(idx_args))?;
                    let item_ptr = unsafe { item.Anonymous.Anonymous.Anonymous.punkVal };
                    if !item_ptr.is_null() {
                        let item_str = to_string.invoke(Some(item), None)?;
                        let s = unsafe { item_str.Anonymous.Anonymous.Anonymous.bstrVal.to_string() };
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

        let had_errors = pipeline.invoke(
            s!("get_HadErrors"),
            Some(pipe),
            None,
            Invocation::Instance,
        )?;
        let had_errors_bool = unsafe { had_errors.Anonymous.Anonymous.Anonymous.boolVal } != 0;

        if had_errors_bool {
            let state_info = pipeline.invoke(
                s!("get_PipelineStateInfo"),
                Some(pipe),
                None,
                Invocation::Instance,
            )?;

            let state_ptr = unsafe { state_info.Anonymous.Anonymous.Anonymous.punkVal };
            if !state_ptr.is_null() {
                let state_info_type = self.automation.resolve_type(s!(
                    "System.Management.Automation.Runspaces.PipelineStateInfo"
                ))?;
                let get_reason = state_info_type.method_signature(s!(
                    "System.Exception get_Reason()"
                ))?;

                if let Ok(reason) = get_reason.invoke(Some(state_info), None) {
                    let reason_ptr = unsafe { reason.Anonymous.Anonymous.Anonymous.punkVal };
                    if !reason_ptr.is_null() {
                        let exception_type = mscorlib.resolve_type(s!("System.Exception"))?;
                        let get_message = exception_type.method_signature(s!(
                            "System.String get_Message()"
                        ))?;

                        if let Ok(msg) = get_message.invoke(Some(reason), None) {
                            let s = unsafe { msg.Anonymous.Anonymous.Anonymous.bstrVal.to_string() };
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

        assembly_runspace.invoke(
            s!("Close"),
            Some(runspace),
            None,
            Invocation::Instance,
        )?;

        Ok(result)
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