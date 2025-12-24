use alloc::{string::{String, ToString}, vec::Vec};
use core::{ffi::c_void, ptr::null_mut, sync::atomic::{AtomicU32, Ordering}};

use obfstr::obfstr as s;
use spin::Mutex;
use windows_core::*;
use windows_sys::Win32::UI::Shell::SHCreateMemStream;

use crate::com::*;

/// Shared state for the current assembly being loaded.
/// Updated before each assembly load, cleared after.
static CURRENT_ASSEMBLY: Mutex<Option<AssemblyData>> = Mutex::new(None);

/// Counter for streams created (for debugging)
static STREAM_COUNT: AtomicU32 = AtomicU32::new(0);

/// Holds the current assembly's buffer and identity.
struct AssemblyData {
    buffer: Vec<u8>,
    identity: String,
}

/// Wrapper for raw pointer that implements Send
struct StreamPtr(*mut c_void);
unsafe impl Send for StreamPtr {}

/// Streams created during ProvideAssembly that need to be released after domain unload
static PENDING_STREAMS: Mutex<Vec<StreamPtr>> = Mutex::new(Vec::new());

/// Sets the current assembly to be provided by the host store.
/// Must be called before loading an assembly.
pub fn set_current_assembly(buffer: &[u8], identity: &str) {
    let mut guard = CURRENT_ASSEMBLY.lock();
    dinvk::println!("[rustclr] set_current_assembly: {} bytes, identity={identity}", buffer.len());
    *guard = Some(AssemblyData {
        buffer: buffer.to_vec(),
        identity: identity.to_string(),
    });
}

/// Clears the current assembly from the host store.
/// Should be called after assembly execution completes.
pub fn clear_current_assembly() {
    let mut guard = CURRENT_ASSEMBLY.lock();
    let was_set = guard.is_some();
    *guard = None;
    drop(guard);

    // Release any pending streams
    let mut streams = PENDING_STREAMS.lock();
    let stream_count = streams.len();
    for StreamPtr(stream_ptr) in streams.drain(..) {
        if !stream_ptr.is_null() {
            dinvk::println!("[rustclr] Releasing stream {stream_ptr:?}");
            unsafe {
                let stream = IUnknown::from_raw(stream_ptr);
                drop(stream); // Release() called on drop
            }
        }
    }

    let total_streams = STREAM_COUNT.load(Ordering::SeqCst);
    dinvk::println!("[rustclr] clear_current_assembly: was_set={was_set}, released={stream_count}, total_created={total_streams}");
}

/// Implements `IHostControl`.
#[implement(IHostControl)]
pub struct RustClrControl {
    /// Host manager responsible for resolving assemblies.
    manager: IHostAssemblyManager,
}

impl RustClrControl {
    /// Creates a new `RustClrControl`.
    pub fn new() -> Self {
        Self {
            manager: RustClrManager::new().into(),
        }
    }
}

impl Default for RustClrControl {
    fn default() -> Self {
        Self::new()
    }
}

impl IHostControl_Impl for RustClrControl_Impl {
    /// Returns `IHostAssemblyManager` when requested.
    fn GetHostManager(&self, riid: *const GUID, ppobject: *mut *mut c_void) -> Result<()> {
        unsafe {
            if *riid == IHostAssemblyManager::IID {
                *ppobject = self.manager.as_raw();
                return Ok(());
            }

            *ppobject = null_mut();
            Err(Error::new(
                HRESULT(0x80004002u32 as i32),
                s!("E_NOINTERFACE"),
            ))
        }
    }

    fn SetAppDomainManager(
        &self,
        _dwappdomainid: u32,
        _punkappdomainmanager: Ref<'_, IUnknown>,
    ) -> Result<()> {
        Ok(())
    }
}

/// Implements `IHostAssemblyManager`.
#[implement(IHostAssemblyManager)]
pub struct RustClrManager {
    /// Store responsible for resolving and serving assemblies.
    store: IHostAssemblyStore,
}

impl RustClrManager {
    /// Creates a new [`RustClrManager`].
    pub fn new() -> Self {
        Self {
            store: RustClrStore::new().into(),
        }
    }
}

impl Default for RustClrManager {
    fn default() -> Self {
        Self::new()
    }
}

impl IHostAssemblyManager_Impl for RustClrManager_Impl {
    fn GetNonHostStoreAssemblies(&self) -> Result<()> {
        Ok(())
    }

    /// Returns the custom assembly store used to resolve in-memory assemblies.
    fn GetAssemblyStore(&self) -> Result<IHostAssemblyStore> {
        Ok(self.store.clone())
    }
}

/// Implements `IHostAssemblyStore`.
/// Reads from the shared CURRENT_ASSEMBLY state.
#[implement(IHostAssemblyStore)]
pub struct RustClrStore;

impl RustClrStore {
    /// Creates a new `RustClrStore`.
    pub fn new() -> Self {
        Self
    }
}

impl Default for RustClrStore {
    fn default() -> Self {
        Self::new()
    }
}

impl IHostAssemblyStore_Impl for RustClrStore_Impl {
    /// Returns the managed assembly image from memory when the identity matches.
    fn ProvideAssembly(
        &self,
        pbindinfo: *const AssemblyBindInfo,
        passemblyid: *mut u64,
        pcontext: *mut u64,
        ppstmassemblyimage: *mut *mut c_void,
        _ppstmpdb: *mut *mut c_void,
    ) -> Result<()> {
        let identity = unsafe { (*pbindinfo).lpPostPolicyIdentity.to_string() }?;
        dinvk::println!("[rustclr] ProvideAssembly called for: {identity}");

        let guard = CURRENT_ASSEMBLY.lock();
        if let Some(data) = guard.as_ref() {
            if data.identity == identity {
                let stream = unsafe {
                    SHCreateMemStream(data.buffer.as_ptr(), data.buffer.len() as u32)
                };
                let count = STREAM_COUNT.fetch_add(1, Ordering::SeqCst) + 1;
                dinvk::println!("[rustclr] ProvideAssembly: created stream #{count} {stream:?} for {} bytes", data.buffer.len());

                // AddRef and track stream for later release (we hold our own reference)
                unsafe {
                    let unk = IUnknown::from_raw(stream);
                    let cloned = unk.clone(); // AddRef
                    core::mem::forget(unk); // Don't release the original
                    PENDING_STREAMS.lock().push(StreamPtr(cloned.into_raw()));
                }

                unsafe { *passemblyid = 800 };
                unsafe { *pcontext = 0 };
                unsafe { *ppstmassemblyimage = stream };
                return Ok(());
            }
        }
        drop(guard);

        dinvk::println!("[rustclr] ProvideAssembly: identity not matched, returning E_FILENOTFOUND");
        Err(Error::new(
            HRESULT(0x80070002u32 as i32),
            s!("assembly not recognized"),
        ))
    }

    /// Always returns `ERROR_FILE_NOT_FOUND` as this implementation does not
    /// support module resolution.
    fn ProvideModule(
        &self,
        _pbindinfo: *const ModuleBindInfo,
        _pdwmoduleid: *mut u32,
        _ppstmmoduleimage: *mut *mut c_void,
        _ppstmpdb: *mut *mut c_void,
    ) -> Result<()> {
        Err(Error::new(
            HRESULT(0x80070002u32 as i32),
            s!("module not recognized"),
        ))
    }
}
