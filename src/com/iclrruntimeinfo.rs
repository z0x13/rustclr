use alloc::ffi::CString;
use alloc::string::ToString;
use core::{ffi::c_void, ops::Deref};

use windows::Win32::Foundation::{HANDLE, HMODULE};
use windows::core::{BOOL, GUID, HRESULT, Interface, PCSTR, PCWSTR, PWSTR};

use const_encrypt::obf;

use crate::error::{ClrError, Result};

/// This struct represents the COM `ICLRRuntimeInfo` interface.
#[repr(C)]
#[derive(Clone)]
pub struct ICLRRuntimeInfo(windows::core::IUnknown);

impl ICLRRuntimeInfo {
    /// Checks if the CLR runtime has been started.
    #[inline]
    pub fn is_started(&self) -> bool {
        let mut started = BOOL::default();
        let mut startup_flags = 0;
        self.IsStarted(&mut started, &mut startup_flags).is_ok() && started.as_bool()
    }

    /// Checks if the .NET runtime is loadable in the current process.
    #[inline]
    pub fn IsLoadable(&self) -> Result<BOOL> {
        unsafe {
            let mut result = BOOL::default();
            let hr = (Interface::vtable(self).IsLoadable)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError(obf!("IsLoadable").to_string(), hr.0))
            }
        }
    }

    /// Retrieves a COM interface by its class identifier.
    #[inline]
    pub fn GetInterface<T>(&self, rclsid: *const GUID) -> Result<T>
    where
        T: Interface,
    {
        unsafe {
            let mut result = core::ptr::null_mut();
            let hr = (Interface::vtable(self).GetInterface)(
                Interface::as_raw(self),
                rclsid,
                &T::IID,
                &mut result,
            );
            if hr.is_ok() {
                Ok(core::mem::transmute_copy(&result))
            } else {
                Err(ClrError::ApiError(obf!("GetInterface").to_string(), hr.0))
            }
        }
    }

    /// Retrieves the version string of the CLR runtime.
    #[inline]
    pub fn GetVersionString(&self, pwzbuffer: PWSTR, pcchbuffer: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetVersionString)(
                Interface::as_raw(self),
                pwzbuffer,
                pcchbuffer,
            );
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError(
                    obf!("GetVersionString").to_string(),
                    hr.0,
                ))
            }
        }
    }

    /// Retrieves the directory where the CLR runtime is installed.
    #[inline]
    pub fn GetRuntimeDirectory(&self, pwzbuffer: PWSTR, pcchbuffer: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetRuntimeDirectory)(
                Interface::as_raw(self),
                pwzbuffer,
                pcchbuffer,
            );
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError(
                    obf!("GetRuntimeDirectory").to_string(),
                    hr.0,
                ))
            }
        }
    }

    /// Checks if the runtime is loaded in a specified process.
    #[inline]
    pub fn IsLoaded(&self, hndProcess: HANDLE) -> Result<BOOL> {
        unsafe {
            let mut pbLoaded = BOOL::default();
            let hr = (Interface::vtable(self).IsLoaded)(
                Interface::as_raw(self),
                hndProcess,
                &mut pbLoaded,
            );
            if hr.is_ok() {
                Ok(pbLoaded)
            } else {
                Err(ClrError::ApiError(obf!("IsLoaded").to_string(), hr.0))
            }
        }
    }

    /// Loads an error string by its resource ID.
    #[inline]
    pub fn LoadErrorString(
        &self,
        iResourceID: u32,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
        iLocaleID: i32,
    ) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).LoadErrorString)(
                Interface::as_raw(self),
                iResourceID,
                pwzBuffer,
                pcchBuffer,
                iLocaleID,
            );
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError(
                    obf!("LoadErrorString").to_string(),
                    hr.0,
                ))
            }
        }
    }

    /// Loads a DLL by name.
    #[inline]
    pub fn LoadLibraryA(&self, pwzDllName: PCWSTR) -> Result<HMODULE> {
        unsafe {
            let mut result = HMODULE::default();
            let hr = (Interface::vtable(self).LoadLibraryA)(
                Interface::as_raw(self),
                pwzDllName,
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError(obf!("LoadLibraryA").to_string(), hr.0))
            }
        }
    }

    /// Retrieves the address of a procedure in a loaded DLL.
    #[inline]
    pub fn GetProcAddress(&self, pszProcName: &str) -> Result<*mut c_void> {
        unsafe {
            let mut result = core::ptr::null_mut();
            let cstr = CString::new(pszProcName)
                .map_err(|_| ClrError::Msg(obf!("invalid String").to_string()))?;
            let hr = (Interface::vtable(self).GetProcAddress)(
                Interface::as_raw(self),
                PCSTR(cstr.as_ptr().cast()),
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError(obf!("GetProcAddress").to_string(), hr.0))
            }
        }
    }

    /// Sets the default startup flags for the runtime.
    #[inline]
    pub fn SetDefaultStartupFlags(
        &self,
        dwstartupflags: u32,
        pwzhostconfigfile: PCWSTR,
    ) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).SetDefaultStartupFlags)(
                Interface::as_raw(self),
                dwstartupflags,
                pwzhostconfigfile,
            );
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError(
                    obf!("SetDefaultStartupFlags").to_string(),
                    hr.0,
                ))
            }
        }
    }

    /// Retrieves the default startup flags for the runtime.
    #[inline]
    pub fn GetDefaultStartupFlags(
        &self,
        pdwstartupflags: *mut u32,
        pwzhostconfigfile: PWSTR,
        pcchhostconfigfile: *mut u32,
    ) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetDefaultStartupFlags)(
                Interface::as_raw(self),
                pdwstartupflags,
                pwzhostconfigfile,
                pcchhostconfigfile,
            );
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError(
                    obf!("GetDefaultStartupFlags").to_string(),
                    hr.0,
                ))
            }
        }
    }

    /// Configures the runtime to behave as a legacy v2 runtime.
    #[inline]
    pub fn BindAsLegacyV2Runtime(&self) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).BindAsLegacyV2Runtime)(Interface::as_raw(self));
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError(
                    obf!("BindAsLegacyV2Runtime").to_string(),
                    hr.0,
                ))
            }
        }
    }

    /// Checks if the runtime has started and retrieves startup flags.
    #[inline]
    pub fn IsStarted(&self, pbstarted: *mut BOOL, pdwstartupflags: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).IsStarted)(
                Interface::as_raw(self),
                pbstarted,
                pdwstartupflags,
            );
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError(obf!("IsStarted").to_string(), hr.0))
            }
        }
    }
}

unsafe impl Interface for ICLRRuntimeInfo {
    type Vtable = ICLRRuntimeInfo_Vtbl;
    const IID: GUID = GUID::from_u128(0xbd39d1d2_ba2f_486a_89b0_b4b0cb466891);
}

impl Deref for ICLRRuntimeInfo {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

#[repr(C)]
pub struct ICLRRuntimeInfo_Vtbl {
    pub base__: windows::core::IUnknown_Vtbl,
    pub GetVersionString: unsafe extern "system" fn(
        this: *mut c_void,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
    ) -> HRESULT,
    pub GetRuntimeDirectory: unsafe extern "system" fn(
        this: *mut c_void,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
    ) -> HRESULT,
    pub IsLoaded: unsafe extern "system" fn(
        this: *mut c_void,
        hndProcess: HANDLE,
        pbLoaded: *mut BOOL,
    ) -> HRESULT,
    pub LoadErrorString: unsafe extern "system" fn(
        this: *mut c_void,
        iResourceID: u32,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
        iLocaleID: i32,
    ) -> HRESULT,
    pub LoadLibraryA: unsafe extern "system" fn(
        this: *mut c_void,
        pwzDllName: PCWSTR,
        phndModule: *mut HMODULE,
    ) -> HRESULT,
    pub GetProcAddress: unsafe extern "system" fn(
        this: *mut c_void,
        pszProcName: PCSTR,
        ppProc: *mut *mut c_void,
    ) -> HRESULT,
    pub GetInterface: unsafe extern "system" fn(
        this: *mut c_void,
        rclsid: *const GUID,
        riid: *const GUID,
        ppUnk: *mut *mut c_void,
    ) -> HRESULT,
    pub IsLoadable: unsafe extern "system" fn(this: *mut c_void, pbLoadable: *mut BOOL) -> HRESULT,
    pub SetDefaultStartupFlags: unsafe extern "system" fn(
        this: *mut c_void,
        dwStartupFlags: u32,
        pwzHostConfigFile: PCWSTR,
    ) -> HRESULT,
    pub GetDefaultStartupFlags: unsafe extern "system" fn(
        this: *mut c_void,
        dwStartupFlags: *mut u32,
        pwzHostConfigFile: PWSTR,
        pcchHostConfigFile: *mut u32,
    ) -> HRESULT,
    pub BindAsLegacyV2Runtime: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub IsStarted: unsafe extern "system" fn(
        this: *mut c_void,
        pbStarted: *mut BOOL,
        pdwStartupFlags: *mut u32,
    ) -> HRESULT,
}
