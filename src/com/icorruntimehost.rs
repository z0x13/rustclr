use alloc::vec::Vec;
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows::core::{GUID, HRESULT, IUnknown, Interface, PCWSTR};
use windows::Win32::Foundation::{HANDLE, HMODULE};

use super::_AppDomain;
use crate::error::{ClrError, Result};

/// This struct represents the COM `ICorRuntimeHost` interface.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct ICorRuntimeHost(windows::core::IUnknown);

impl ICorRuntimeHost {
    /// Creates a new .NET AppDomain with the specified name.
    #[inline]
    pub fn create_domain(&self, name: &str) -> Result<_AppDomain> {
        let name = name.encode_utf16().chain(Some(0)).collect::<Vec<u16>>();
        let domain_name = PCWSTR(name.as_ptr());
        self.CreateDomain(domain_name, null_mut())
    }

    /// Starts the .NET runtime host.
    #[inline]
    pub fn Start(&self) -> HRESULT {
        unsafe { (Interface::vtable(self).Start)(Interface::as_raw(self)) }
    }

    /// Stops the .NET runtime host.
    #[inline]
    pub fn Stop(&self) -> HRESULT {
        unsafe { (Interface::vtable(self).Stop)(Interface::as_raw(self)) }
    }

    /// Retrieves the default application domain for the runtime host.
    pub fn GetDefaultDomain(&self) -> Result<_AppDomain> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetDefaultDomain)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                _AppDomain::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("GetDefaultDomain", hr.0))
            }
        }
    }

    /// Creates a new application domain with the specified name and identity.
    #[inline]
    pub fn CreateDomain(
        &self,
        pwzFriendlyName: PCWSTR,
        pIdentityArray: *mut IUnknown,
    ) -> Result<_AppDomain> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CreateDomain)(
                Interface::as_raw(self),
                pwzFriendlyName,
                pIdentityArray,
                &mut result,
            );
            if hr.is_ok() {
                _AppDomain::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("CreateDomain", hr.0))
            }
        }
    }

    /// Creates a logical thread state within the runtime host.
    #[inline]
    pub fn CreateLogicalThreadState(&self) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).CreateLogicalThreadState)(Interface::as_raw(self));
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError("CreateLogicalThreadState", hr.0))
            }
        }
    }

    /// Deletes the current logical thread state.
    #[inline]
    pub fn DeleteLogicalThreadState(&self) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).DeleteLogicalThreadState)(Interface::as_raw(self));
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError("DeleteLogicalThreadState", hr.0))
            }
        }
    }

    /// Switches into the logical thread state.
    #[inline]
    pub fn SwitchInLogicalThreadState(&self) -> Result<u32> {
        unsafe {
            let mut result = 0;
            let hr = (Interface::vtable(self).SwitchInLogicalThreadState)(
                Interface::as_raw(self),
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("SwitchInLogicalThreadState", hr.0))
            }
        }
    }

    /// Switches out of the logical thread state.
    #[inline]
    pub fn SwitchOutLogicalThreadState(&self) -> Result<*mut u32> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).SwitchOutLogicalThreadState)(
                Interface::as_raw(self),
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("SwitchOutLogicalThreadState", hr.0))
            }
        }
    }

    /// Retrieves the number of locks held by the current logical thread.
    #[inline]
    pub fn LocksHeldByLogicalThread(&self) -> Result<u32> {
        unsafe {
            let mut result = 0;
            let hr = (Interface::vtable(self).LocksHeldByLogicalThread)(
                Interface::as_raw(self),
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("LocksHeldByLogicalThread", hr.0))
            }
        }
    }

    /// Maps a file handle to an `HMODULE`.
    #[inline]
    pub fn MapFile(&self, h_file: HANDLE) -> Result<HMODULE> {
        unsafe {
            let mut result = HMODULE::default();
            let hr = (Interface::vtable(self).MapFile)(Interface::as_raw(self), h_file, &mut result);
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("MapFile", hr.0))
            }
        }
    }

    /// Retrieves the configuration for the runtime host.
    #[inline]
    pub fn GetConfiguration(&self) -> Result<*mut c_void> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetConfiguration)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetConfiguration", hr.0))
            }
        }
    }

    /// Enumerates application domains managed by the runtime host.
    #[inline]
    pub fn EnumDomains(&self) -> Result<*mut c_void> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).EnumDomains)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("EnumDomains", hr.0))
            }
        }
    }

    /// Retrieves the next application domain in the enumeration.
    #[inline]
    pub fn NextDomain(&self, hEnum: *mut c_void) -> Result<IUnknown> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).NextDomain)(Interface::as_raw(self), hEnum, &mut result);
            if hr.is_ok() {
                Ok(IUnknown::from_raw(result as *mut c_void))
            } else {
                Err(ClrError::ApiError("NextDomain", hr.0))
            }
        }
    }

    /// Closes an application domain enumeration.
    #[inline]
    pub fn CloseEnum(&self, hEnum: *mut c_void) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).CloseEnum)(Interface::as_raw(self), hEnum);
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError("CloseEnum", hr.0))
            }
        }
    }

    /// Creates a new application domain with specified configuration.
    #[inline]
    pub fn CreateDomainEx(
        &self,
        pwzFriendlyName: PCWSTR,
        psSetup: *mut IUnknown,
        pEvidence: *mut IUnknown,
    ) -> Result<_AppDomain> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CreateDomainEx)(
                Interface::as_raw(self),
                pwzFriendlyName,
                psSetup,
                pEvidence,
                &mut result,
            );
            if hr.is_ok() {
                _AppDomain::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("CreateDomainEx", hr.0))
            }
        }
    }

    /// Creates a setup configuration object for application domains.
    #[inline]
    pub fn CreateDomainSetup(&self) -> Result<IUnknown> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CreateDomainSetup)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                Ok(IUnknown::from_raw(result as *mut c_void))
            } else {
                Err(ClrError::ApiError("CreateDomainSetup", hr.0))
            }
        }
    }

    /// Creates an evidence object for application domains.
    #[inline]
    pub fn CreateEvidence(&self) -> Result<IUnknown> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CreateEvidence)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                Ok(IUnknown::from_raw(result as *mut c_void))
            } else {
                Err(ClrError::ApiError("CreateEvidence", hr.0))
            }
        }
    }

    /// Unloads the specified application domain.
    #[inline]
    pub fn UnloadDomain(&self, pAppDomain: *mut IUnknown) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).UnloadDomain)(Interface::as_raw(self), pAppDomain);
            if hr.is_ok() {
                Ok(())
            } else {
                Err(ClrError::ApiError("UnloadDomain", hr.0))
            }
        }
    }

    /// Retrieves the current application domain.
    #[inline]
    pub fn CurrentDomain(&self) -> Result<_AppDomain> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CurrentDomain)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                _AppDomain::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("CurrentDomain", hr.0))
            }
        }
    }
}

unsafe impl Interface for ICorRuntimeHost {
    type Vtable = ICorRuntimeHost_Vtbl;
    const IID: GUID = GUID::from_u128(0xCB2F6722_AB3A_11D2_9C40_00C04FA30A3E);
}

impl Deref for ICorRuntimeHost {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

#[repr(C)]
pub struct ICorRuntimeHost_Vtbl {
    pub base__: windows::core::IUnknown_Vtbl,
    pub CreateLogicalThreadState: unsafe extern "system" fn(*mut c_void) -> HRESULT,
    pub DeleteLogicalThreadState: unsafe extern "system" fn(*mut c_void) -> HRESULT,
    pub SwitchInLogicalThreadState: unsafe extern "system" fn(
        this: *mut c_void,
        pFiberCookie: *mut u32,
    ) -> HRESULT,
    pub SwitchOutLogicalThreadState: unsafe extern "system" fn(
        this: *mut c_void,
        pFiberCookie: *mut *mut u32,
    ) -> HRESULT,
    pub LocksHeldByLogicalThread: unsafe extern "system" fn(
        this: *mut c_void,
        pCount: *mut u32,
    ) -> HRESULT,
    pub MapFile: unsafe extern "system" fn(
        this: *mut c_void,
        hFile: HANDLE,
        hMapAddress: *mut HMODULE,
    ) -> HRESULT,
    pub GetConfiguration: unsafe extern "system" fn(
        this: *mut c_void,
        pConfiguration: *mut *mut c_void,
    ) -> HRESULT,
    pub Start: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub Stop: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub CreateDomain: unsafe extern "system" fn(
        this: *mut c_void,
        pwzFriendlyName: PCWSTR,
        pIdentityArray: *mut IUnknown,
        pAppDomain: *mut *mut IUnknown,
    ) -> HRESULT,
    pub GetDefaultDomain: unsafe extern "system" fn(
        this: *mut c_void,
        pAppDomain: *mut *mut IUnknown,
    ) -> HRESULT,
    pub EnumDomains: unsafe extern "system" fn(
        this: *mut c_void,
        hEnum: *mut *mut c_void,
    ) -> HRESULT,
    pub NextDomain: unsafe extern "system" fn(
        this: *mut c_void,
        hEnum: *mut c_void,
        pAppDomain: *mut *mut IUnknown,
    ) -> HRESULT,
    pub CloseEnum: unsafe extern "system" fn(this: *mut c_void, hEnum: *mut c_void) -> HRESULT,
    pub CreateDomainEx: unsafe extern "system" fn(
        this: *mut c_void,
        pwzFriendlyName: PCWSTR,
        pSetup: *mut IUnknown,
        pEvidence: *mut IUnknown,
        pAppDomain: *mut *mut IUnknown,
    ) -> HRESULT,
    pub CreateDomainSetup: unsafe extern "system" fn(
        this: *mut c_void,
        pAppDomainSetup: *mut *mut IUnknown,
    ) -> HRESULT,
    pub CreateEvidence: unsafe extern "system" fn(
        this: *mut c_void,
        pEvidence: *mut *mut IUnknown,
    ) -> HRESULT,
    pub UnloadDomain: unsafe extern "system" fn(
        this: *mut c_void,
        pAppDomain: *mut IUnknown,
    ) -> HRESULT,
    pub CurrentDomain: unsafe extern "system" fn(
        this: *mut c_void,
        pAppDomain: *mut *mut IUnknown,
    ) -> HRESULT,
}
