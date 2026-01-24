use alloc::{string::String, vec::Vec};
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::Ole::{
    SafeArrayDestroy, SafeArrayGetElement, SafeArrayGetLBound, SafeArrayGetUBound,
};
use windows::core::{BSTR, GUID, HRESULT, IUnknown, Interface};

use super::{_Assembly, _Type};
use crate::error::{ClrError, Result};
use crate::variant::create_safe_array_buffer;

/// This struct represents the COM `_AppDomain` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _AppDomain(windows::core::IUnknown);

impl _AppDomain {
    /// Loads an assembly into the current application domain from a byte slice.
    #[inline]
    pub fn load_bytes(&self, buffer: &[u8]) -> Result<_Assembly> {
        let safe_array = create_safe_array_buffer(buffer)?;
        self.Load_3(safe_array.as_ptr())
    }

    /// Loads an assembly by its name in the current application domain.
    #[inline]
    pub fn load_name(&self, name: &str) -> Result<_Assembly> {
        let lib_name = BSTR::from(name);
        self.Load_2(lib_name.as_ptr())
    }

    /// Creates an `_AppDomain` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_AppDomain> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_AppDomain>()
            .map_err(|_| ClrError::CastingError("_AppDomain"))
    }

    /// Searches for an assembly by name within the current AppDomain.
    #[inline]
    pub fn get_assembly(&self, assembly_name: &str) -> Result<_Assembly> {
        let assemblies = self.assemblies()?;
        for (name, assembly) in assemblies {
            if name.contains(assembly_name) {
                return Ok(assembly);
            }
        }

        Err(ClrError::Msg("Assembly Not Found"))
    }

    /// Retrieves all assemblies currently loaded in the AppDomain.
    #[inline]
    pub fn assemblies(&self) -> Result<Vec<(String, _Assembly)>> {
        let sa_assemblies = self.GetAssemblies()?;
        if sa_assemblies.is_null() {
            return Err(ClrError::NullPointerError("GetAssemblies"));
        }

        let mut assemblies = Vec::new();
        unsafe {
            let lbound = SafeArrayGetLBound(sa_assemblies, 1)
                .map_err(|err| ClrError::ApiError("SafeArrayGetLBound", err.code().0))?;
            let ubound = SafeArrayGetUBound(sa_assemblies, 1)
                .map_err(|err| ClrError::ApiError("SafeArrayGetUBound", err.code().0))?;

            for i in lbound..=ubound {
                let mut p_assembly = null_mut::<_Assembly>();
                if let Err(err) =
                    SafeArrayGetElement(sa_assemblies, &i, &mut p_assembly as *mut _ as *mut _)
                {
                    let _ = SafeArrayDestroy(sa_assemblies);
                    return Err(ClrError::ApiError("SafeArrayGetElement", err.code().0));
                }
                if p_assembly.is_null() {
                    let _ = SafeArrayDestroy(sa_assemblies);
                    return Err(ClrError::NullPointerError("SafeArrayGetElement"));
                }

                let _assembly = _Assembly::from_raw(p_assembly as *mut c_void)?;
                let assembly_name = _assembly.ToString()?;
                assemblies.push((assembly_name, _assembly));
            }

            let _ = SafeArrayDestroy(sa_assemblies);
        }

        Ok(assemblies)
    }

    /// Calls the `Load_3` method from the vtable of the `_AppDomain` interface.
    #[inline]
    pub fn Load_3(&self, rawAssembly: *mut SAFEARRAY) -> Result<_Assembly> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).Load_3)(Interface::as_raw(self), rawAssembly, &mut result)
        };
        if hr.is_ok() {
            _Assembly::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("Load_3", hr.0))
        }
    }

    /// Calls the `Load_2` method from the vtable of the `_AppDomain` interface.
    #[inline]
    pub fn Load_2(&self, assemblyString: *const u16) -> Result<_Assembly> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).Load_2)(Interface::as_raw(self), assemblyString, &mut result)
        };
        if hr.is_ok() {
            _Assembly::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("Load_2", hr.0))
        }
    }

    /// Calls the `GetHashCode` method from the vtable of the `_AppDomain` interface.
    #[inline]
    pub fn GetHashCode(&self) -> Result<u32> {
        let mut result = 0;
        let hr =
            unsafe { (Interface::vtable(self).GetHashCode)(Interface::as_raw(self), &mut result) };
        if hr.is_ok() {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetHashCode", hr.0))
        }
    }

    /// Retrieves the primary type associated with the current app domain.
    #[inline]
    pub fn GetType(&self) -> Result<_Type> {
        let mut result = null_mut();
        let hr = unsafe { (Interface::vtable(self).GetType)(Interface::as_raw(self), &mut result) };
        if hr.is_ok() {
            _Type::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("GetType", hr.0))
        }
    }

    /// Retrieves the assemblies currently loaded into the current AppDomain.
    #[inline]
    pub fn GetAssemblies(&self) -> Result<*mut SAFEARRAY> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).GetAssemblies)(Interface::as_raw(self), &mut result)
        };
        if hr.is_ok() {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetAssemblies", hr.0))
        }
    }
}

unsafe impl Interface for _AppDomain {
    type Vtable = _AppDomainVtbl;
    const IID: GUID = GUID::from_u128(0x05F696DC_2B29_3663_AD8B_C4389CF2A713);
}

impl Deref for _AppDomain {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

type BSTR_PTR = *const u16;

#[repr(C)]
pub struct _AppDomainVtbl {
    pub base__: windows::core::IUnknown_Vtbl,
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,
    get_ToString: *const c_void,
    Equals: *const c_void,
    GetHashCode: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut u32) -> HRESULT,
    GetType: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut _Type) -> HRESULT,
    InitializeLifetimeService: *const c_void,
    GetLifetimeService: *const c_void,
    get_Evidence: *const c_void,
    add_DomainUnload: *const c_void,
    remove_DomainUnload: *const c_void,
    add_AssemblyLoad: *const c_void,
    remove_AssemblyLoad: *const c_void,
    add_ProcessExit: *const c_void,
    remove_ProcessExit: *const c_void,
    add_TypeResolve: *const c_void,
    remove_TypeResolve: *const c_void,
    add_ResourceResolve: *const c_void,
    remove_ResourceResolve: *const c_void,
    add_AssemblyResolve: *const c_void,
    remove_AssemblyResolve: *const c_void,
    add_UnhandledException: *const c_void,
    remove_UnhandledException: *const c_void,
    DefineDynamicAssembly: *const c_void,
    DefineDynamicAssembly_2: *const c_void,
    DefineDynamicAssembly_3: *const c_void,
    DefineDynamicAssembly_4: *const c_void,
    DefineDynamicAssembly_5: *const c_void,
    DefineDynamicAssembly_6: *const c_void,
    DefineDynamicAssembly_7: *const c_void,
    DefineDynamicAssembly_8: *const c_void,
    DefineDynamicAssembly_9: *const c_void,
    CreateInstance: *const c_void,
    CreateInstanceFrom: *const c_void,
    CreateInstance_2: *const c_void,
    CreateInstanceFrom_2: *const c_void,
    CreateInstance_3: *const c_void,
    CreateInstanceFrom_3: *const c_void,
    Load: *const c_void,
    Load_2: unsafe extern "system" fn(
        this: *mut c_void,
        assemblyString: BSTR_PTR,
        pRetVal: *mut *mut _Assembly,
    ) -> HRESULT,
    Load_3: unsafe extern "system" fn(
        this: *mut c_void,
        rawAssembly: *mut SAFEARRAY,
        pRetVal: *mut *mut _Assembly,
    ) -> HRESULT,
    Load_4: *const c_void,
    Load_5: *const c_void,
    Load_6: *const c_void,
    Load_7: *const c_void,
    ExecuteAssembly: *const c_void,
    ExecuteAssembly_2: *const c_void,
    ExecuteAssembly_3: *const c_void,
    get_FriendlyName: *const c_void,
    get_BaseDirectory: *const c_void,
    get_RelativeSearchPath: *const c_void,
    get_ShadowCopyFiles: *const c_void,
    GetAssemblies:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut SAFEARRAY) -> HRESULT,
    AppendPrivatePath: *const c_void,
    ClearPrivatePath: *const c_void,
    SetShadowCopyPath: *const c_void,
    ClearShadowCopyPath: *const c_void,
    SetCachePath: *const c_void,
    SetData: *const c_void,
    GetData: *const c_void,
    SetAppDomainPolicy: *const c_void,
    SetThreadPrincipal: *const c_void,
    SetPrincipalPolicy: *const c_void,
    DoCallBack: *const c_void,
    get_DynamicDirectory: *const c_void,
}
