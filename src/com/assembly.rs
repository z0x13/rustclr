use alloc::string::{String, ToString};
use alloc::vec::Vec;
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use const_encrypt::obf;
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::Ole::{
    SafeArrayDestroy, SafeArrayGetElement, SafeArrayGetLBound, SafeArrayGetUBound,
};
use windows::Win32::System::Variant::VARIANT;
use windows::core::{BSTR, GUID, HRESULT, IUnknown, Interface};

use super::{_MethodInfo, _Type};
use crate::error::{ClrError, Result};
use crate::wrappers::SafeArray as SafeArrayWrapper;

/// This struct represents the COM `_Assembly` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _Assembly(windows::core::IUnknown);

impl _Assembly {
    /// Resolves a type by name within the assembly.
    #[inline]
    pub fn resolve_type(&self, name: &str) -> Result<_Type> {
        let type_name = BSTR::from(name);
        self.GetType_2(type_name.as_ptr())
    }

    /// Executes the entry point of the assembly.
    #[inline]
    pub fn run(&self, args: &SafeArrayWrapper) -> Result<VARIANT> {
        let entrypoint = self.get_EntryPoint()?;
        let str = entrypoint.ToString()?;
        match str.as_str() {
            str if str.ends_with(&*obf!("Main()").as_str()) => entrypoint.invoke(None, None),
            str if str.ends_with(&*obf!("Main(System.String[])").as_str()) => {
                entrypoint.invoke(None, Some(args))
            }
            _ => Err(ClrError::MethodNotFound),
        }
    }

    /// Creates an instance of a type within the assembly.
    #[inline]
    pub fn create_instance(&self, name: &str) -> Result<VARIANT> {
        let type_name = BSTR::from(name);
        self.CreateInstance(type_name.as_ptr())
    }

    /// Retrieves all types within the assembly.
    #[inline]
    pub fn types(&self) -> Result<Vec<String>> {
        let sa_types = self.GetTypes()?;
        if sa_types.is_null() {
            return Err(ClrError::NullPointerError(obf!("GetTypes").to_string()));
        }

        let mut types = Vec::new();
        unsafe {
            let lbound = SafeArrayGetLBound(sa_types, 1).map_err(|err| {
                ClrError::ApiError(obf!("SafeArrayGetLBound").to_string(), err.code().0)
            })?;
            let ubound = SafeArrayGetUBound(sa_types, 1).map_err(|err| {
                ClrError::ApiError(obf!("SafeArrayGetUBound").to_string(), err.code().0)
            })?;

            for i in lbound..=ubound {
                let mut p_type = null_mut::<_Type>();
                if let Err(err) = SafeArrayGetElement(sa_types, &i, &mut p_type as *mut _ as *mut _)
                {
                    let _ = SafeArrayDestroy(sa_types);
                    return Err(ClrError::ApiError(
                        obf!("SafeArrayGetElement").to_string(),
                        err.code().0,
                    ));
                }
                if p_type.is_null() {
                    let _ = SafeArrayDestroy(sa_types);
                    return Err(ClrError::NullPointerError(
                        obf!("SafeArrayGetElement").to_string(),
                    ));
                }

                let _type = _Type::from_raw(p_type as *mut c_void)?;
                let type_name = _type.ToString()?;
                types.push(type_name);
            }

            let _ = SafeArrayDestroy(sa_types);
        }

        Ok(types)
    }

    /// Creates an `_Assembly` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_Assembly> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_Assembly>()
            .map_err(|_| ClrError::CastingError(obf!("_Assembly").to_string()))
    }

    /// Retrieves the string representation of the assembly.
    #[inline]
    pub fn ToString(&self) -> Result<String> {
        unsafe {
            let mut result: *const u16 = core::ptr::null();
            let hr = (Interface::vtable(self).get_ToString)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                let bstr = BSTR::from_raw(result);
                Ok(bstr.to_string())
            } else {
                Err(ClrError::ApiError(obf!("ToString").to_string(), hr.0))
            }
        }
    }

    /// Calls the `GetHashCode` method from the vtable of the `_Assembly` interface.
    #[inline]
    pub fn GetHashCode(&self) -> Result<u32> {
        let mut result = 0;
        let hr =
            unsafe { (Interface::vtable(self).GetHashCode)(Interface::as_raw(self), &mut result) };
        if hr.is_ok() {
            Ok(result)
        } else {
            Err(ClrError::ApiError(obf!("GetHashCode").to_string(), hr.0))
        }
    }

    /// Retrieves the entry point method of the assembly.
    #[inline]
    pub fn get_EntryPoint(&self) -> Result<_MethodInfo> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).get_EntryPoint)(Interface::as_raw(self), &mut result)
        };
        if hr.is_ok() {
            _MethodInfo::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError(obf!("get_EntryPoint").to_string(), hr.0))
        }
    }

    /// Resolves a specific type by name within the assembly.
    #[inline]
    pub fn GetType_2(&self, name: *const u16) -> Result<_Type> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).GetType_2)(Interface::as_raw(self), name, &mut result)
        };
        if hr.is_ok() {
            _Type::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError(obf!("GetType_2").to_string(), hr.0))
        }
    }

    /// Retrieves all types defined within the assembly as a `SAFEARRAY`.
    #[inline]
    pub fn GetTypes(&self) -> Result<*mut SAFEARRAY> {
        let mut result = null_mut();
        let hr =
            unsafe { (Interface::vtable(self).GetTypes)(Interface::as_raw(self), &mut result) };
        if hr.is_ok() {
            Ok(result)
        } else {
            Err(ClrError::ApiError(obf!("GetTypes").to_string(), hr.0))
        }
    }

    /// Creates an instance of a type using its name.
    #[inline]
    pub fn CreateInstance(&self, typeName: *const u16) -> Result<VARIANT> {
        let mut result = VARIANT::default();
        let hr = unsafe {
            (Interface::vtable(self).CreateInstance)(Interface::as_raw(self), typeName, &mut result)
        };
        if hr.is_ok() {
            Ok(result)
        } else {
            Err(ClrError::ApiError(obf!("CreateInstance").to_string(), hr.0))
        }
    }

    /// Retrieves the main type associated with the assembly.
    #[inline]
    pub fn GetType(&self) -> Result<_Type> {
        let mut result = null_mut();
        let hr = unsafe { (Interface::vtable(self).GetType)(Interface::as_raw(self), &mut result) };
        if hr.is_ok() {
            _Type::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError(obf!("GetType").to_string(), hr.0))
        }
    }

    /// Retrieves the assembly's codebase as a URI.
    #[inline]
    pub fn get_CodeBase(&self) -> Result<String> {
        unsafe {
            let mut result: *const u16 = core::ptr::null();
            let hr = (Interface::vtable(self).get_CodeBase)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                let bstr = BSTR::from_raw(result);
                Ok(bstr.to_string())
            } else {
                Err(ClrError::ApiError(obf!("get_CodeBase").to_string(), hr.0))
            }
        }
    }

    /// Retrieves the escaped codebase of the assembly as a URI.
    #[inline]
    pub fn get_EscapedCodeBase(&self) -> Result<String> {
        unsafe {
            let mut result: *const u16 = core::ptr::null();
            let hr =
                (Interface::vtable(self).get_EscapedCodeBase)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                let bstr = BSTR::from_raw(result);
                Ok(bstr.to_string())
            } else {
                Err(ClrError::ApiError(
                    obf!("get_EscapedCodeBase").to_string(),
                    hr.0,
                ))
            }
        }
    }

    /// Retrieves the name of the assembly.
    #[inline]
    pub fn GetName(&self) -> Result<*mut c_void> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetName)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError(obf!("GetName").to_string(), hr.0))
            }
        }
    }

    /// Retrieves the name of the assembly, with an option to copy the name.
    #[inline]
    pub fn GetName_2(&self, copiedName: VARIANT_BOOL) -> Result<*mut c_void> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetName_2)(
                Interface::as_raw(self),
                copiedName,
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError(obf!("GetName_2").to_string(), hr.0))
            }
        }
    }

    /// Retrieves the full name of the assembly.
    #[inline]
    pub fn get_FullName(&self) -> Result<String> {
        unsafe {
            let mut result: *const u16 = core::ptr::null();
            let hr = (Interface::vtable(self).get_FullName)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                let bstr = BSTR::from_raw(result);
                Ok(bstr.to_string())
            } else {
                Err(ClrError::ApiError(obf!("get_FullName").to_string(), hr.0))
            }
        }
    }

    /// Retrieves the file location of the assembly.
    #[inline]
    pub fn get_Location(&self) -> Result<String> {
        unsafe {
            let mut result: *const u16 = core::ptr::null();
            let hr = (Interface::vtable(self).get_Location)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                let bstr = BSTR::from_raw(result);
                Ok(bstr.to_string())
            } else {
                Err(ClrError::ApiError(obf!("get_Location").to_string(), hr.0))
            }
        }
    }
}

unsafe impl Interface for _Assembly {
    type Vtable = _Assembly_Vtbl;
    const IID: GUID = GUID::from_u128(0x17156360_2f1a_384a_bc52_fde93c215c5b);
}

impl Deref for _Assembly {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

type BSTR_PTR = *const u16;

#[repr(C)]
pub struct _Assembly_Vtbl {
    base__: windows::core::IUnknown_Vtbl,

    // IDispatch methods
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,

    // Methods specific to the COM interface
    get_ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR_PTR) -> HRESULT,
    Equals: *const c_void,
    GetHashCode: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut u32) -> HRESULT,
    GetType: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut _Type) -> HRESULT,
    get_CodeBase: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR_PTR) -> HRESULT,
    get_EscapedCodeBase:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR_PTR) -> HRESULT,
    GetName: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut c_void) -> HRESULT,
    GetName_2: unsafe extern "system" fn(
        this: *mut c_void,
        copiedName: VARIANT_BOOL,
        pRetVal: *mut *mut c_void,
    ) -> HRESULT,
    get_FullName: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR_PTR) -> HRESULT,
    get_EntryPoint:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut _MethodInfo) -> HRESULT,
    GetType_2: unsafe extern "system" fn(
        this: *mut c_void,
        name: BSTR_PTR,
        pRetVal: *mut *mut _Type,
    ) -> HRESULT,
    GetType_3: *const c_void,
    GetExportedTypes: *const c_void,
    GetTypes: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut SAFEARRAY) -> HRESULT,
    GetManifestResourceStream: *const c_void,
    GetManifestResourceStream_2: *const c_void,
    GetFile: *const c_void,
    GetFiles: *const c_void,
    GetFiles_2: *const c_void,
    GetManifestResourceNames: *const c_void,
    GetManifestResourceInfo: *const c_void,
    get_Location: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR_PTR) -> HRESULT,
    get_Evidence: *const c_void,
    GetCustomAttributes: *const c_void,
    GetCustomAttributes_2: *const c_void,
    IsDefined: *const c_void,
    GetObjectData: *const c_void,
    add_ModuleResolve: *const c_void,
    remove_ModuleResolve: *const c_void,
    GetType_4: *const c_void,
    GetSatelliteAssembly: *const c_void,
    GetSatelliteAssembly_2: *const c_void,
    LoadModule: *const c_void,
    LoadModule_2: *const c_void,
    CreateInstance: unsafe extern "system" fn(
        this: *mut c_void,
        typeName: BSTR_PTR,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    CreateInstance_2: *const c_void,
    CreateInstance_3: *const c_void,
    GetLoadedModules: *const c_void,
    GetLoadedModules_2: *const c_void,
    GetModules: *const c_void,
    GetModules_2: *const c_void,
    GetModule: *const c_void,
    GetReferencedAssemblies: *const c_void,
    get_GlobalAssemblyCache: *const c_void,
}
