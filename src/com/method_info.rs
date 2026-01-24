use alloc::string::{String, ToString};
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::Variant::VARIANT;
use windows::core::{BSTR, GUID, HRESULT, IUnknown, Interface};

use super::_Type;
use crate::error::{ClrError, Result};
use crate::wrappers::SafeArray as SafeArrayWrapper;

/// This struct represents the COM `_MethodInfo` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _MethodInfo(windows::core::IUnknown);

impl _MethodInfo {
    /// Invokes the method represented by this `_MethodInfo` instance.
    #[inline]
    pub fn invoke(
        &self,
        obj: Option<VARIANT>,
        parameters: Option<&SafeArrayWrapper>,
    ) -> Result<VARIANT> {
        let variant_obj = obj.unwrap_or_default();
        let params_ptr = parameters.map_or(null_mut(), |p| p.as_ptr());
        self.Invoke_3(variant_obj, params_ptr)
    }

    /// Creates an `_MethodInfo` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_MethodInfo> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_MethodInfo>()
            .map_err(|_| ClrError::CastingError("_MethodInfo"))
    }

    /// Retrieves the string representation of the method (equivalent to `ToString` in .NET).
    #[inline]
    pub fn ToString(&self) -> Result<String> {
        unsafe {
            let mut result: *const u16 = core::ptr::null();
            let hr = (Interface::vtable(self).get_ToString)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                let bstr = BSTR::from_raw(result);
                Ok(bstr.to_string())
            } else {
                Err(ClrError::ApiError("ToString", hr.0))
            }
        }
    }

    /// Retrieves the name of the method.
    #[inline]
    pub fn get_name(&self) -> Result<String> {
        unsafe {
            let mut result: *const u16 = core::ptr::null();
            let hr = (Interface::vtable(self).get_name)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                let bstr = BSTR::from_raw(result);
                Ok(bstr.to_string())
            } else {
                Err(ClrError::ApiError("get_name", hr.0))
            }
        }
    }

    /// Internal invocation method for the method, used by `invoke`.
    #[inline]
    pub fn Invoke_3(&self, obj: VARIANT, parameters: *mut SAFEARRAY) -> Result<VARIANT> {
        unsafe {
            let mut result = VARIANT::default();
            let hr = (Interface::vtable(self).Invoke_3)(
                Interface::as_raw(self),
                obj,
                parameters,
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("Invoke_3", hr.0))
            }
        }
    }

    /// Retrieves the parameters of the method as a `SAFEARRAY`.
    #[inline]
    pub fn GetParameters(&self) -> Result<*mut SAFEARRAY> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).GetParameters)(Interface::as_raw(self), &mut result)
        };
        if hr.is_ok() {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetParameters", hr.0))
        }
    }

    /// Calls the `GetHashCode` method from the vtable of the `_MethodInfo` interface.
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

    /// Calls the `GetBaseDefinition` method from the vtable of the `_MethodInfo` interface.
    #[inline]
    pub fn GetBaseDefinition(&self) -> Result<_MethodInfo> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).GetBaseDefinition)(Interface::as_raw(self), &mut result)
        };
        if hr.is_ok() {
            _MethodInfo::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("GetBaseDefinition", hr.0))
        }
    }

    /// Retrieves the main type associated with the method.
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
}

unsafe impl Interface for _MethodInfo {
    type Vtable = _MethodInfo_Vtbl;
    const IID: GUID = GUID::from_u128(0xffcc1b5d_ecb8_38dd_9b01_3dc8abc2aa5f);
}

impl Deref for _MethodInfo {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

type BSTR_PTR = *const u16;

#[repr(C)]
pub struct _MethodInfo_Vtbl {
    pub base__: windows::core::IUnknown_Vtbl,

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
    get_MemberType: *const c_void,
    get_name: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR_PTR) -> HRESULT,
    get_DeclaringType: *const c_void,
    get_ReflectedType: *const c_void,
    GetCustomAttributes: *const c_void,
    GetCustomAttributes_2: *const c_void,
    IsDefined: *const c_void,
    GetParameters:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut SAFEARRAY) -> HRESULT,
    GetMethodImplementationFlags: *const c_void,
    get_MethodHandle: *const c_void,
    get_Attributes: *const c_void,
    get_CallingConvention: *const c_void,
    Invoke_2: *const c_void,
    get_IsPublic: *const c_void,
    get_IsPrivate: *const c_void,
    get_IsFamily: *const c_void,
    get_IsAssembly: *const c_void,
    get_IsFamilyAndAssembly: *const c_void,
    get_IsFamilyOrAssembly: *const c_void,
    get_IsStatic: *const c_void,
    get_IsFinal: *const c_void,
    get_IsVirtual: *const c_void,
    get_IsHideBySig: *const c_void,
    get_IsAbstract: *const c_void,
    get_IsSpecialName: *const c_void,
    get_IsConstructor: *const c_void,
    Invoke_3: unsafe extern "system" fn(
        this: *mut c_void,
        obj: VARIANT,
        parameters: *mut SAFEARRAY,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    get_returnType: *const c_void,
    get_ReturnTypeCustomAttributes: *const c_void,
    GetBaseDefinition:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut _MethodInfo) -> HRESULT,
}
