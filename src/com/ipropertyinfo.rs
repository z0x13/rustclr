use alloc::string::{String, ToString};
use alloc::vec::Vec;
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::Variant::VARIANT;
use windows::core::{BSTR, GUID, HRESULT, IUnknown, Interface};

use const_encrypt::obf;

use crate::error::{ClrError, Result};
use crate::variant::create_safe_args;

/// This struct represents the COM `_PropertyInfo` interface.
#[repr(C)]
#[derive(Clone)]
pub struct _PropertyInfo(windows::core::IUnknown);

impl _PropertyInfo {
    /// Retrieves the value of the property.
    #[inline]
    pub fn value(&self, instance: Option<VARIANT>, args: Option<Vec<VARIANT>>) -> Result<VARIANT> {
        let args_array = args.map(create_safe_args).transpose()?;
        let args_ptr = args_array.as_ref().map_or(null_mut(), |a| a.as_ptr());

        let instance_var = instance.unwrap_or_default();
        self.GetValue(instance_var, args_ptr)
    }

    /// Creates an `_PropertyInfo` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_PropertyInfo> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_PropertyInfo>()
            .map_err(|_| ClrError::CastingError(obf!("_PropertyInfo").to_string()))
    }

    /// Retrieves the string representation of the property (equivalent to `ToString` in .NET).
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

    /// Retrieves a property value.
    #[inline]
    pub fn GetValue(&self, instance: VARIANT, args: *mut SAFEARRAY) -> Result<VARIANT> {
        unsafe {
            let mut result = VARIANT::default();
            let hr = (Interface::vtable(self).GetValue)(
                Interface::as_raw(self),
                instance,
                args,
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError(obf!("GetValue").to_string(), hr.0))
            }
        }
    }
}

unsafe impl Interface for _PropertyInfo {
    type Vtable = _PropertyInfo_Vtbl;
    const IID: GUID = GUID::from_u128(0xF59ED4E4_E68F_3218_BD77_061AA82824BF);
}

impl Deref for _PropertyInfo {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

type BSTR_PTR = *const u16;

#[repr(C)]
pub struct _PropertyInfo_Vtbl {
    pub base__: windows::core::IUnknown_Vtbl,

    // IDispatch methods
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,

    // Methods specific to the COM interface
    get_ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR_PTR) -> HRESULT,
    Equals: *const c_void,
    GetHashCode: *const c_void,
    GetType: *const c_void,
    get_MemberType: *const c_void,
    get_name: *const c_void,
    get_DeclaringType: *const c_void,
    get_ReflectedType: *const c_void,
    GetCustomAttributes: *const c_void,
    GetCustomAttributes_2: *const c_void,
    IsDefined: *const c_void,
    get_PropertyType: *const c_void,
    GetValue: unsafe extern "system" fn(
        this: *mut c_void,
        obj: VARIANT,
        index: *mut SAFEARRAY,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    GetValue_2: *const c_void,
    SetValue: *const c_void,
    SetValue_2: *const c_void,
    GetAccessors: *const c_void,
    GetGetMethod: *const c_void,
    GetSetMethod: *const c_void,
    GetIndexParameters: *const c_void,
    get_Attributes: *const c_void,
    get_CanRead: *const c_void,
    get_CanWrite: *const c_void,
    GetAccessors_2: *const c_void,
    GetGetMethod_2: *const c_void,
    GetSetMethod_2: *const c_void,
    get_IsSpecialName: *const c_void,
}
