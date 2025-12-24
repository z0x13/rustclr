use alloc::{string::String, vec::Vec};
use core::{
    ffi::c_void,
    ops::Deref,
    ptr::{null, null_mut},
};

use windows_core::{GUID, IUnknown, Interface};
use windows_sys::{
    Win32::System::{Com::SAFEARRAY, Variant::VARIANT},
    core::{BSTR, HRESULT},
};

use crate::wrappers::Bstr;
use crate::variant::create_safe_args;
use crate::error::{ClrError, Result};

/// This struct represents the COM `_PropertyInfo` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _PropertyInfo(windows_core::IUnknown);

impl _PropertyInfo {
    /// Retrieves the value of the property.
    ///
    /// Note: `args` VARIANTs are copied to a SAFEARRAY and cleared.
    /// The caller is responsible for clearing returned VARIANTs and `instance` when done.
    #[inline]
    pub fn value(&self, instance: Option<VARIANT>, args: Option<Vec<VARIANT>>) -> Result<VARIANT> {
        // create_safe_args takes ownership and clears the VARIANTs
        let args_array = args
            .map(create_safe_args)
            .transpose()?;
        let args_ptr = args_array.as_ref().map_or(null_mut(), |a| a.as_ptr());

        let instance_var = instance.unwrap_or(unsafe { core::mem::zeroed::<VARIANT>() });
        self.GetValue(instance_var, args_ptr)
        // args_array drops here, SafeArrayDestroy is called automatically
    }

    /// Creates an `_PropertyInfo` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_PropertyInfo> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_PropertyInfo>()
            .map_err(|_| ClrError::CastingError("_PropertyInfo"))
    }

    /// Retrieves the string representation of the method (equivalent to `ToString` in .NET).
    #[inline]
    pub fn ToString(&self) -> Result<String> {
        unsafe {
            let mut result = null::<u16>();
            let hr = (Interface::vtable(self).get_ToString)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                let bstr = Bstr::from_raw(result);
                Ok(bstr.to_string_lossy())
            } else {
                Err(ClrError::ApiError("ToString", hr))
            }
        }
    }

    /// Retrieves a method by name.
    #[inline]
    pub fn GetValue(&self, instance: VARIANT, args: *mut SAFEARRAY) -> Result<VARIANT> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).GetValue)(
                Interface::as_raw(self),
                instance,
                args,
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetValue", hr))
            }
        }
    }
}

unsafe impl Interface for _PropertyInfo {
    type Vtable = _PropertyInfo_Vtbl;

    /// The interface identifier (IID) for the `_PropertyInfo` COM interface.
    ///
    /// This GUID is used to identify the `_PropertyInfo` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `_PropertyInfo` interface.
    const IID: GUID = GUID::from_u128(0xF59ED4E4_E68F_3218_BD77_061AA82824BF);
}

impl Deref for _PropertyInfo {
    type Target = windows_core::IUnknown;

    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `_PropertyInfo` to be used as an `IUnknown`
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`,
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `_PropertyInfo` interface.
#[repr(C)]
pub struct _PropertyInfo_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    
    // IDispatch methods
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,

    // Methods specific to the COM interface
    get_ToString: unsafe extern "system" fn(
        this: *mut c_void, 
        pRetVal: *mut BSTR
    ) -> HRESULT,
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
