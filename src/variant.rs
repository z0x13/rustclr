//! Helper to convert Rust types into COM `VARIANT` and build `SAFEARRAY`.

use alloc::{string::String, vec::Vec};
use core::ffi::c_void;
use core::ptr::{copy_nonoverlapping, null_mut};

use windows_core::Interface;
use windows_sys::Win32::{
    Foundation::{VARIANT_FALSE, VARIANT_TRUE},
    System::{
        Com::SAFEARRAYBOUND,
        Ole::{
            SafeArrayAccessData, SafeArrayCreate, SafeArrayCreateVector,
            SafeArrayPutElement, SafeArrayUnaccessData,
        },
        Variant::{
            VARIANT, VT_ARRAY, VT_BOOL, VT_BSTR, VT_I4, VT_I8,
            VT_UI1, VT_UNKNOWN, VT_VARIANT,
        },
    },
};

use crate::error::{ClrError, Result};
use crate::wrappers::{Bstr, SafeArray as SafeArrayWrapper};

/// Trait to convert various Rust types to Windows COM-compatible `VARIANT` types.
pub trait Variant {
    /// Converts the Rust type to a `VARIANT`.
    fn to_variant(&self) -> VARIANT;

    /// Returns the `u16` representing the VARIANT type.
    fn var_type() -> u16;
}

impl Variant for String {
    /// Converts a `String` to a BSTR-based `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let bstr = Bstr::from(self.as_str());
        let mut variant = unsafe { core::mem::zeroed::<VARIANT>() };
        variant.Anonymous.Anonymous.vt = Self::var_type();
        // Transfer ownership of BSTR to VARIANT - VariantClear will free it
        variant.Anonymous.Anonymous.Anonymous.bstrVal = bstr.into_raw();
        variant
    }

    /// Returns the VARIANT type ID for BSTRs.
    fn var_type() -> u16 {
        VT_BSTR
    }
}

impl Variant for &str {
    /// Converts a `&str` to a BSTR-based `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let bstr = Bstr::from(*self);
        let mut variant = unsafe { core::mem::zeroed::<VARIANT>() };
        variant.Anonymous.Anonymous.vt = Self::var_type();
        // Transfer ownership of BSTR to VARIANT - VariantClear will free it
        variant.Anonymous.Anonymous.Anonymous.bstrVal = bstr.into_raw();
        variant
    }

    /// Returns the VARIANT type ID for BSTRs.
    fn var_type() -> u16 {
        VT_BSTR
    }
}

impl Variant for bool {
    /// Converts a `bool` to a boolean `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let mut variant = unsafe { core::mem::zeroed::<VARIANT>() };
        variant.Anonymous.Anonymous.vt = Self::var_type();
        variant.Anonymous.Anonymous.Anonymous.boolVal =
            if *self { VARIANT_TRUE } else { VARIANT_FALSE };
        variant
    }

    /// Returns the VARIANT type ID for booleans.
    fn var_type() -> u16 {
        VT_BOOL
    }
}

impl Variant for i32 {
    /// Converts an `i32` to an integer `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let mut variant = unsafe { core::mem::zeroed::<VARIANT>() };
        variant.Anonymous.Anonymous.vt = Self::var_type();
        variant.Anonymous.Anonymous.Anonymous.lVal = *self;
        variant
    }

    /// Returns the VARIANT type ID for integers.
    fn var_type() -> u16 {
        VT_I4
    }
}

impl Variant for i64 {
    /// Converts an `i64` to an integer `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let mut variant = unsafe { core::mem::zeroed::<VARIANT>() };
        variant.Anonymous.Anonymous.vt = Self::var_type();
        variant.Anonymous.Anonymous.Anonymous.llVal = *self;
        variant
    }

    /// Returns the VARIANT type ID for integers.
    fn var_type() -> u16 {
        VT_I8
    }
}

impl Variant for windows_core::IUnknown {
    /// Converts an `IUnknown` to an integer `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let mut variant = unsafe { core::mem::zeroed::<VARIANT>() };
        variant.Anonymous.Anonymous.vt = VT_UNKNOWN;
        variant.Anonymous.Anonymous.Anonymous.punkVal = self.as_raw() as *mut _;
        variant
    }

    /// Returns the VARIANT type ID for integers.
    fn var_type() -> u16 {
        VT_UNKNOWN
    }
}

/// Creates a `SAFEARRAY` from a vector of elements implementing the `Variant` trait.
/// Returns an owned SafeArray wrapper that auto-destroys on drop.
pub fn create_safe_array_args<T: Variant>(args: Vec<T>) -> Result<SafeArrayWrapper> {
    unsafe {
        let vartype = T::var_type();
        let psa = SafeArrayCreateVector(vartype, 0, args.len() as u32);
        if psa.is_null() {
            return Err(ClrError::NullPointerError("SafeArrayCreateVector"));
        }

        for (i, arg) in args.iter().enumerate() {
            let variant = arg.to_variant();
            let index = i as i32;
            let value_ptr = match vartype {
                VT_BOOL => {
                    &variant.Anonymous.Anonymous.Anonymous.boolVal as *const _ as *const c_void
                }
                VT_I4 => &variant.Anonymous.Anonymous.Anonymous.lVal as *const _ as *const c_void,
                VT_BSTR => variant.Anonymous.Anonymous.Anonymous.bstrVal as *const c_void,
                _ => return Err(ClrError::VariantUnsupported),
            };

            let hr = SafeArrayPutElement(psa, &index, value_ptr);
            if hr != 0 {
                // SafeArrayPutElement copies the BSTR, so we need to clean up the variant
                if vartype == VT_BSTR {
                    let bstr = variant.Anonymous.Anonymous.Anonymous.bstrVal;
                    if !bstr.is_null() {
                        windows_sys::Win32::Foundation::SysFreeString(bstr);
                    }
                }
                // Destroy the partially filled array
                windows_sys::Win32::System::Ole::SafeArrayDestroy(psa);
                return Err(ClrError::ApiError("SafeArrayPutElement", hr));
            }

            // SafeArrayPutElement copies the BSTR, so we need to free our copy
            if vartype == VT_BSTR {
                let bstr = variant.Anonymous.Anonymous.Anonymous.bstrVal;
                if !bstr.is_null() {
                    windows_sys::Win32::Foundation::SysFreeString(bstr);
                }
            }
        }

        // Wrap in VT_ARRAY | vartype VARIANT, then wrap in SAFEARRAY of VARIANT
        let args_array = SafeArrayCreateVector(VT_VARIANT, 0, 1);
        if args_array.is_null() {
            windows_sys::Win32::System::Ole::SafeArrayDestroy(psa);
            return Err(ClrError::NullPointerError("SafeArrayCreateVector (2)"));
        }

        let mut var_array = core::mem::zeroed::<VARIANT>();
        var_array.Anonymous.Anonymous.vt = VT_ARRAY | vartype;
        var_array.Anonymous.Anonymous.Anonymous.parray = psa;

        let index = 0;
        let hr = SafeArrayPutElement(
            args_array,
            &index,
            &mut var_array as *const VARIANT as *const c_void,
        );
        if hr != 0 {
            // var_array owns psa now via the parray field, but SafeArrayPutElement failed
            // so we need to clean up both
            windows_sys::Win32::System::Ole::SafeArrayDestroy(psa);
            windows_sys::Win32::System::Ole::SafeArrayDestroy(args_array);
            return Err(ClrError::ApiError("SafeArrayPutElement (2)", hr));
        }

        Ok(SafeArrayWrapper::from_raw(args_array))
    }
}

/// Creates a `SAFEARRAY` from a vector of `VARIANT` elements.
/// Returns an owned SafeArray wrapper that auto-destroys on drop.
///
/// Note: The VARIANTs are copied into the SAFEARRAY and then cleared.
/// This function takes ownership and cleans up all input VARIANTs.
pub fn create_safe_args(mut args: Vec<VARIANT>) -> Result<SafeArrayWrapper> {
    unsafe {
        let arg = SafeArrayCreateVector(VT_VARIANT, 0, args.len() as u32);
        if arg.is_null() {
            // Clean up all VARIANTs before returning error
            for var in args.iter_mut() {
                windows_sys::Win32::System::Variant::VariantClear(var);
            }
            return Err(ClrError::NullPointerError("SafeArrayCreateVector"));
        }

        for (i, var) in args.iter().enumerate() {
            let index = i as i32;
            let hr = SafeArrayPutElement(
                arg,
                &index,
                var as *const VARIANT as *const c_void,
            );
            if hr != 0 {
                windows_sys::Win32::System::Ole::SafeArrayDestroy(arg);
                // Clean up all VARIANTs before returning error
                for var in args.iter_mut() {
                    windows_sys::Win32::System::Variant::VariantClear(var);
                }
                return Err(ClrError::ApiError("SafeArrayPutElement", hr));
            }
        }

        // SafeArrayPutElement copies VARIANTs, so clear the originals to prevent leaks
        for var in args.iter_mut() {
            windows_sys::Win32::System::Variant::VariantClear(var);
        }

        Ok(SafeArrayWrapper::from_raw(arg))
    }
}

/// Creates a `SAFEARRAY` from a byte buffer for loading assemblies.
/// Returns an owned SafeArray wrapper that auto-destroys on drop.
pub fn create_safe_array_buffer(data: &[u8]) -> Result<SafeArrayWrapper> {
    let len: u32 = data.len() as u32;
    let bounds = SAFEARRAYBOUND {
        cElements: data.len() as _,
        lLbound: 0,
    };

    unsafe {
        let sa = SafeArrayCreate(VT_UI1, 1, &bounds);
        if sa.is_null() {
            return Err(ClrError::NullPointerError("SafeArrayCreate"));
        }

        let mut p_data = null_mut();
        let mut hr = SafeArrayAccessData(sa, &mut p_data);
        if hr != 0 {
            windows_sys::Win32::System::Ole::SafeArrayDestroy(sa);
            return Err(ClrError::ApiError("SafeArrayAccessData", hr));
        }

        copy_nonoverlapping(data.as_ptr(), p_data as *mut u8, len as usize);
        hr = SafeArrayUnaccessData(sa);
        if hr != 0 {
            windows_sys::Win32::System::Ole::SafeArrayDestroy(sa);
            return Err(ClrError::ApiError("SafeArrayUnaccessData", hr));
        }

        Ok(SafeArrayWrapper::from_raw(sa))
    }
}
