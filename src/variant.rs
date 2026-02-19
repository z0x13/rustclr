//! Helper functions to build `SAFEARRAY` from Rust types.

use alloc::string::{String, ToString};
use alloc::vec::Vec;
use core::ffi::c_void;
use core::ptr::{copy_nonoverlapping, null_mut};

use windows::Win32::System::Com::SAFEARRAYBOUND;
use windows::Win32::System::Ole::{
    SafeArrayAccessData, SafeArrayCreate, SafeArrayCreateVector, SafeArrayDestroy,
    SafeArrayPutElement, SafeArrayUnaccessData,
};
use windows::Win32::System::Variant::{InitVariantFromStringArray, VARIANT, VT_UI1, VT_VARIANT};
use windows::core::PCWSTR;

use const_encrypt::obf;

use crate::error::{ClrError, Result};
use crate::wrappers::SafeArray as SafeArrayWrapper;

/// Creates a `SAFEARRAY` of VARIANTs from a vector of values that implement `Into<VARIANT>`.
/// Returns an owned SafeArray wrapper that auto-destroys on drop.
pub fn create_safe_array_args<T: Into<VARIANT>>(args: Vec<T>) -> Result<SafeArrayWrapper> {
    let variants: Vec<VARIANT> = args.into_iter().map(Into::into).collect();
    create_safe_args(variants)
}

/// Creates a `SAFEARRAY` from a vector of `VARIANT` elements.
/// Returns an owned SafeArray wrapper that auto-destroys on drop.
pub fn create_safe_args(args: Vec<VARIANT>) -> Result<SafeArrayWrapper> {
    unsafe {
        let sa = SafeArrayCreateVector(VT_VARIANT, 0, args.len() as u32);
        if sa.is_null() {
            return Err(ClrError::NullPointerError(
                obf!("SafeArrayCreateVector").to_string(),
            ));
        }

        for (i, var) in args.iter().enumerate() {
            let index = i as i32;
            if let Err(err) =
                SafeArrayPutElement(sa, &index, var as *const VARIANT as *const c_void)
            {
                let _ = SafeArrayDestroy(sa);
                return Err(ClrError::ApiError(
                    obf!("SafeArrayPutElement").to_string(),
                    err.code().0,
                ));
            }
        }

        Ok(SafeArrayWrapper::from_raw(sa))
    }
}

/// Creates a `SAFEARRAY` from a byte buffer for loading assemblies.
/// Returns an owned SafeArray wrapper that auto-destroys on drop.
pub fn create_safe_array_buffer(data: &[u8]) -> Result<SafeArrayWrapper> {
    let bounds = SAFEARRAYBOUND {
        cElements: data.len() as u32,
        lLbound: 0,
    };

    unsafe {
        let sa = SafeArrayCreate(VT_UI1, 1, &bounds);
        if sa.is_null() {
            return Err(ClrError::NullPointerError(
                obf!("SafeArrayCreate").to_string(),
            ));
        }

        let mut p_data = null_mut();
        if let Err(err) = SafeArrayAccessData(sa, &mut p_data) {
            let _ = SafeArrayDestroy(sa);
            return Err(ClrError::ApiError(
                obf!("SafeArrayAccessData").to_string(),
                err.code().0,
            ));
        }

        copy_nonoverlapping(data.as_ptr(), p_data as *mut u8, data.len());

        if let Err(err) = SafeArrayUnaccessData(sa) {
            let _ = SafeArrayDestroy(sa);
            return Err(ClrError::ApiError(
                obf!("SafeArrayUnaccessData").to_string(),
                err.code().0,
            ));
        }

        Ok(SafeArrayWrapper::from_raw(sa))
    }
}

/// Creates a VARIANT containing a string array (VT_ARRAY | VT_BSTR).
/// Used for passing string[] arguments to .NET methods like Main(string[]).
pub fn create_string_array_variant(strings: Vec<String>) -> Result<VARIANT> {
    let wide_strings: Vec<Vec<u16>> = strings
        .iter()
        .map(|s| s.encode_utf16().chain(Some(0)).collect())
        .collect();

    let pcwstr_vec: Vec<PCWSTR> = wide_strings
        .iter()
        .map(|s| PCWSTR::from_raw(s.as_ptr()))
        .collect();

    let variant = unsafe { InitVariantFromStringArray(&pcwstr_vec) }.map_err(|err| {
        ClrError::ApiError(obf!("InitVariantFromStringArray").to_string(), err.code().0)
    })?;

    Ok(variant)
}
