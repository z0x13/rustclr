use alloc::string::ToString;
use core::{ffi::c_void, mem::transmute, ops::Deref, ptr::null_mut};

use windows::core::{GUID, HRESULT, IUnknown, Interface};

use const_encrypt::obf;

use crate::error::{ClrError, Result};

/// This struct represents the COM `IEnumUnknown` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct IEnumUnknown(windows::core::IUnknown);

impl IEnumUnknown {
    /// Creates an `IEnumUnknown` from a raw pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Self {
        unsafe { Self(IUnknown::from_raw(raw)) }
    }

    /// Retrieves the next set of interfaces from the enumerator.
    #[inline]
    pub fn Next(
        &self,
        rgelt: &mut [Option<windows::core::IUnknown>],
        pceltfetched: Option<*mut u32>,
    ) -> HRESULT {
        unsafe {
            (Interface::vtable(self).Next)(
                Interface::as_raw(self),
                rgelt.len() as u32,
                transmute(rgelt.as_ptr()),
                transmute(pceltfetched.unwrap_or_default()),
            )
        }
    }

    /// Skips a specified number of elements in the enumeration sequence.
    #[inline]
    pub fn Skip(&self, celt: u32) -> Result<()> {
        let hr = unsafe { (Interface::vtable(self).Skip)(Interface::as_raw(self), celt) };
        if hr.is_ok() {
            Ok(())
        } else {
            Err(ClrError::ApiError(obf!("Skip").to_string(), hr.0))
        }
    }

    /// Resets the enumeration sequence to the beginning.
    #[inline]
    pub fn Reset(&self) -> Result<()> {
        let hr = unsafe { (Interface::vtable(self).Reset)(Interface::as_raw(self)) };
        if hr.is_ok() {
            Ok(())
        } else {
            Err(ClrError::ApiError(obf!("Reset").to_string(), hr.0))
        }
    }

    /// Creates a new enumerator with the same state as the current one.
    #[inline]
    pub fn Clone(&self) -> Result<*mut IEnumUnknown> {
        let mut result = null_mut();
        let hr = unsafe { (Interface::vtable(self).Clone)(Interface::as_raw(self), &mut result) };
        if hr.is_ok() {
            Ok(result)
        } else {
            Err(ClrError::ApiError(obf!("Clone").to_string(), hr.0))
        }
    }
}

unsafe impl Interface for IEnumUnknown {
    type Vtable = IEnumUnknown_Vtbl;
    const IID: GUID = GUID::from_u128(0x00000100_0000_0000_c000_000000000046);
}

impl Deref for IEnumUnknown {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

#[repr(C)]
pub struct IEnumUnknown_Vtbl {
    pub base__: windows::core::IUnknown_Vtbl,
    pub Next: unsafe extern "system" fn(
        this: *mut c_void,
        celt: u32,
        rgelt: *mut *mut IUnknown,
        pceltFetched: *mut u32,
    ) -> HRESULT,
    pub Skip: unsafe extern "system" fn(this: *mut c_void, celt: u32) -> HRESULT,
    pub Reset: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub Clone:
        unsafe extern "system" fn(this: *mut c_void, ppenum: *mut *mut IEnumUnknown) -> HRESULT,
}
