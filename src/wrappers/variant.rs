//! Owned VARIANT wrapper with automatic cleanup.

use alloc::string::String;
use windows_sys::Win32::Foundation::{SysStringLen, VARIANT_FALSE, VARIANT_TRUE};
use windows_sys::Win32::System::Variant::{
    VARIANT, VariantClear, VariantCopy,
    VT_BOOL, VT_BSTR, VT_EMPTY, VT_I4, VT_I8, VT_UNKNOWN,
};

use super::Bstr;

/// Owned VARIANT that calls `VariantClear` on drop.
#[repr(transparent)]
pub struct OwnedVariant(VARIANT);

impl OwnedVariant {
    /// Creates an empty VARIANT (VT_EMPTY).
    pub fn empty() -> Self {
        Self(unsafe { core::mem::zeroed() })
    }

    /// Returns the variant type.
    #[inline]
    pub fn vt(&self) -> u16 {
        unsafe { self.0.Anonymous.Anonymous.vt }
    }

    /// Returns true if VT_EMPTY.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.vt() == VT_EMPTY
    }

    /// Returns a reference to the inner VARIANT for reading.
    #[inline]
    pub fn as_raw(&self) -> &VARIANT {
        &self.0
    }

    /// Returns a mutable reference to the inner VARIANT.
    #[inline]
    pub fn as_raw_mut(&mut self) -> &mut VARIANT {
        &mut self.0
    }

    /// Returns a mutable pointer for COM out parameters.
    #[inline]
    pub fn as_mut_ptr(&mut self) -> *mut VARIANT {
        &mut self.0
    }

    /// Consumes self and returns the inner VARIANT, preventing Drop.
    /// Caller takes ownership and must call VariantClear.
    #[inline]
    pub fn into_raw(self) -> VARIANT {
        let var = unsafe { core::ptr::read(&self.0) };
        core::mem::forget(self);
        var
    }

    /// Takes ownership of an existing VARIANT.
    ///
    /// # Safety
    ///
    /// The VARIANT must be valid and caller transfers ownership.
    #[inline]
    pub unsafe fn from_raw(var: VARIANT) -> Self {
        Self(var)
    }

    /// Extracts BSTR if VT_BSTR, taking ownership.
    /// Returns None if not VT_BSTR or null.
    pub fn take_bstr(&mut self) -> Option<Bstr> {
        unsafe {
            if self.0.Anonymous.Anonymous.vt != VT_BSTR {
                return None;
            }
            let bstr = self.0.Anonymous.Anonymous.Anonymous.bstrVal;
            if bstr.is_null() {
                return None;
            }
            // Clear from variant to prevent double-free
            self.0.Anonymous.Anonymous.Anonymous.bstrVal = core::ptr::null();
            self.0.Anonymous.Anonymous.vt = VT_EMPTY;
            Some(Bstr::from_raw(bstr))
        }
    }

    /// Gets BSTR value as String without taking ownership (copies the string).
    pub fn get_bstr_string(&self) -> Option<String> {
        unsafe {
            if self.0.Anonymous.Anonymous.vt != VT_BSTR {
                return None;
            }
            let bstr = self.0.Anonymous.Anonymous.Anonymous.bstrVal;
            if bstr.is_null() {
                return Some(String::new());
            }
            let len = SysStringLen(bstr) as usize;
            let slice = core::slice::from_raw_parts(bstr, len);
            Some(String::from_utf16_lossy(slice))
        }
    }

    /// Gets the i32 value if VT_I4.
    pub fn get_i32(&self) -> Option<i32> {
        unsafe {
            if self.0.Anonymous.Anonymous.vt == VT_I4 {
                Some(self.0.Anonymous.Anonymous.Anonymous.lVal)
            } else {
                None
            }
        }
    }

    /// Gets the i64 value if VT_I8.
    pub fn get_i64(&self) -> Option<i64> {
        unsafe {
            if self.0.Anonymous.Anonymous.vt == VT_I8 {
                Some(self.0.Anonymous.Anonymous.Anonymous.llVal)
            } else {
                None
            }
        }
    }

    /// Gets the bool value if VT_BOOL.
    pub fn get_bool(&self) -> Option<bool> {
        unsafe {
            if self.0.Anonymous.Anonymous.vt == VT_BOOL {
                Some(self.0.Anonymous.Anonymous.Anonymous.boolVal != 0)
            } else {
                None
            }
        }
    }

    /// Gets the IUnknown pointer if VT_UNKNOWN.
    /// Does NOT take ownership - the pointer is still owned by the VARIANT.
    pub fn get_unknown_ptr(&self) -> Option<*mut core::ffi::c_void> {
        unsafe {
            if self.0.Anonymous.Anonymous.vt == VT_UNKNOWN {
                Some(self.0.Anonymous.Anonymous.Anonymous.punkVal)
            } else {
                None
            }
        }
    }

    /// Gets the byref pointer.
    pub fn get_byref(&self) -> *mut core::ffi::c_void {
        unsafe { self.0.Anonymous.Anonymous.Anonymous.byref }
    }
}

impl Clone for OwnedVariant {
    fn clone(&self) -> Self {
        unsafe {
            let mut value = Self::empty();
            VariantCopy(&mut value.0, &self.0);
            value
        }
    }
}

impl Default for OwnedVariant {
    fn default() -> Self {
        Self::empty()
    }
}

impl Drop for OwnedVariant {
    fn drop(&mut self) {
        unsafe { VariantClear(&mut self.0) };
    }
}

// From implementations for OwnedVariant

impl From<i32> for OwnedVariant {
    fn from(value: i32) -> Self {
        let mut v = Self::empty();
        v.0.Anonymous.Anonymous.vt = VT_I4;
        v.0.Anonymous.Anonymous.Anonymous.lVal = value;
        v
    }
}

impl From<i64> for OwnedVariant {
    fn from(value: i64) -> Self {
        let mut v = Self::empty();
        v.0.Anonymous.Anonymous.vt = VT_I8;
        v.0.Anonymous.Anonymous.Anonymous.llVal = value;
        v
    }
}

impl From<bool> for OwnedVariant {
    fn from(value: bool) -> Self {
        let mut v = Self::empty();
        v.0.Anonymous.Anonymous.vt = VT_BOOL;
        v.0.Anonymous.Anonymous.Anonymous.boolVal =
            if value { VARIANT_TRUE } else { VARIANT_FALSE };
        v
    }
}

impl From<Bstr> for OwnedVariant {
    fn from(value: Bstr) -> Self {
        let mut v = Self::empty();
        v.0.Anonymous.Anonymous.vt = VT_BSTR;
        v.0.Anonymous.Anonymous.Anonymous.bstrVal = value.into_raw();
        v
    }
}

impl From<&str> for OwnedVariant {
    fn from(value: &str) -> Self {
        Bstr::from(value).into()
    }
}

impl From<String> for OwnedVariant {
    fn from(value: String) -> Self {
        Bstr::from(value).into()
    }
}
