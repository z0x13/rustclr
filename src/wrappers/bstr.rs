//! Owned BSTR wrapper with automatic cleanup.

use alloc::{string::String, vec::Vec};
use core::ops::Deref;
use windows_sys::Win32::Foundation::{SysAllocStringLen, SysFreeString, SysStringLen};

/// Owned BSTR that calls `SysFreeString` on drop.
#[repr(transparent)]
pub struct Bstr(*const u16);

impl Bstr {
    /// Creates an empty BSTR (null pointer, no allocation).
    #[inline]
    pub const fn new() -> Self {
        Self(core::ptr::null())
    }

    /// Creates a BSTR from a UTF-16 slice.
    pub fn from_wide(value: &[u16]) -> Self {
        if value.is_empty() {
            return Self::new();
        }
        unsafe { Self(SysAllocStringLen(value.as_ptr(), value.len() as u32)) }
    }

    /// Creates a BSTR from a UTF-8 string.
    pub fn from_str(value: &str) -> Self {
        if value.is_empty() {
            return Self::new();
        }
        let wide: Vec<u16> = value.encode_utf16().collect();
        Self::from_wide(&wide)
    }

    /// Returns true if the BSTR is null/empty.
    #[inline]
    pub fn is_null(&self) -> bool {
        self.0.is_null()
    }

    /// Returns the raw pointer for passing to COM methods (borrowing).
    #[inline]
    pub fn as_ptr(&self) -> *const u16 {
        self.0
    }

    /// Consumes self and returns raw pointer, preventing Drop.
    /// Caller takes ownership and must call `SysFreeString`.
    #[inline]
    pub fn into_raw(self) -> *const u16 {
        let ptr = self.0;
        core::mem::forget(self);
        ptr
    }

    /// Takes ownership of an existing BSTR pointer.
    ///
    /// # Safety
    ///
    /// Pointer must be a valid BSTR allocated by `SysAllocString` or null.
    #[inline]
    pub unsafe fn from_raw(raw: *const u16) -> Self {
        Self(raw)
    }

    /// Converts to Rust String (lossy UTF-16 to UTF-8).
    pub fn to_string_lossy(&self) -> String {
        if self.0.is_null() {
            return String::new();
        }
        let len = unsafe { SysStringLen(self.0) } as usize;
        if len == 0 {
            return String::new();
        }
        let slice = unsafe { core::slice::from_raw_parts(self.0, len) };
        String::from_utf16_lossy(slice)
    }
}

impl Deref for Bstr {
    type Target = [u16];

    fn deref(&self) -> &[u16] {
        if self.0.is_null() {
            return &[];
        }
        let len = unsafe { SysStringLen(self.0) } as usize;
        if len > 0 {
            unsafe { core::slice::from_raw_parts(self.0, len) }
        } else {
            &[]
        }
    }
}

impl Clone for Bstr {
    fn clone(&self) -> Self {
        Self::from_wide(self)
    }
}

impl Default for Bstr {
    fn default() -> Self {
        Self::new()
    }
}

impl Drop for Bstr {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe { SysFreeString(self.0) };
        }
    }
}

impl From<&str> for Bstr {
    fn from(value: &str) -> Self {
        Self::from_str(value)
    }
}

impl From<String> for Bstr {
    fn from(value: String) -> Self {
        Self::from_str(&value)
    }
}

impl From<&String> for Bstr {
    fn from(value: &String) -> Self {
        Self::from_str(value)
    }
}
