//! Defines the `ComString` trait for converting between Rust strings and BSTRs.

use alloc::string::String;
use windows_sys::Win32::Foundation::SysStringLen;

use crate::wrappers::Bstr;

/// The `ComString` trait provides methods for working with BSTRs.
pub trait ComString {
    /// Converts a Rust string into an owned BSTR wrapper.
    fn to_bstr_owned(&self) -> Bstr;

    /// Converts a BSTR to a Rust string.
    fn to_string(&self) -> String {
        String::new()
    }
}

impl ComString for &str {
    fn to_bstr_owned(&self) -> Bstr {
        Bstr::from(*self)
    }
}

impl ComString for String {
    fn to_bstr_owned(&self) -> Bstr {
        Bstr::from(self.as_str())
    }
}

impl ComString for *const u16 {
    fn to_bstr_owned(&self) -> Bstr {
        if self.is_null() {
            return Bstr::new();
        }
        let len = unsafe { SysStringLen(*self) } as usize;
        if len == 0 {
            return Bstr::new();
        }
        let slice = unsafe { core::slice::from_raw_parts(*self, len) };
        Bstr::from_wide(slice)
    }

    fn to_string(&self) -> String {
        if self.is_null() {
            return String::new();
        }
        let len = unsafe { SysStringLen(*self) };
        if len == 0 {
            return String::new();
        }
        let slice = unsafe { core::slice::from_raw_parts(*self, len as usize) };
        String::from_utf16_lossy(slice)
    }
}
