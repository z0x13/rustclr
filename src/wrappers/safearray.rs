//! Owned SAFEARRAY wrapper with automatic cleanup.

use alloc::string::ToString;
use core::ptr::NonNull;
use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::Ole::{SafeArrayAccessData, SafeArrayDestroy, SafeArrayUnaccessData};

use const_encrypt::obf;

use crate::error::{ClrError, Result};

/// Owned SAFEARRAY that calls `SafeArrayDestroy` on drop.
pub struct SafeArray(NonNull<SAFEARRAY>);

impl SafeArray {
    /// Takes ownership of a SAFEARRAY pointer.
    /// Returns None if null.
    #[inline]
    pub fn from_ptr(ptr: *mut SAFEARRAY) -> Option<Self> {
        NonNull::new(ptr).map(Self)
    }

    /// Takes ownership of a SAFEARRAY pointer.
    ///
    /// # Safety
    ///
    /// Pointer must be valid and non-null.
    #[inline]
    pub unsafe fn from_raw(ptr: *mut SAFEARRAY) -> Self {
        unsafe { Self(NonNull::new_unchecked(ptr)) }
    }

    /// Returns raw pointer for COM methods (borrowing).
    #[inline]
    pub fn as_ptr(&self) -> *mut SAFEARRAY {
        self.0.as_ptr()
    }

    /// Consumes self and returns raw pointer, preventing Drop.
    /// Caller takes ownership and must call `SafeArrayDestroy`.
    #[inline]
    pub fn into_raw(self) -> *mut SAFEARRAY {
        let ptr = self.0.as_ptr();
        core::mem::forget(self);
        ptr
    }

    /// Returns the number of elements (for 1D arrays).
    pub fn len(&self) -> u32 {
        unsafe { (*self.0.as_ptr()).rgsabound[0].cElements }
    }

    /// Returns true if empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

impl Drop for SafeArray {
    fn drop(&mut self) {
        unsafe {
            let _ = SafeArrayDestroy(self.0.as_ptr());
        }
    }
}

/// RAII accessor that locks a SafeArray for data access.
/// Unlocks automatically on drop.
pub struct SafeArrayAccessor<'a, T> {
    array: &'a SafeArray,
    data: *mut T,
}

impl<'a, T> SafeArrayAccessor<'a, T> {
    /// Locks the array for access.
    ///
    /// # Safety
    ///
    /// Array must contain elements of type T.
    pub unsafe fn new(array: &'a SafeArray) -> Result<Self> {
        let mut data = core::ptr::null_mut();
        unsafe { SafeArrayAccessData(array.as_ptr(), &mut data) }.map_err(|err| {
            ClrError::ApiError(obf!("SafeArrayAccessData").to_string(), err.code().0)
        })?;
        Ok(Self {
            array,
            data: data as *mut T,
        })
    }

    /// Returns an iterator over the elements.
    pub fn iter(&self) -> impl Iterator<Item = &T> {
        let len = self.array.len() as usize;
        (0..len).map(move |i| unsafe { &*self.data.add(i) })
    }

    /// Returns a mutable iterator over the elements.
    pub fn iter_mut(&mut self) -> impl Iterator<Item = &mut T> {
        let len = self.array.len() as usize;
        let data = self.data;
        (0..len).map(move |i| unsafe { &mut *data.add(i) })
    }
}

impl<T> Drop for SafeArrayAccessor<'_, T> {
    fn drop(&mut self) {
        unsafe {
            let _ = SafeArrayUnaccessData(self.array.as_ptr());
        }
    }
}
