//! RAII wrappers for COM resources (BSTR, VARIANT, SAFEARRAY).

mod bstr;
mod variant;
mod safearray;

pub use bstr::Bstr;
pub use variant::OwnedVariant;
pub use safearray::SafeArray;
