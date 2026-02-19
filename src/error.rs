use alloc::string::String;
use const_encrypt::obf;
use core::fmt;

pub type Result<T> = core::result::Result<T, ClrError>;

#[derive(Debug)]
pub enum ClrError {
    FileReadError(String),
    ApiError(String, i32),
    MissingArguments,
    CastingError(String),
    InvalidExecutable,
    MethodNotFound,
    PropertyNotFound,
    NotDotNet,
    MetaHostCreationError(String),
    RuntimeInfoError(String),
    RuntimeHostError(String),
    RuntimeStartError,
    DomainCreationError(String),
    DefaultDomainError(String),
    NoDomainAvailable,
    NullPointerError(String),
    SafeArrayError(String),
    VariantUnsupported,
    Msg(String),
    Message(String),
    InvalidNtHeader,
}

impl fmt::Display for ClrError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::FileReadError(s) => write!(f, "{}: {s}", obf!("file read error")),
            Self::ApiError(name, hr) => {
                write!(f, "{name} {}: {hr}", obf!("failed with HRESULT"))
            }
            Self::MissingArguments => {
                write!(
                    f,
                    "{}",
                    obf!("entrypoint expected arguments but received none")
                )
            }
            Self::CastingError(s) => write!(f, "{}: {s}", obf!("interface cast failed")),
            Self::InvalidExecutable => {
                write!(f, "{}", obf!("invalid or unsupported executable buffer"))
            }
            Self::MethodNotFound => write!(f, "{}", obf!("method not found")),
            Self::PropertyNotFound => write!(f, "{}", obf!("property not found")),
            Self::NotDotNet => write!(f, "{}", obf!("not a .NET application")),
            Self::MetaHostCreationError(s) => {
                write!(f, "{}: {s}", obf!("metahost creation failed"))
            }
            Self::RuntimeInfoError(s) => {
                write!(f, "{}: {s}", obf!("runtime info retrieval failed"))
            }
            Self::RuntimeHostError(s) => {
                write!(f, "{}: {s}", obf!("runtime host retrieval failed"))
            }
            Self::RuntimeStartError => write!(f, "{}", obf!("failed to start CLR runtime")),
            Self::DomainCreationError(s) => write!(f, "{}: {s}", obf!("domain creation failed")),
            Self::DefaultDomainError(s) => {
                write!(f, "{}: {s}", obf!("default domain retrieval failed"))
            }
            Self::NoDomainAvailable => {
                write!(f, "{}", obf!("no domain available in current runtime"))
            }
            Self::NullPointerError(s) => write!(f, "{}: {s}", obf!("null pointer")),
            Self::SafeArrayError(s) => write!(f, "{}: {s}", obf!("safearray creation failed")),
            Self::VariantUnsupported => write!(f, "{}", obf!("unsupported VARIANT type")),
            Self::Msg(s) => f.write_str(s),
            Self::Message(s) => f.write_str(s),
            Self::InvalidNtHeader => write!(f, "{}", obf!("invalid PE file")),
        }
    }
}

impl core::error::Error for ClrError {}
