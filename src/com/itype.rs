use alloc::string::{String, ToString};
use alloc::vec::Vec;
use bitflags::bitflags;
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows::core::{BSTR, GUID, HRESULT, IUnknown, Interface};
use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::Ole::{SafeArrayDestroy, SafeArrayGetElement, SafeArrayGetLBound, SafeArrayGetUBound};
use windows::Win32::System::Variant::VARIANT;

use crate::com::{_MethodInfo, _PropertyInfo};
use crate::error::{ClrError, Result};
use crate::variant::create_safe_args;
use crate::Invocation;

/// This struct represents the COM `_Type` interface.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct _Type(windows::core::IUnknown);

impl _Type {
    /// Retrieves a method by its name from the type.
    #[inline]
    pub fn method(&self, name: &str) -> Result<_MethodInfo> {
        let method_name = BSTR::from(name);
        self.GetMethod_6(method_name.as_ptr())
    }

    /// Finds a method by signature from the type.
    #[inline]
    pub fn method_signature(&self, name: &str) -> Result<_MethodInfo> {
        let methods = self.methods();
        if let Ok(methods) = methods {
            for (method_name, method_info) in methods {
                if method_name == name {
                    return Ok(method_info);
                }
            }
        }

        Err(ClrError::MethodNotFound)
    }

    /// Finds a property by signature from the type.
    #[inline]
    pub fn property_signature(&self, name: &str) -> Result<_PropertyInfo> {
        let properties = self.properties();
        if let Ok(properties) = properties {
            for (property_name, property_info) in properties {
                if property_name == name {
                    return Ok(property_info);
                }
            }
        }

        Err(ClrError::PropertyNotFound)
    }

    /// Retrieves a property by name from the type.
    #[inline]
    pub fn property(&self, name: &str) -> Result<_PropertyInfo> {
        unsafe {
            let binding_flags = BindingFlags::Public
                | BindingFlags::Instance
                | BindingFlags::Static
                | BindingFlags::FlattenHierarchy
                | BindingFlags::NonPublic;

            let property_name = BSTR::from(name);
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetProperty)(
                Interface::as_raw(self),
                property_name.as_ptr(),
                binding_flags,
                &mut result,
            );

            if hr.is_ok() && !result.is_null() {
                Ok(_PropertyInfo::from_raw(result)?)
            } else {
                Err(ClrError::ApiError("GetProperty", hr.0))
            }
        }
    }

    /// Invokes a method on the type.
    #[inline]
    pub fn invoke(
        &self,
        name: &str,
        instance: Option<VARIANT>,
        args: Option<Vec<VARIANT>>,
        invocation_type: Invocation,
    ) -> Result<VARIANT> {
        let flags = match invocation_type {
            Invocation::Static => {
                BindingFlags::NonPublic
                    | BindingFlags::Public
                    | BindingFlags::Static
                    | BindingFlags::InvokeMethod
            }
            Invocation::Instance => {
                BindingFlags::NonPublic
                    | BindingFlags::Public
                    | BindingFlags::Instance
                    | BindingFlags::InvokeMethod
            }
        };

        let method_name = BSTR::from(name);
        let args_array = args.map(create_safe_args).transpose()?;
        let args_ptr = args_array.as_ref().map_or(null_mut(), |a| a.as_ptr());

        let instance_var = instance.unwrap_or_default();
        self.InvokeMember_3(method_name.as_ptr(), flags, instance_var, args_ptr)
    }

    /// Retrieves all methods of the type.
    #[inline]
    pub fn methods(&self) -> Result<Vec<(String, _MethodInfo)>> {
        let binding_flags = BindingFlags::Public
            | BindingFlags::Instance
            | BindingFlags::Static
            | BindingFlags::FlattenHierarchy
            | BindingFlags::NonPublic;

        let sa_methods = self.GetMethods(binding_flags)?;
        if sa_methods.is_null() {
            return Err(ClrError::NullPointerError("GetMethods"));
        }

        let mut methods = Vec::new();
        unsafe {
            let lbound = SafeArrayGetLBound(sa_methods, 1)
                .map_err(|err| ClrError::ApiError("SafeArrayGetLBound", err.code().0))?;
            let ubound = SafeArrayGetUBound(sa_methods, 1)
                .map_err(|err| ClrError::ApiError("SafeArrayGetUBound", err.code().0))?;

            let mut p_method = null_mut::<_MethodInfo>();
            for i in lbound..=ubound {
                if let Err(err) = SafeArrayGetElement(sa_methods, &i, &mut p_method as *mut _ as *mut _) {
                    let _ = SafeArrayDestroy(sa_methods);
                    return Err(ClrError::ApiError("SafeArrayGetElement", err.code().0));
                }
                if p_method.is_null() {
                    let _ = SafeArrayDestroy(sa_methods);
                    return Err(ClrError::NullPointerError("SafeArrayGetElement"));
                }

                let method = _MethodInfo::from_raw(p_method as *mut c_void)?;
                let method_name = method.ToString()?;
                methods.push((method_name, method));
            }

            let _ = SafeArrayDestroy(sa_methods);
        }

        Ok(methods)
    }

    /// Retrieves all properties of the type.
    #[inline]
    pub fn properties(&self) -> Result<Vec<(String, _PropertyInfo)>> {
        let binding_flags = BindingFlags::Public
            | BindingFlags::Instance
            | BindingFlags::Static
            | BindingFlags::FlattenHierarchy
            | BindingFlags::NonPublic;

        let sa_properties = self.GetProperties(binding_flags)?;
        if sa_properties.is_null() {
            return Err(ClrError::NullPointerError("GetProperties"));
        }

        let mut properties = Vec::new();
        unsafe {
            let lbound = SafeArrayGetLBound(sa_properties, 1)
                .map_err(|err| ClrError::ApiError("SafeArrayGetLBound", err.code().0))?;
            let ubound = SafeArrayGetUBound(sa_properties, 1)
                .map_err(|err| ClrError::ApiError("SafeArrayGetUBound", err.code().0))?;

            let mut p_property = null_mut::<_PropertyInfo>();
            for i in lbound..=ubound {
                if let Err(err) = SafeArrayGetElement(sa_properties, &i, &mut p_property as *mut _ as *mut _) {
                    let _ = SafeArrayDestroy(sa_properties);
                    return Err(ClrError::ApiError("SafeArrayGetElement", err.code().0));
                }
                if p_property.is_null() {
                    let _ = SafeArrayDestroy(sa_properties);
                    return Err(ClrError::NullPointerError("SafeArrayGetElement"));
                }

                let property = _PropertyInfo::from_raw(p_property as *mut c_void)?;
                let name = property.ToString()?;
                properties.push((name, property));
            }

            let _ = SafeArrayDestroy(sa_properties);
        }

        Ok(properties)
    }

    /// Creates an `_Type` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_Type> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_Type>()
            .map_err(|_| ClrError::CastingError("_Type"))
    }

    /// Retrieves the string representation of the type.
    #[inline]
    pub fn ToString(&self) -> Result<String> {
        unsafe {
            let mut result: *const u16 = core::ptr::null();
            let hr = (Interface::vtable(self).get_ToString)(Interface::as_raw(self), &mut result);
            if hr.is_ok() {
                let bstr = BSTR::from_raw(result);
                Ok(bstr.to_string())
            } else {
                Err(ClrError::ApiError("ToString", hr.0))
            }
        }
    }

    /// Retrieves all properties matching the specified `BindingFlags`.
    #[inline]
    pub fn GetProperties(&self, bindingAttr: BindingFlags) -> Result<*mut SAFEARRAY> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetProperties)(
                Interface::as_raw(self),
                bindingAttr,
                &mut result,
            );

            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetProperties", hr.0))
            }
        }
    }

    /// Retrieves all methods matching the specified `BindingFlags`.
    #[inline]
    pub fn GetMethods(&self, bindingAttr: BindingFlags) -> Result<*mut SAFEARRAY> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetMethods)(
                Interface::as_raw(self),
                bindingAttr,
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetMethods", hr.0))
            }
        }
    }

    /// Retrieves a method by name.
    #[inline]
    pub fn GetMethod_6(&self, name: *const u16) -> Result<_MethodInfo> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).GetMethod_6)(Interface::as_raw(self), name, &mut result);
            if hr.is_ok() {
                _MethodInfo::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("GetMethod_6", hr.0))
            }
        }
    }

    /// Invokes a method (static or instance) by name on the specified type or object.
    #[inline]
    pub fn InvokeMember_3(
        &self,
        name: *const u16,
        invoke_attr: BindingFlags,
        instance: VARIANT,
        args: *mut SAFEARRAY,
    ) -> Result<VARIANT> {
        unsafe {
            let mut result = VARIANT::default();
            let hr = (Interface::vtable(self).InvokeMember_3)(
                Interface::as_raw(self),
                name,
                invoke_attr,
                null_mut(),
                instance,
                args,
                &mut result,
            );
            if hr.is_ok() {
                Ok(result)
            } else {
                Err(ClrError::ApiError("InvokeMember_3", hr.0))
            }
        }
    }
}

unsafe impl Interface for _Type {
    type Vtable = _Type_Vtbl;
    const IID: GUID = GUID::from_u128(0xbca8b44d_aad6_3a86_8ab7_03349f4f2da2);
}

impl Deref for _Type {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

bitflags! {
    #[repr(transparent)]
    pub struct BindingFlags: u32 {
        const Default = 0;
        const IgnoreCase = 1;
        const DeclaredOnly = 2;
        const Instance = 4;
        const Static = 8;
        const Public = 16;
        const NonPublic = 32;
        const FlattenHierarchy = 64;
        const InvokeMethod = 256;
        const CreateInstance = 512;
        const GetField = 1024;
        const SetField = 2048;
        const GetProperty = 4096;
        const SetProperty = 8192;
        const PutDispProperty = 16384;
        const PutRefDispProperty = 32768;
        const ExactBinding = 65536;
        const SuppressChangeType = 131072;
        const OptionalParamBinding = 262144;
        const IgnoreReturn = 16777216;
    }
}

type BSTR_PTR = *const u16;

#[repr(C)]
pub struct _Type_Vtbl {
    pub base__: windows::core::IUnknown_Vtbl,

    // IDispatch methods
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,

    // Methods specific to the COM interface
    get_ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR_PTR) -> HRESULT,
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
    get_Guid: *const c_void,
    get_Module: *const c_void,
    get_Assembly: *const c_void,
    get_TypeHandle: *const c_void,
    get_FullName: *const c_void,
    get_Namespace: *const c_void,
    get_AssemblyQualifiedName: *const c_void,
    GetArrayRank: *const c_void,
    get_BaseType: *const c_void,
    GetConstructors: *const c_void,
    GetInterface: *const c_void,
    GetInterfaces: *const c_void,
    FindInterfaces: *const c_void,
    GetEvent: *const c_void,
    GetEvents: *const c_void,
    GetEvents_2: *const c_void,
    GetNestedTypes: *const c_void,
    GetNestedType: *const c_void,
    GetMember: *const c_void,
    GetDefaultMembers: *const c_void,
    FindMembers: *const c_void,
    GetElementType: *const c_void,
    IsSubclassOf: *const c_void,
    IsInstanceOfType: *const c_void,
    IsAssignableFrom: *const c_void,
    GetInterfaceMap: *const c_void,
    GetMethod: *const c_void,
    GetMethod_2: *const c_void,
    GetMethods: unsafe extern "system" fn(
        this: *mut c_void,
        bindingAttr: BindingFlags,
        pRetVal: *mut *mut SAFEARRAY,
    ) -> HRESULT,
    GetField: *const c_void,
    GetFields: *const c_void,
    pub GetProperty: unsafe extern "system" fn(
        this: *mut c_void,
        name: BSTR_PTR,
        bindingAttr: BindingFlags,
        result: *mut *mut c_void,
    ) -> HRESULT,
    GetProperty_2: *const c_void,
    GetProperties: unsafe extern "system" fn(
        this: *mut c_void,
        bindingAttr: BindingFlags,
        pRetVal: *mut *mut SAFEARRAY,
    ) -> HRESULT,
    GetMember_2: *const c_void,
    GetMembers: *const c_void,
    InvokeMember: *const c_void,
    get_UnderlyingSystemType: *const c_void,
    InvokeMember_2: *const c_void,
    InvokeMember_3: unsafe extern "system" fn(
        this: *mut c_void,
        name: BSTR_PTR,
        invokeAttr: BindingFlags,
        Binder: *mut c_void,
        Target: VARIANT,
        args: *mut SAFEARRAY,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    GetConstructor: *const c_void,
    GetConstructor_2: *const c_void,
    GetConstructor_3: *const c_void,
    GetConstructors_2: *const c_void,
    get_TypeInitializer: *const c_void,
    GetMethod_3: *const c_void,
    GetMethod_4: *const c_void,
    GetMethod_5: *const c_void,
    GetMethod_6: unsafe extern "system" fn(
        this: *mut c_void,
        name: BSTR_PTR,
        pRetVal: *mut *mut _MethodInfo,
    ) -> HRESULT,
    GetMethods_2: *const c_void,
    GetField_2: *const c_void,
    GetFields_2: *const c_void,
    GetInterface_2: *const c_void,
    GetEvent_2: *const c_void,
    GetProperty_3: *const c_void,
    GetProperty_4: *const c_void,
    GetProperty_5: *const c_void,
    GetProperty_6: *const c_void,
    GetProperty_7: *const c_void,
    GetProperties_2: *const c_void,
    GetNestedTypes_2: *const c_void,
    GetNestedType_2: *const c_void,
    GetMember_3: *const c_void,
    GetMembers_2: *const c_void,
    get_Attributes: *const c_void,
    get_IsNotPublic: *const c_void,
    get_IsPublic: *const c_void,
    get_IsNestedPublic: *const c_void,
    get_IsNestedPrivate: *const c_void,
    get_IsNestedFamily: *const c_void,
    get_IsNestedAssembly: *const c_void,
    get_IsNestedFamANDAssem: *const c_void,
    get_IsNestedFamORAssem: *const c_void,
    get_IsAutoLayout: *const c_void,
    get_IsLayoutSequential: *const c_void,
    get_IsExplicitLayout: *const c_void,
    get_IsClass: *const c_void,
    get_IsInterface: *const c_void,
    get_IsValueType: *const c_void,
    get_IsAbstract: *const c_void,
    get_IsSealed: *const c_void,
    get_IsEnum: *const c_void,
    get_IsSpecialName: *const c_void,
    get_IsImport: *const c_void,
    get_IsSerializable: *const c_void,
    get_IsAnsiClass: *const c_void,
    get_IsUnicodeClass: *const c_void,
    get_IsArray: *const c_void,
    get_IsByRef: *const c_void,
    get_IsPointer: *const c_void,
    get_IsPrimitive: *const c_void,
    get_IsCOMObject: *const c_void,
    get_HasElementType: *const c_void,
    get_IsContextful: *const c_void,
    get_IsMarshalByRef: *const c_void,
    Equals_2: *const c_void,
}
