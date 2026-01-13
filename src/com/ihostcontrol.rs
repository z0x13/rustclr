use core::{ffi::c_void, ops::Deref};
use windows::core::{GUID, HRESULT, Interface};

/// This struct represents the COM `IHostControl` interface.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct IHostControl(windows::core::IUnknown);

/// Trait representing the implementation of the `IHostControl` interface.
pub trait IHostControl_Impl: windows::core::IUnknownImpl {
    /// Requests a host-provided manager object that implements the interface specified by `riid`.
    fn GetHostManager(
        &self,
        riid: *const GUID,
        ppobject: *mut *mut c_void,
    ) -> windows::core::Result<()>;

    /// Notifies the host that the CLR has created an `AppDomainManager` for a new AppDomain.
    fn SetAppDomainManager(
        &self,
        dwappdomainid: u32,
        punkappdomainmanager: windows::core::Ref<'_, windows::core::IUnknown>,
    ) -> windows::core::Result<()>;
}

impl IHostControl_Vtbl {
    /// Creates a new virtual table for the `IHostControl` implementation.
    pub const fn new<Identity: IHostControl_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn GetHostManager<
            Identity: IHostControl_Impl,
            const OFFSET: isize,
        >(
            this: *mut c_void,
            riid: *const GUID,
            ppobject: *mut *mut c_void,
        ) -> HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IHostControl_Impl::GetHostManager(
                    this,
                    core::mem::transmute_copy(&riid),
                    core::mem::transmute_copy(&ppobject),
                )
                .into()
            }
        }

        unsafe extern "system" fn SetAppDomainManager<
            Identity: IHostControl_Impl,
            const OFFSET: isize,
        >(
            this: *mut c_void,
            dwappdomainid: u32,
            punkappdomainmanager: *mut c_void,
        ) -> HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IHostControl_Impl::SetAppDomainManager(
                    this,
                    core::mem::transmute_copy(&dwappdomainid),
                    core::mem::transmute_copy(&punkappdomainmanager),
                )
                .into()
            }
        }

        Self {
            base__: windows::core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            GetHostManager: GetHostManager::<Identity, OFFSET>,
            SetAppDomainManager: SetAppDomainManager::<Identity, OFFSET>,
        }
    }

    /// Verifies if a given interface ID matches `IHostControl`.
    pub fn matches(iid: &windows::core::GUID) -> bool {
        iid == &<IHostControl as windows::core::Interface>::IID
    }
}

impl windows::core::RuntimeName for IHostControl {}

unsafe impl Interface for IHostControl {
    type Vtable = IHostControl_Vtbl;
    const IID: GUID = GUID::from_u128(0x02ca073c_7079_4860_880a_c2f7a449c991);
}

impl Deref for IHostControl {
    type Target = windows::core::IUnknown;

    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

#[repr(C)]
pub struct IHostControl_Vtbl {
    pub base__: windows::core::IUnknown_Vtbl,
    pub GetHostManager: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const GUID,
        ppobject: *mut *mut c_void,
    ) -> HRESULT,
    pub SetAppDomainManager: unsafe extern "system" fn(
        this: *mut c_void,
        dwappdomainid: u32,
        punkappdomainmanager: *mut c_void,
    ) -> HRESULT,
}
