use xladd_core::Reg;

mod actuarial;
mod excel_wrappers;
// mod excel_wrappers_copy;

pub use actuarial::*;

// use log::*;

#[unsafe(no_mangle)]
pub extern "stdcall" fn xlAutoOpen() -> i32 {
    let reg = Reg::new();
    reg.register_all_functions();  // Automatically finds and registers all #[xl_func] functions
    1
}


// xlAutoClose and xlAutoFree12 are defined in xlauto.rs - don't duplicate them here