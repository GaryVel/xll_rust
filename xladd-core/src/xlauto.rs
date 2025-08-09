//! Functions that are exported from the xll and invoked by Excel
//! The only two essential functions are xlAutoOpen and xlAutoFree12.

use crate::xlcall::LPXLOPER12;

// pub extern "stdcall" fn xlAutoOpen() implemented in lib.rs as it calls the 
// registration of all used defined functions

#[unsafe(no_mangle)]
pub extern "stdcall" fn xlAutoFree12(px_free: LPXLOPER12) {
    // take ownership of this xloper. Then when our xloper goes
    // out of scope, its drop method will free any resources.
    let _ = unsafe { Box::from_raw(px_free) };
}

/// Excel exit point - called when Excel unloads the add-in
#[unsafe(no_mangle)]
pub extern "stdcall" fn xlAutoClose() -> i32 {
    1 // Success
}
