//! Functions that are exported from the xll and invoked by Excel
//! The only two essential functions are xlAutoOpen and xlAutoFree12.

use crate::xlcall::LPXLOPER12;
use crate::variant::Variant;

// pub extern "stdcall" fn xlAutoOpen() implemented in lib.rs as it calls the 
// registration of all used defined functions

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoFree12(px_free: LPXLOPER12) {
    // Rebuild the Box<Variant> so Variant::drop runs and frees string/array memory
    // let _ = unsafe { Box::<Variant>::from_raw(px_free.cast()) };
    let _ = unsafe { Box::<Variant>::from_raw(px_free.cast()) };
}

/// Excel exit point - called when Excel unloads the add-in
#[unsafe(no_mangle)]
pub extern "system" fn xlAutoClose() -> i32 {
    1 // Success
}
