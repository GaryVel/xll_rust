//! Entry point code for xladd-core, based on the sample C++ code
//! supplied with the Microsoft Excel12 SDK

use crate::registrator::debug_print;
use crate::variant::Variant;
use crate::xlcall::{xlFree, xlretFailed, LPXLOPER12, XLOPER12};
use libc::c_int;

use std::{mem, ptr, sync::OnceLock};
use windows::{
    core::{s, w, PCWSTR},
    Win32::System::LibraryLoader::{GetModuleHandleW, GetProcAddress},
};

type EXCEL12PROC = extern "system" fn(
    xlfn: c_int,
    count: c_int,
    rgpxloper12: *const LPXLOPER12,
    xloper12res: LPXLOPER12,
) -> c_int;

// Excel12 entry point thunk provider (returns the Excel12 pointer as usize).
type FNGETEXCEL12ENTRYPT = unsafe extern "C" fn() -> usize;

/// Cached Excel12 function pointer (usize).
static PEXCEL12: OnceLock<usize> = OnceLock::new();

/// Resolve and cache the Excel12 entry point.
/// Returns the Excel12 function pointer (0 if not found).
fn fetch_excel12_entry_pt() -> usize {
    *PEXCEL12.get_or_init(|| unsafe {
        let mut pexcel12: usize = 0;

        // Try XLCALL32.DLL first
        if let Ok(hmod) = GetModuleHandleW(w!("XLCALL32.DLL")) {
            if let Some(proc_addr) = GetProcAddress(hmod, s!("MdCallBack12")) {
                let get_excel12: FNGETEXCEL12ENTRYPT = mem::transmute(proc_addr);
                let entry_pt = get_excel12();
                if entry_pt != 0 {
                    pexcel12 = entry_pt;
                }
            }
        }

        // Fallback: resolve Excel12 directly from the current process
        if pexcel12 == 0 {
            if let Ok(hmod) = GetModuleHandleW(PCWSTR::null()) {
                if let Some(proc_addr) = GetProcAddress(hmod, s!("Excel12")) {
                    pexcel12 = proc_addr as usize;
                }
            }
        }

        pexcel12
    })
}

/// Call into Excel, passing a function number as defined in xlcall and a slice
/// of Variant, and returning a Variant. To find out the number and type of
/// parameters and the expected result, please consult the Excel SDK documentation.
pub fn excel12(xlfn: u32, opers: &mut [Variant]) -> Variant {
    debug_print(&format!("FuncID:{}, {} args)", xlfn, opers.len()));
    let mut args: Vec<LPXLOPER12> = Vec::with_capacity(opers.len());
    for oper in opers.iter_mut() {
        debug_print(&format!("arg: {}", oper));
        args.push(oper.as_mut_xloper());
    }
    let mut result = Variant::default();
    let res = excel12v(xlfn as i32, result.as_mut_xloper(), &args);
    match res {
        0 => result,
        v => {
            debug_print(&format!("ReturnCode {}", v));
            result
        }
    }
}

pub fn excel12_1(xlfn: u32, mut oper: Variant) -> Variant {
    let mut result = Variant::default();
    excel12v(xlfn as i32, result.as_mut_xloper(), &[oper.as_mut_xloper()]);
    result
}

pub fn excel12v(xlfn: i32, oper_res: &mut XLOPER12, opers: &[LPXLOPER12]) -> i32 {
    let pexcel12 = fetch_excel12_entry_pt();

    unsafe {
        if pexcel12 == 0 {
            xlretFailed as i32
        } else {
            let p = opers.as_ptr();
            let len = opers.len();
            let f: EXCEL12PROC = mem::transmute(pexcel12);
            f(xlfn, len as i32, p, oper_res)
        }
    }
}

pub fn excel_free(xloper: LPXLOPER12) -> i32 {
    let pexcel12 = fetch_excel12_entry_pt();

    unsafe {
        if pexcel12 == 0 {
            xlretFailed as i32
        } else {
            let f: EXCEL12PROC = mem::transmute(pexcel12);
            f(xlFree as i32, 1, &xloper, ptr::null_mut())
        }
    }
}
