#![allow(non_snake_case, non_camel_case_types, non_upper_case_globals)]

use std::{fmt, mem, slice};
//#[cfg(feature = "try_from")]
use crate::entrypoint::excel_free;
use crate::xlcall::{
    xlbitDLLFree, xlbitXLFree, xlerrDiv0, xlerrGettingData, xlerrNA, xlerrName, xlerrNull,
    xlerrNum, xlerrRef, xlerrValue, Xloper12Value, Xloper12SRef,
    Xloper12Array, xltypeBool, xltypeErr, xltypeInt, xltypeMissing,
    xltypeMulti, xltypeNil, xltypeNum, xltypeRef, xltypeSRef, xltypeStr, xltypeMask,
    LPXLOPER12, XLMREF12, XLOPER12, xloper12, xlref12,
};
use std::convert::TryFrom;
use std::f64;
use thiserror::Error;

// ####################################################################################################################
// 1. CORE TYPES AND CONSTANTS
// ####################################################################################################################

// Custom error types for conversion failures
// Implements std::error::Error for proper error propagation

#[derive(Error, Debug)]
pub enum XLAddError {
    #[error("Could not convert parameter [{0}] to f64")]
    F64ConversionFailed(String),
    
    #[error("Could not convert parameter [{0}] to bool")]
    BoolConversionFailed(String),
    
    #[error("Could not convert parameter [{0}] to integer")]
    IntConversionFailed(String),
    
    #[error("Could not convert parameter [{0}] to string")]
    StringConversionFailed(String),
    
    #[error("Function [{func}] is missing parameter {param}")]
    MissingArgument { func: String, param: String },
    
    #[error("Invalid data: {0}")]
    InvalidData(String),
    
    #[error("Array dimension error: {0}")]
    DimensionError(String),
}

const xltypeStr_xlbitDLLFree: u32 = xltypeStr | xlbitDLLFree;
const xltypeMulti_xlbitDLLFree: u32 = xltypeMulti | xlbitDLLFree;

/// Variant is a wrapper around a Excel's XLOPER12 union type. It can contain a string, i32
/// or f64, or a two dimensional of any mixture of these.
pub struct Variant(XLOPER12);

// --------------------------------------------------------------------------------------------------------------------
// 2. CORE VARIANT METHODS
// --------------------------------------------------------------------------------------------------------------------

impl Variant {
    /// Construct a variant containing a missing entry. This is used in function calls to
    /// signal that a parameter should be defaulted.
    pub fn missing() -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeMissing,
            val: Xloper12Value { w: 0 },
        })
    }

    pub fn is_missing_or_null(&self) -> bool {
        self.0.xltype & xltypeMask == xltypeMissing || self.0.xltype & xltypeMask == xltypeNil
    }

    /// Construct a variant containing an error. This is used in Excel to represent standard errors
    /// that are shown as #DIV0 etc. Currently supported error codes are:
    /// xlerrNull, xlerrDiv0, xlerrValue, xlerrRef, xlerrName, xlerrNum, xlerrNA, xlerrGettingData
    pub fn from_err(xlerr: u32) -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeErr,
            val: Xloper12Value { err: xlerr as i32 },
        })
    }

    /// Construct a variant containing an array from a slice of other variants. The variants
    /// may contain arrays or scalar strings or numbers, which are treated like single-cell
    /// arrays. They are glued either horizontally (horiz=true) or vertically. If the arrays
    /// do not match sizes in the other dimension, they are padded with blanks.
    pub fn concat(from: &[Variant], horiz: bool) -> Variant {
        // first find the size of the resulting array
        let mut columns: usize = 0;
        let mut rows: usize = 0;
        for xloper in from.iter() {
            let dim = xloper.dim();
            if horiz {
                columns += dim.0;
                rows = rows.max(dim.1);
            } else {
                columns = columns.max(dim.0);
                rows += dim.1;
            }
        }

        // Zero-sized arrays cause Excel to crash. Arrays with a dimension of
        // one (either rows or cols) are confusing to Excel, which repeats them
        // when using array paste. Solve both problems by padding with NA and
        // setting the min rows or cols to two.
        rows = rows.max(2);
        columns = columns.max(2);

        // If the array is too big, return an error string
        if rows > 1_048_576 || columns > 16384 {
            return Self::from("#ERR resulting array is too big");
        }

        // now clone the components into place
        let size = rows * columns;
        let mut array = Vec::with_capacity(size);
        array.resize_with(size, || Variant::from_err(xlerrNA));
        let mut col = 0;
        let mut row = 0;
        for var in from.iter() {
            match var.0.xltype & xltypeMask {
                xltypeMulti => {
                    if let Some(arr) = var.0.val.as_array(var.0.xltype) {
                        let var_cols = arr.columns as usize;
                        let var_rows = arr.rows as usize;
                        for x in 0..var_cols {
                            for y in 0..var_rows {
                                let src_index = y * var_cols + x;
                                let dest = (row + y) * columns + col + x;
                                if let Some(xloper) = arr.get(src_index) {
                                    array[dest] = Variant::from(xloper as *const _ as LPXLOPER12).clone();
                                }
                            }
                        }
                        if horiz {
                            col += var_cols;
                        } else {
                            row += var_rows;
                        }
                    }
                },
                xltypeMissing => {}
                _ => {
                    let dest = row * columns + col;
                    array[dest] = var.clone();
                    if horiz {
                        col += 1;
                    } else {
                        row += 1;
                    }
                }
            }
        }

        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);

        Variant(XLOPER12 {
            xltype: xltypeMulti,
            val: Xloper12Value {
                array: Xloper12Array {
                    lparray,
                    rows: rows as i32,
                    columns: columns as i32,
                },
            },
        })
    }

    /// Creates a transposed clone of this Variant. If this Variant is a scalar type,
    /// simply returns it unchanged.
    pub fn transpose(&self) -> Variant {
        // simply clone any scalar type, including errors
        if (self.0.xltype & xltypeMask) != xltypeMulti {
            return self.clone();
        }

        // We have an array that we need to transpose. Create a vector of
        // Variant to contain the elements.
        let dim = self.dim();
        if dim.0 > 1_048_576 || dim.1 > 16384 {
            return Self::from("#ERR resulting array is too big");
        }

        let len = dim.0 * dim.1;
        let mut array = Vec::with_capacity(len);

        // Copy the elements transposed, cloning each one
        for i in 0..dim.1 {
            for j in 0..dim.0 {
                array.push(self.at(j, i));
            }
        }

        // Return as a Variant
        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);

        Variant(XLOPER12 {
            xltype: xltypeMulti,
            val: Xloper12Value {
                array: Xloper12Array {
                    lparray,
                    rows: dim.0 as i32,
                    columns: dim.1 as i32,
                },
            },
        })
    }

    /// Exposes the underlying XLOPER12
    pub fn as_mut_xloper(&mut self) -> &mut XLOPER12 {
        &mut self.0
    }

    /// Gets the count of rows and columns. Scalars are treated as 1x1. Missing values are
    /// treated as 0x0.
    pub fn dim(&self) -> (usize, usize) {
        match self.0.xltype & xltypeMask {
            xltypeMulti => {
                self.0.val.as_array(self.0.xltype)
                    .map(|arr| arr.dim())
                    .unwrap_or((0, 0))
            },
            xltypeSRef => {
                self.0.val.as_sref(self.0.xltype)
                    .map(|sref| sref.dim())
                    .unwrap_or((0, 0))
            },
            xltypeRef => {
                self.0.val.as_mref(self.0.xltype)
                    .and_then(|mref| get_mref_dim_safe(mref.lpmref))
                    .unwrap_or((0, 0))
            },
            xltypeMissing => (0, 0),
            _ => (1, 1),
        }
    }

    /// Gets the element at the given column and row. If this is a scalar, treat it as a one-element
    /// array. If the column or row is out of bounds, return NA. The returned element is always cloned
    /// so it can be returned as a value
    pub fn at(&self, column: usize, row: usize) -> Variant {
        if (self.0.xltype & xltypeMask) != xltypeMulti {
            if column == 0 && row == 0 {
                self.clone()
            } else {
                Self::from_err(xlerrNA)
            }
        } else if let Some(array) = self.0.val.as_array(self.0.xltype) {
            array.get_2d(row, column)
                .map(|xloper| Variant::from(xloper as *const _ as LPXLOPER12).clone())
                .unwrap_or_else(|| Self::from_err(xlerrNA))
        } else {
            Self::from_err(xlerrNA)
        }
    }

    pub fn location(&self) -> (i32, i32) {
        self.0.val.as_sref(self.0.xltype)
            .map(|sref| (sref.ref_.rwFirst, sref.ref_.colFirst))
            .unwrap_or((0, 0))
    }

    pub fn as_sref(rowStart: i32, rowEnd: i32, colStart: i32, colEnd: i32) -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeSRef,
            val: Xloper12Value {
                sref: Xloper12SRef {
                    count: 1,
                    ref_: xlref12 {
                        rwFirst: rowStart,
                        rwLast: rowEnd,
                        colFirst: colStart,
                        colLast: colEnd,
                    },
                },
            },
        })
    }

    pub fn is_ref(&self) -> bool {
        let xltype = self.0.xltype & xltypeMask;
        xltype == xltypeRef || xltype == xltypeSRef
    }
}

/// Construct a variant containing nil. This is used in Excel to represent cells that have
/// nothing in them. It is also a sensible starting state for an uninitialized variant.
impl Default for Variant {
    fn default() -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeNil,
            val: Xloper12Value { w: 0 },
        })
    }
}

/// We need to implement Drop, as Variant is a wrapper around a union type that does
/// not know how to handle its contained pointers.
impl Drop for Variant {
    fn drop(&mut self) {
        if (self.0.xltype & xlbitXLFree) != 0 {
            excel_free(&mut self.0);
            return;
        }

        match self.0.xltype {
            xltypeStr_xlbitDLLFree => {
                // We have a 16bit string that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so its drop method
                // will clean up the memory for us.
                if let Some(ptr) = self.0.val.as_str_ptr(self.0.xltype) {
                    unsafe {
                        let len = *ptr as usize + 1;
                        let cap = len;
                        Vec::from_raw_parts(ptr, len, cap);
                    }
                }
            }
            xltypeMulti_xlbitDLLFree => {
                // We have an array that was originally allocated as a vector of
                // Variant but then forgotten. Reconstruct the vector, so its drop method
                // will clean up the vector and its elements for us.
                if let Some(array) = self.0.val.as_array(self.0.xltype) {
                    unsafe {
                        let p = array.lparray as *mut Variant;
                        let len = (array.rows * array.columns) as usize;
                        let cap = len;
                        Vec::from_raw_parts(p, len, cap);
                    }
                }
            }
            _ => {
                // nothing to do
            }
        }
    }
}

/// We need to hand-code Clone, because of the ownership issues for strings and multi.
impl Clone for Variant {
    fn clone(&self) -> Variant {
        // a simple copy is good enough for most variant types, but make sure the addin
        // is the owner
        let mut copy = Variant(self.0);
        copy.0.xltype &= !xlbitXLFree;
        copy.0.xltype |= xlbitDLLFree;

        // Special handling for string and mult, to avoid double delete of the member
        match copy.0.xltype {
            xltypeStr_xlbitDLLFree => {
                // We have a 16bit string that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so we can clone it.
                if let Some(ptr) = copy.0.val.as_str_ptr(copy.0.xltype) {
                    unsafe {
                        let len = *ptr as usize + 1;
                        let cap = len;
                        let string_vec = Vec::from_raw_parts(ptr, len, cap);
                        let mut cloned = string_vec.clone();
                        copy.0.val.str = cloned.as_mut_ptr();

                        // now forget everything -- we do not want either string deallocated
                        mem::forget(string_vec);
                        mem::forget(cloned);
                    }
                }
            }
            xltypeMulti_xlbitDLLFree => {
                // We have an array that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so we can clone it.
                if let Some(array) = self.0.val.as_array(self.0.xltype) {
                    unsafe {
                        let p = array.lparray as *mut Variant;
                        let len = (array.rows * array.columns) as usize;
                        let cap = len;
                        let array_vec = Vec::from_raw_parts(p, len, cap);
                        let mut cloned = array_vec.clone();
                        copy.0.val.array.lparray = cloned.as_mut_ptr() as LPXLOPER12;

                        // now forget everything -- we do not want either string deallocated
                        mem::forget(array_vec);
                        mem::forget(cloned);
                    }
                }
            }
            _ => {
                // nothing to do
            }
        }

        copy
    }
}

/// Implement Display, which means we do not need a method for converting to strings. Just use
/// to_string.
impl fmt::Display for Variant {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self.0.xltype & xltypeMask {
            xltypeErr => {
                self.0.val.as_err(self.0.xltype)
                    .map(|err| match err as u32 {
                        xlerrNull => write!(f, "#NULL"),
                        xlerrDiv0 => write!(f, "#DIV0"),
                        xlerrValue => write!(f, "#VALUE"),
                        xlerrRef => write!(f, "#REF"),
                        xlerrName => write!(f, "#NAME"),
                        xlerrNum => write!(f, "#NUM"),
                        xlerrNA => write!(f, "#NA"),
                        xlerrGettingData => write!(f, "#DATA"),
                        v => write!(f, "#BAD_ERR {}", v),
                    })
                    .unwrap_or_else(|| write!(f, "#ERR"))
            },
            xltypeInt => {
                self.0.val.as_int(self.0.xltype)
                    .map(|i| write!(f, "{}", i))
                    .unwrap_or_else(|| write!(f, "#INT_ERR"))
            },
            xltypeMissing => write!(f, "#MISSING"),
            xltypeMulti => write!(f, "#MULTI"),
            xltypeNil => write!(f, "#NIL"),
            xltypeNum => {
                self.0.val.as_num(self.0.xltype)
                    .map(|n| write!(f, "{}", n))
                    .unwrap_or_else(|| write!(f, "#NUM_ERR"))
            },
            xltypeStr => write!(f, "{}", String::try_from(&self.clone()).unwrap()),
            xlerrNull => write!(f, "#NULL"),
            xltypeSRef => {
                self.0.val.as_sref(self.0.xltype)
                    .map(|sref| write!(f, "Sref:({},{}) -> ({},{})",
                        sref.ref_.rwFirst, sref.ref_.colFirst,
                        sref.ref_.rwLast, sref.ref_.colLast))
                    .unwrap_or_else(|| write!(f, "#SREF_ERR"))
            },
            v => write!(f, "#BAD_ERR {}", v),
        }
    }
}

impl fmt::Debug for Variant {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self.0.xltype & xltypeMask {
            xltypeErr => {
                self.0.val.as_err(self.0.xltype)
                    .map(|err| match err as u32 {
                        xlerrNull => write!(f, "#NULL"),
                        xlerrDiv0 => write!(f, "#DIV0"),
                        xlerrValue => write!(f, "#VALUE"),
                        xlerrRef => write!(f, "#REF"),
                        xlerrName => write!(f, "#NAME"),
                        xlerrNum => write!(f, "#NUM"),
                        xlerrNA => write!(f, "#NA"),
                        xlerrGettingData => write!(f, "#DATA"),
                        v => write!(f, "#BAD_ERR {}", v),
                    })
                    .unwrap_or_else(|| write!(f, "#ERR"))
            },
            xlerrNull => write!(f, "#NULL"),
            xltypeInt => {
                self.0.val.as_int(self.0.xltype)
                    .map(|i| write!(f, "{}", i))
                    .unwrap_or_else(|| write!(f, "#INT_ERR"))
            },
            xltypeMissing => write!(f, "#MISSING"),
            xltypeMulti => write!(f, "#MULTI"),
            xltypeNil => write!(f, "#NIL"),
            xltypeBool => {
                self.0.val.as_bool(self.0.xltype)
                    .map(|b| write!(f, "{}", b))
                    .unwrap_or_else(|| write!(f, "#BOOL_ERR"))
            },
            xltypeNum => {
                self.0.val.as_num(self.0.xltype)
                    .map(|n| write!(f, "{}", n))
                    .unwrap_or_else(|| write!(f, "#NUM_ERR"))
            },
            xltypeStr => write!(f, "{}", String::try_from(&self.clone()).unwrap()),
            v => write!(f, "#BAD_XLOPER {}", v),
        }
    }
}


// --------------------------------------------------------------------------------------------------------------------
// 3. FROM EXCEL TO RUST (Excel -> Rust conversions)
// --------------------------------------------------------------------------------------------------------------------

/// Construct a variant from an LPXLOPER12, for example supplied by Excel. The assumption
/// is that Excel continues to own the XLOPER12 and its lifetime is greater than that of
/// the Variant we construct here. For example, the LPXLOPER may be an argument to one
/// of our functions. We therefore do not want to own any of the data in this variant, so
/// we clear all ownership bits. This means we treat it as a kind of dynamic mut ref.
impl From<LPXLOPER12> for Variant {
    fn from(xloper: LPXLOPER12) -> Variant {
        let mut result = Variant(unsafe { *xloper });
        result.0.xltype &= xltypeMask; // no ownership bits
        result
    }
}

// For async functions
#[derive(Debug)]
pub struct XLOPERPtr(pub *mut xloper12);
unsafe impl Send for XLOPERPtr {}

impl From<XLOPERPtr> for Variant {
    fn from(xloper: XLOPERPtr) -> Variant {
        Variant(unsafe { std::mem::transmute::<XLOPER12, xloper12>(*xloper.0) })
    }
}

// --------------------------------------------------------------------------------------------------------------------
// 4. BASIC TYPE CONVERSIONS (Excel types -> Rust primitives)
// --------------------------------------------------------------------------------------------------------------------

// From xloper12 (infallible)

impl From<&xloper12> for String {
    fn from(v: &xloper12) -> String {
        match v.xltype & xltypeMask {
            xltypeNum => v.val.as_num(v.xltype)
                .map(|n| n.to_string())
                .unwrap_or_default(),
            xltypeStr => {
                v.val.as_str_ptr(v.xltype)
                    .and_then(|ptr| unsafe {
                        let cstr_len = *ptr as usize;
                        let cstr_slice = slice::from_raw_parts(ptr.offset(1), cstr_len);
                        String::from_utf16(cstr_slice).ok()
                    })
                    .unwrap_or_default()
            },
            xltypeMulti => {
                v.val.as_array(v.xltype)
                    .and_then(|arr| arr.get(0))
                    .map(|first| String::from(first))
                    .unwrap_or_default()
            },
            xltypeBool => v.val.as_bool(v.xltype)
                .map(|b| b.to_string())
                .unwrap_or_default(),
            _ => String::new(),
        }
    }
}

impl From<&xloper12> for f64 {
    fn from(v: &xloper12) -> f64 {
        match v.xltype & xltypeMask {
            xltypeNum => v.val.as_num(v.xltype).unwrap_or(f64::NAN),
            xltypeInt => v.val.as_int(v.xltype).map(|i| i as f64).unwrap_or(f64::NAN),
            xltypeStr => f64::NAN,
            xltypeBool => v.val.as_bool(v.xltype).map(|b| if b { 1.0 } else { 0.0 }).unwrap_or(f64::NAN),
            xltypeMulti => {
                v.val.as_array(v.xltype)
                    .and_then(|arr| arr.get(0))
                    .map(|first| f64::from(first))
                    .unwrap_or(f64::NAN)
            },
            _ => f64::NAN,
        }
    }
}

impl From<&xloper12> for i32 {
    fn from(v: &xloper12) -> i32 {
        match v.xltype & xltypeMask {
            xltypeNum => v.val.as_num(v.xltype).map(|n| n as i32).unwrap_or(0),
            xltypeInt => v.val.as_int(v.xltype).unwrap_or(0),
            xltypeStr => 0,
            xltypeBool => v.val.as_bool(v.xltype).map(|b| if b { 1 } else { 0 }).unwrap_or(0),
            xltypeMulti => {
                v.val.as_array(v.xltype)
                    .and_then(|arr| arr.get(0))
                    .map(|first| i32::from(first))
                    .unwrap_or(0)
            },
            _ => 0,
        }
    }
}

impl From<&xloper12> for bool {
    fn from(v: &xloper12) -> bool {
        match v.xltype & xltypeMask {
            xltypeNum => v.val.as_num(v.xltype).map(|n| n != 0.0).unwrap_or(false),
            xltypeStr => false,
            xltypeBool => v.val.as_bool(v.xltype).unwrap_or(false),
            xltypeMulti => {
                v.val.as_array(v.xltype)
                    .and_then(|arr| arr.get(0))
                    .map(|first| bool::from(first))
                    .unwrap_or(false)
            },
            _ => false,
        }
    }
}

impl<'a> From<&'a xloper12> for Vec<f64> {
    fn from(xloper: &'a xloper12) -> Vec<f64> {
        let result = Variant(*xloper);
        Vec::<f64>::try_from(&Variant::from(result)).unwrap_or_default()
    }
}

// From Variant (infallible) 

impl From<&Variant> for String {
    fn from(v: &Variant) -> String {
        String::from(&v.0)
    }
}

impl From<&Variant> for i32 {
    fn from(v: &Variant) -> i32 {
        // Delegate to the f64 TryFrom, then convert
        f64::try_from(v)
            .unwrap_or(0.0) // Default to 0 for invalid values
            as i32
    }
}

impl From<&Variant> for u32 {
    fn from(v: &Variant) -> u32 {
        // Delegate to the f64 TryFrom, then convert
        f64::try_from(v)
            .unwrap_or(0.0) // Default to 0 for invalid values
            .max(0.0) // Ensure non-negative
            as u32
    }
}

/// Converts a variant into a f64 array filling the missing or invalid with f64::NAN.
/// This is so that you can handle those appropriately for your application (for example fill with the mean value or 0)
impl<'a> TryFrom<&'a Variant> for Vec<f64> {
    type Error = XLAddError;
    
    fn try_from(v: &'a Variant) -> Result<Vec<f64>, Self::Error> {
        let (cols, rows) = v.dim();
        let mut res = Vec::with_capacity(cols * rows);
        
        if cols == 1 && rows == 1 {
            res.push(f64::try_from(v)?);
        } else if let Some(array) = v.0.val.as_array(v.0.xltype) {
            for j in 0..rows {
                for i in 0..cols {
                    let index = j * cols + i;
                    if let Some(xloper) = array.get(index) {
                        let val = match xloper.xltype & xltypeMask {
                            xltypeNum => xloper.val.as_num(xloper.xltype)
                                .ok_or_else(|| XLAddError::F64ConversionFailed(
                                    format!("Failed to get number at [{}, {}]", i, j)
                                ))?,
                            xltypeInt => xloper.val.as_int(xloper.xltype)
                                .map(|i| i as f64)
                                .ok_or_else(|| XLAddError::F64ConversionFailed(
                                    format!("Failed to get integer at [{}, {}]", i, j)
                                ))?,
                            _ => return Err(XLAddError::F64ConversionFailed(
                                format!("Invalid type at [{}, {}]", i, j)
                            ))
                        };
                        res.push(val);
                    } else {
                        return Err(XLAddError::F64ConversionFailed(
                            format!("Failed to access element at [{}, {}]", i, j)
                        ));
                    }
                }
            }
        } else {
            return Err(XLAddError::F64ConversionFailed("Not an array".to_string()));
        }
        Ok(res)
    }
}

/// Converts a variant into a string array filling the missing or invalid with f64::NAN.
/// This is so that you can handle those appropriately for your application (for example fill with the mean value or 0)
impl<'a> From<&'a Variant> for Vec<String> {
    fn from(v: &'a Variant) -> Vec<String> {
        let (x, y) = v.dim();
        let mut res = Vec::with_capacity(x * y);
        if x == 1 && y == 1 {
            res.push(String::from(v));
        } else if let Some(array) = v.0.val.as_array(v.0.xltype) {
            for j in 0..y {
                for i in 0..x {
                    let index = j * x + i;
                    if let Some(xloper) = array.get(index) {
                        res.push(String::from(xloper));
                    } else {
                        res.push(String::new());
                    }
                }
            }
        }
        res
    }
}

// --------------------------------------------------------------------------------------------------------------------
// 5. FALLIBLE CONVERSIONS (TryFrom - can fail)
// --------------------------------------------------------------------------------------------------------------------

// GV: additional conversion traits for Variant
// Support for converting Variant to u32 and i32
// NB: this is not necessary as Excel passes all numbers (integers, decimals) as double (f64), and
//     so using u32 for parameter in function called by Excel is not recommended. Rather cast as
//     u32 when calling the underlying function.

impl TryFrom<&Variant> for bool {
    type Error = XLAddError;
    
    fn try_from(v: &Variant) -> Result<Self, Self::Error> {
        match v.0.xltype & xltypeMask {
            xltypeBool => v.0.val.as_bool(v.0.xltype)
                .ok_or_else(|| XLAddError::BoolConversionFailed("Failed to extract boolean".to_string())),
            xltypeNum => {
                v.0.val.as_num(v.0.xltype)
                    .and_then(|num| {
                        if num == 0.0 {
                            Some(false)
                        } else if num == 1.0 {
                            Some(true)
                        } else {
                            None
                        }
                    })
                    .ok_or_else(|| XLAddError::BoolConversionFailed(
                        "Number must be 0.0 or 1.0 for boolean conversion".to_string()
                    ))
            }
            xltypeInt => {
                v.0.val.as_int(v.0.xltype)
                    .and_then(|int_val| {
                        if int_val == 0 {
                            Some(false)
                        } else if int_val == 1 {
                            Some(true)
                        } else {
                            None
                        }
                    })
                    .ok_or_else(|| XLAddError::BoolConversionFailed(
                        "Integer must be 0 or 1 for boolean conversion".to_string()
                    ))
            }
            xltypeStr => {
                let str_val = String::from(&v.0);
                match str_val.to_lowercase().as_str() {
                    "true" | "yes" | "1" => Ok(true),
                    "false" | "no" | "0" => Ok(false),
                    _ => Err(XLAddError::BoolConversionFailed(format!(
                        "String '{}' is not a valid boolean", str_val
                    )))
                }
            }
            xltypeErr => Err(XLAddError::BoolConversionFailed("Cannot convert Excel error to boolean".to_string())),
            xltypeMissing | xltypeNil => Err(XLAddError::BoolConversionFailed("Cannot convert missing/nil value to boolean".to_string())),
            _ => Err(XLAddError::BoolConversionFailed("Invalid Excel type for boolean conversion".to_string()))
        }
    }
}

// Removed this:
// impl From<&Variant> for f64 { ... }
// Added this instead:
impl TryFrom<&Variant> for f64 {
    type Error = XLAddError;
    
    fn try_from(v: &Variant) -> Result<Self, Self::Error> {
        match v.0.xltype & xltypeMask {
            xltypeNum => v.0.val.as_num(v.0.xltype)
                .ok_or_else(|| XLAddError::F64ConversionFailed("Failed to extract number".to_string())),
            xltypeInt => v.0.val.as_int(v.0.xltype)
                .map(|i| i as f64)
                .ok_or_else(|| XLAddError::F64ConversionFailed("Failed to extract integer".to_string())),
            xltypeStr => {
                let str_val = String::from(&v.0);
                str_val.parse::<f64>().map_err(|_| {
                    XLAddError::F64ConversionFailed(format!("Cannot convert '{}' to number", str_val))
                })
            }
            xltypeBool => {
                v.0.val.as_bool(v.0.xltype)
                    .map(|b| if b { 1.0 } else { 0.0 })
                    .ok_or_else(|| XLAddError::F64ConversionFailed("Failed to extract boolean".to_string()))
            }
            xltypeErr => Err(XLAddError::F64ConversionFailed("Cannot convert Excel error to number".to_string())),
            xltypeMissing | xltypeNil => Err(XLAddError::F64ConversionFailed("Missing or empty value - number required".to_string())),
            _ => Err(XLAddError::F64ConversionFailed("Invalid Excel type for number conversion".to_string()))
        }
    }
}

// --------------------------------------------------------------------------------------------------------------------
// 6. RUST TO EXCEL CONVERSIONS (Rust -> Excel)
// --------------------------------------------------------------------------------------------------------------------

// Convert i32 to Variant via f64 (Excel's native number type)
impl From<i32> for Variant {
    fn from(val: i32) -> Variant {
        Variant::from(val as f64)
    }
}

// Convert u32 to Variant via f64 (Excel's native number type)  
impl From<u32> for Variant {
    fn from(val: u32) -> Variant {
        Variant::from(val as f64)
    }
}

/// Construct a variant containing an bool (i32)
impl From<bool> for Variant {
    fn from(xbool: bool) -> Variant {
        Variant(XLOPER12 {
            xltype: xltypeBool,
            val: Xloper12Value {
                xbool: xbool as i32,
            },
        })
    }
}

/// Construct a variant containing an float (f64)
impl From<f64> for Variant {
    fn from(num: f64) -> Variant {
        match num {
            num if num.is_nan() => Variant::from_err(xlerrNA),
            num if num.is_infinite() => Variant::from_err(xlerrNA),
            num => Variant(XLOPER12 {
                xltype: xltypeNum,
                val: Xloper12Value { num },
            }),
        }
    }
}

/// Construct a variant containing a string. Strings in Excel (at least after Excel 97) are 16bit
/// Unicode starting with a 16-bit length. The length is treated as signed, which means that
/// strings can be no longer than 32k characters. If a string longer than this is supplied, or a
/// string that is not valid 16bit Unicode, an xlerrValue error is stored instead.
impl From<&str> for Variant {
    fn from(s: &str) -> Variant {
        let mut wstr: Vec<u16> = s.encode_utf16().collect();
        if wstr.len() > 65534 {
            return Variant::from_err(xlerrValue);
        }
        // Pascal-style string with length at the start. Forget the string so we do not delete it.
        // We are now relying on the drop method of Variant to clean it up for us. Note that the
        // shrink_to_fit is essential, so the capacity is the same as the length. We have no way
        // of storing the capacity otherwise.
        wstr.insert(0, wstr.len() as u16);
        wstr.shrink_to_fit();
        let p = wstr.as_mut_ptr();
        mem::forget(wstr);
        Variant(XLOPER12 {
            xltype: xltypeStr | xlbitDLLFree,
            val: Xloper12Value { str: p },
        })
    }
}

impl From<String> for Variant {
    fn from(s: String) -> Variant {
        Variant::from(s.as_str())
    }
}

impl From<&String> for Variant {
    fn from(s: &String) -> Variant {
        Variant::from(s.as_str())
    }
}

// --------------------------------------------------------------------------------------------------------------------
// 7. ARRAY CONVERSIONS
// --------------------------------------------------------------------------------------------------------------------

// Simple arrays

// converting two dimensional array of f64 to Variant
impl From<Vec<Vec<f64>>> for Variant {
    fn from(arr: Vec<Vec<f64>>) -> Variant {
        if arr.is_empty() || arr[0].is_empty() {
            return Variant::from_err(xlerrNull);
        }
        
        let rows = arr.len();
        let cols = arr[0].len();
        let mut flat_variants = Vec::with_capacity(rows * cols);
        
        // Flatten row by row
        for row in arr {
            for val in row {
                flat_variants.push(Variant::from(val));
            }
        }
        
        let lparray = flat_variants.as_mut_ptr() as LPXLOPER12;
        mem::forget(flat_variants);
        
        Variant(XLOPER12 {
            xltype: xltypeMulti | xlbitDLLFree,
            val: Xloper12Value {
                array: Xloper12Array {
                    lparray,
                    rows: rows as i32,
                    columns: cols as i32,
                },
            },
        })
    }
}


// converting one dimensional array of f64 to Variant
impl From<Vec<f64>> for Variant {
    fn from(arr: Vec<f64>) -> Variant {
        if arr.is_empty() {
            return Variant::from_err(xlerrNull);
        }
        
        if arr.len() == 1 {
            // Single value - return as scalar
            return Variant::from(arr[0]);
        }
        
        // Create as a horizontal array (1 row, n columns)
        let mut variants = arr.into_iter().map(|v| Variant::from(v)).collect::<Vec<_>>();
        let lparray = variants.as_mut_ptr() as LPXLOPER12;
        let columns = variants.len();
        mem::forget(variants);
        
        Variant(XLOPER12 {
            xltype: xltypeMulti | xlbitDLLFree,
            val: Xloper12Value {
                array: Xloper12Array {
                    lparray,
                    rows: 1,
                    columns: std::cmp::min(16383, columns as i32),
                },
            },
        })
    }
}

impl From<Vec<&str>> for Variant {
    fn from(arr: Vec<&str>) -> Variant {
        let mut array = Vec::new();
        arr.iter().for_each(|&v| {
            array.push(Variant::from(v));
        });

        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);
        let rows = 1;
        let columns = arr.len();
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else {
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: Xloper12Value {
                    array: Xloper12Array {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

// Complex arrays

// Construct 2d variant array from (string,f64)
impl From<Vec<(String, f64)>> for Variant {
    fn from(arr: Vec<(String, f64)>) -> Variant {
        let mut array = Vec::new();
        arr.iter().for_each(|v| {
            array.push(Variant::from(v.0.as_str()));
            array.push(Variant::from(v.1))
        });

        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);
        let rows = arr.len();
        let columns = 2;
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else {
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: Xloper12Value {
                    array: Xloper12Array {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

impl From<Vec<(Variant, f64)>> for Variant {
    fn from(arr: Vec<(Variant, f64)>) -> Variant {
        let mut array = Vec::new();
        for (var, num) in arr.iter() {
            array.push(var.clone());
            array.push(Variant::from(*num));
        }

        let lparray = array.as_mut_ptr() as LPXLOPER12;
        mem::forget(array);
        let rows = arr.len();
        let columns = 2;
        if rows == 0 || columns == 0 {
            Variant::from_err(xlerrNull)
        } else {
            Variant(XLOPER12 {
                xltype: xltypeMulti | xlbitDLLFree,
                val: Xloper12Value {
                    array: Xloper12Array {
                        lparray,
                        rows: std::cmp::min(1_048_575, rows as i32),
                        columns: std::cmp::min(16383, columns as i32),
                    },
                },
            })
        }
    }
}

// --------------------------------------------------------------------------------------------------------------------
// 8. EXCEL SPECIFIC CONVERSIONS
// --------------------------------------------------------------------------------------------------------------------

/// Construct a LPXlOPER12 from a Variant. This is just a cast to the underlying union contained within a pointer
///  that we pass back to Excel. Excel will clean up the pointer after us
impl From<Variant> for LPXLOPER12 {
    fn from(v: Variant) -> LPXLOPER12 {
        Box::into_raw(Box::new(v)) as LPXLOPER12
    }
}

// --------------------------------------------------------------------------------------------------------------------
// 9. UTILITY FUNCTIONS
// --------------------------------------------------------------------------------------------------------------------

// Gets the array size of a multi-cell reference. If the reference is badly formed, returns None
fn get_mref_dim_safe(mref: *const XLMREF12) -> Option<(usize, usize)> {
    if mref.is_null() {
        return None;
    }
    
    // currently we only handle single contiguous references
    unsafe {
        if (*mref).count != 1 {
            return None;
        }
        Some((*mref).reftbl[0].dim())
    }
}