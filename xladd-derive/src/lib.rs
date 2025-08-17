use proc_macro::TokenStream;
use quote::quote;
use syn::{parse_macro_input, ItemFn, FnArg, Pat};

#[proc_macro_attribute]
pub fn xl_func(attr: TokenStream, input: TokenStream) -> TokenStream {
    let input_fn = parse_macro_input!(input as ItemFn);

    // Debug prints
    eprintln!("Processing function: {}", input_fn.sig.ident);
    
    // Parse attribute parameters - now includes param descriptions
    let mut category = String::new();
    let mut prefix = "xl".to_string();
    let mut rename = String::new();
    let mut single_threaded = true;
    let mut param_descriptions_from_attr = std::collections::HashMap::new();
    
    // Parse the attribute tokens for options
    let attr_str = attr.to_string();
    if !attr_str.is_empty() {
        // Parse parameters like: category="Math", params(age="Age in years", salary="Annual salary")
        parse_xl_func_attributes(&attr_str, &mut category, &mut prefix, &mut rename, 
                                &mut single_threaded, &mut param_descriptions_from_attr);
    }
    
    // Extract function name
    let fn_name = &input_fn.sig.ident;
    
    // Generate Excel function name
    let excel_fn_name = if !rename.is_empty() {
        rename
    } else {
        format!("{}_{}", prefix, fn_name)
    };
    let xl_fn_name = quote::format_ident!("{}", excel_fn_name);
    let xl_fn_name_str = xl_fn_name.to_string();
    
    // Generate registration function name and static name
    let static_args_name = quote::format_ident!("ARGS_{}", fn_name.to_string().to_uppercase());
    
    // Extract parameter information
    let mut param_names = Vec::new();
    let mut param_types = Vec::new();
    let mut param_descriptions = std::collections::HashMap::new();
    
    for input in &input_fn.sig.inputs {
        if let FnArg::Typed(pat_type) = input {
            if let Pat::Ident(pat_ident) = pat_type.pat.as_ref() {
                let param_name = &pat_ident.ident;
                param_names.push(param_name);
                param_types.push(&pat_type.ty);
                
                // Use description from attribute first, then fall back to default
                let description = param_descriptions_from_attr.get(&param_name.to_string())
                    .cloned()
                    .unwrap_or_else(|| format!("Parameter {}", param_name));
                
                param_descriptions.insert(param_name.to_string(), description);
            }
        }
    }
    
    // Parse documentation from function doc comments
    let mut function_description = String::new();
    let mut return_description = String::new();
    
    for attr in &input_fn.attrs {
        if attr.path().is_ident("doc") {
            // Extract the doc string from the attribute meta
            if let syn::Meta::NameValue(meta_name_value) = &attr.meta {
                if let syn::Expr::Lit(syn::ExprLit { lit: syn::Lit::Str(lit_str), .. }) = &meta_name_value.value {
                    let doc_string = lit_str.value(); 
                    let doc_line = doc_string.trim();
                    
                    if doc_line.starts_with("* ret:") {
                        return_description = doc_line[6..].trim().to_string();
                    } else if doc_line.starts_with("* ") && (doc_line.contains(':') || doc_line.contains("-")) {
                        let content = &doc_line[2..]; // Remove "* "
                        
                        // Try to find either delimiter
                        let delimiter_pos = if let Some(pos) = content.find(':') {
                            Some((pos, ':'))
                        } else if let Some(pos) = content.find("-") {
                            Some((pos, '-'))
                        } else {
                            None
                        };
                        
                        if let Some((pos, _delimiter)) = delimiter_pos {
                            let param_name = content[..pos].trim().replace('`', "");
                            let description = content[pos + 1..].trim();
                            
                            // ALWAYS use doc comment description (overwrite defaults)
                            param_descriptions.insert(param_name.to_string(), description.to_string());
                        }
                    } else if !doc_line.starts_with("*") && !doc_line.starts_with("#") && !doc_line.is_empty() {
                        if !function_description.is_empty() {
                            function_description.push(' ');
                        }
                        function_description.push_str(&doc_line);
                    }
                }
            }
        }
    }
    
    // Create the combined description for Excel
    let excel_description = if return_description.is_empty() {
        if function_description.is_empty() {
            "No description available".to_string()
        } else {
            function_description.clone()
        }
    } else {
        if function_description.is_empty() {
            format!("Returns: {}", return_description)
        } else {
            format!("{} Returns: {}", function_description, return_description)
        }
    };

    // Check length of description and truncate if necessary
    // (calls to xlfRegister will fail if any of the strings are longer than 255 characters)
    let excel_description = if excel_description.len() > 255 {
        eprintln!("⚠️  TRUNCATING description for {}: {} chars -> 255 chars", fn_name, excel_description.len());
        let mut truncated = excel_description.chars().take(252).collect::<String>();
        truncated.push_str("...");
        truncated
    } else {
        excel_description
    };
    
    // Generate argument conversion code
    let arg_conversions = param_names.iter().zip(param_types.iter()).map(|(name, ty)| {
        quote! {
            let #name = {
                let variant = xladd_core::variant::Variant::from(#name);
                if variant.is_missing_or_null() {
                    return xladd_core::xlcall::LPXLOPER12::from(
                        xladd_core::variant::Variant::from("Missing argument")
                    );
                }
                match std::convert::TryInto::<#ty>::try_into(&variant) {
                    Ok(val) => val,
                    Err(e) => {
                        return xladd_core::xlcall::LPXLOPER12::from(
                            // xladd_core::variant::Variant::from(format!("Conversion error: {}", e))
                            xladd_core::variant::Variant::from(&format!("Conversion error: {}", e)) 
                        );
                    }
                }
            };
        }
    });
    
    // Generate Excel function arguments
    let xl_args = param_names.iter().map(|name| {
        quote! { #name: xladd_core::xlcall::LPXLOPER12 }
    });
    
    // Generate function call arguments
    let call_args = param_names.iter().map(|name| quote! { #name });
    
    // Extract return type to determine if it's a Result
    let return_type = &input_fn.sig.output;
    let is_result_type = match return_type {
        syn::ReturnType::Type(_, ty) => {
            if let syn::Type::Path(type_path) = ty.as_ref() {
                type_path.path.segments.first()
                    .map(|seg| seg.ident == "Result")
                    .unwrap_or(false)
            } else {
                false
            }
        }
        _ => false,
    };

    // Generate different wrapper code based on return type
    let function_call = if is_result_type {
        // For Result<T, E> return types
        quote! {
            match #fn_name(#(#call_args),*) {
                Ok(result) => {
                    xladd_core::xlcall::LPXLOPER12::from(xladd_core::variant::Variant::from(result))
                }
                Err(e) => {
                    xladd_core::xlcall::LPXLOPER12::from(
                        xladd_core::variant::Variant::from(&e.to_string())
                    )
                }
            }
        }
    } else {
        // For direct return types (f64, Vec<f64>, etc.)
        quote! {
            let result = #fn_name(#(#call_args),*);
            xladd_core::xlcall::LPXLOPER12::from(xladd_core::variant::Variant::from(result))
        }
    };
    
    // Generate the registration string (Q for each parameter + Q for return)
    let mut reg_string = param_names.iter().map(|_| "Q").collect::<String>();
    reg_string.push('Q'); // Regular return value
    
    if !single_threaded {
        reg_string.push('$'); // Thread-safe marker
    }
    
    // Generate the parameter names string for registration
    let param_names_str = param_names.iter()
        .map(|name| name.to_string())
        .collect::<Vec<_>>()
        .join(",");

    // Check and truncate if needed
    let param_names_str = if param_names_str.len() > 255 {
        eprintln!("⚠️  TRUNCATING param names for {}: {} chars -> 255 chars", fn_name, param_names_str.len());
        let mut truncated = param_names_str.chars().take(252).collect::<String>();
        truncated.push_str("...");
        truncated
    } else {
        param_names_str
    };
    
    // Create argument info structs with Excel bug workaround and length checking
    let arg_infos = param_names.iter().enumerate().map(|(i, name)| {
        let name_str = name.to_string();
        let mut description = param_descriptions.get(&name_str)
            .cloned()
            .unwrap_or_else(|| format!("Parameter {}", name_str));
        
        // Check and truncate parameter description if needed
        if description.len() > 255 {
            eprintln!("⚠️  TRUNCATING param description for '{}': {} chars -> 255 chars", name_str, description.len());
            description = description.chars().take(252).collect::<String>();
            description.push_str("...");
        }
        
        // WORKAROUND: Excel bug truncates the last character of the final parameter
        if i == param_names.len() - 1 && param_names.len() > 0 {
            description.push_str(".."); // Add trailing chars to last parameter
        }
        
        quote! {
            xladd_core::registrator::ArgInfo {
                name: #name_str,
                description: #description,
                excel_type: "Q",
            }
        }
    });
    
    // Generate the complete macro output
    let expanded = quote! {
        // The original user function (unchanged)
        #input_fn
        
        // Excel wrapper function
        #[unsafe(no_mangle)]
        extern "system" fn #xl_fn_name(#(#xl_args),*) -> xladd_core::xlcall::LPXLOPER12 {
            // Convert arguments from Excel types to Rust types
            #(#arg_conversions)*
            
            // Call the original function with appropriate error handling
            #function_call
        }
        
        // Create a static array of argument info with proper UPPER_CASE naming
        static #static_args_name: &[xladd_core::registrator::ArgInfo] = &[#(#arg_infos),*];
        
        // Auto-register this function when the module loads
        inventory::submit! {
            xladd_core::registrator::FunctionRegistration {
                xl_name: #xl_fn_name_str,
                arg_types: #reg_string,
                arg_names: #param_names_str,
                category: #category,
                description: #excel_description,
                arg_infos: #static_args_name,
            }
        }
    };
    
    TokenStream::from(expanded)
}

/// Parse xl_func attribute parameters including param descriptions
fn parse_xl_func_attributes(
    attr_str: &str,
    category: &mut String,
    prefix: &mut String, 
    rename: &mut String,
    single_threaded: &mut bool,
    param_descriptions: &mut std::collections::HashMap<String, String>
) {
    // Simple parser for: category="Math", params(age="Age in years", salary="Annual salary")
    // This is a basic implementation - could be made more robust
    
    if attr_str.contains("category=") {
        if let Some(start) = attr_str.find("category=\"") {
            let start = start + 10; // Skip 'category="'
            if let Some(end) = attr_str[start..].find('"') {
                *category = attr_str[start..start + end].to_string();
            }
        }
    }
    
    if attr_str.contains("prefix=") {
        if let Some(start) = attr_str.find("prefix=\"") {
            let start = start + 8; // Skip 'prefix="'
            if let Some(end) = attr_str[start..].find('"') {
                *prefix = attr_str[start..start + end].to_string();
            }
        }
    }
    
    if attr_str.contains("rename=") {
        if let Some(start) = attr_str.find("rename=\"") {
            let start = start + 8; // Skip 'rename="'
            if let Some(end) = attr_str[start..].find('"') {
                *rename = attr_str[start..start + end].to_string();
            }
        }
    }
    
    if attr_str.contains("threadsafe") {
        *single_threaded = false;
    }
    
    // Parse params(param1="desc1", param2="desc2")
    if let Some(params_start) = attr_str.find("params(") {
        let params_start = params_start + 7; // Skip 'params('
        if let Some(params_end) = attr_str[params_start..].find(')') {
            let params_str = &attr_str[params_start..params_start + params_end];
            
            // Split by commas and parse param="description" pairs
            for pair in params_str.split(',') {
                let pair = pair.trim();
                if let Some(eq_pos) = pair.find('=') {
                    let param_name = pair[..eq_pos].trim().to_string();
                    let desc_part = &pair[eq_pos + 1..].trim();
                    if desc_part.starts_with('"') && desc_part.ends_with('"') {
                        let description = desc_part[1..desc_part.len() - 1].to_string();
                        param_descriptions.insert(param_name, description);
                    }
                }
            }
        }
    }
}