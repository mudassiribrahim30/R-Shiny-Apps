# Data2SPSS - Professional Data Conversion Tool
# Now with bidirectional conversion: Convert from SPSS to other formats

library(shiny)
library(shinythemes)
library(readxl)
library(readr)
library(haven)
library(foreign)
library(rio)
library(DT)
library(dplyr)
library(labelled)
library(purrr)
library(shinyjs)
library(writexl)


# Enhanced function to handle various file formats with better error handling
load_data_file <- function(file_path, file_type) {
  tryCatch({
    # Handle SAS xpt files specifically for JMP compatibility
    if (file_type == "xpt") {
      # Use haven to read xpt files (SAS transport files)
      data <- read_xpt(file_path)
      return(data)
    }
    
    data <- switch(file_type,
                   "csv" = read_csv(file_path, show_col_types = FALSE),
                   "xlsx" = read_excel(file_path),
                   "xls" = read_excel(file_path),
                   "dta" = read_stata(file_path),  # This preserves Stata labels
                   "sav" = read_sav(file_path),
                   "sas7bdat" = read_sas(file_path),
                   "xpt" = read_xpt(file_path),  # SAS transport file
                   "rdata" = {
                     env <- new.env()
                     load(file_path, envir = env)
                     obj_names <- ls(env)
                     if (length(obj_names) == 0) {
                       stop("No objects found in RData file")
                     }
                     get(obj_names[1], envir = env)
                   },
                   "tsv" = read_tsv(file_path, show_col_types = FALSE),
                   "txt" = read_delim(file_path, delim = "\t", show_col_types = FALSE),
                   # Use rio as fallback for other formats
                   rio::import(file_path)
    )
    
    # Ensure we return a data frame
    if (!is.data.frame(data)) {
      data <- as.data.frame(data)
    }
    
    # Convert all character columns to UTF-8 encoding
    data <- data %>%
      mutate(across(where(is.character), ~iconv(., to = "UTF-8", sub = "?")))
    
    return(data)
  }, error = function(e) {
    # Try rio as fallback if specific method fails
    tryCatch({
      data <- rio::import(file_path)
      if (!is.data.frame(data)) {
        data <- as.data.frame(data)
      }
      return(data)
    }, error = function(e2) {
      stop(paste("Error loading file:", e$message, "| Fallback error:", e2$message))
    })
  })
}

# Robust function to clean variable names for SPSS and SAS
clean_variable_names <- function(names, format = "spss") {
  cleaned_names <- sapply(names, function(name) {
    if (is.na(name) || name == "") {
      return("Variable")
    }
    
    # Remove special characters and keep only alphanumeric and underscores
    clean_name <- gsub("[^a-zA-Z0-9_]", "", name)
    
    # Ensure name starts with a letter
    if (!grepl("^[a-zA-Z]", clean_name)) {
      clean_name <- paste0("V", clean_name)
    }
    
    # Truncate based on format requirements
    if (format == "sas") {
      # SAS variable names max 32 characters
      clean_name <- substr(clean_name, 1, 32)
    } else {
      # SPSS limit (64 characters)
      clean_name <- substr(clean_name, 1, 64)
    }
    
    if (clean_name == "") {
      clean_name <- ifelse(format == "sas", "Var", "Variable")
    }
    
    return(clean_name)
  })
  
  # Ensure unique names
  cleaned_names <- make.unique(cleaned_names, sep = "_")
  
  return(cleaned_names)
}

# CRITICAL FIX: Enhanced function to detect if variable is truly categorical vs continuous
is_truly_categorical <- function(x) {
  # If it's labelled from Stata, it's categorical
  if (inherits(x, "haven_labelled")) return(TRUE)
  
  # If it has value labels attribute from Stata, it's categorical
  if (!is.null(attr(x, "labels"))) return(TRUE)
  
  # If it's a factor, it's categorical
  if (is.factor(x)) return(TRUE)
  
  # If it's character with limited unique values, it's categorical
  if (is.character(x)) {
    unique_vals <- unique(na.omit(x))
    if (length(unique_vals) <= 100) {  # Increased threshold for character variables
      return(TRUE)
    }
    return(FALSE)
  }
  
  # If it's not numeric, it's categorical
  if (!is.numeric(x)) return(TRUE)
  
  # Get unique values (excluding NAs)
  unique_vals <- unique(na.omit(x))
  
  # If it has few unique values, it might be categorical
  if (length(unique_vals) <= 50) {  # Increased threshold
    return(TRUE)
  }
  
  # If it has many unique values or decimal values, it's continuous
  return(FALSE)
}

# Function to detect continuous numeric variables (0 to infinity)
is_continuous_numeric <- function(x) {
  if (!is.numeric(x)) return(FALSE)
  
  unique_vals <- unique(na.omit(x))
  
  # If it has many unique values, it's continuous
  if (length(unique_vals) > 50) {
    return(TRUE)
  }
  
  # If it has decimal values, it's continuous
  if (any(unique_vals %% 1 != 0)) {
    return(TRUE)
  }
  
  # If values range widely (like measurements), it's continuous
  if (length(unique_vals) > 1) {
    range_val <- max(unique_vals) - min(unique_vals)
    if (range_val > 1000) {
      return(TRUE)
    }
  }
  
  return(FALSE)
}

# Function to determine SPSS measure type
get_spss_measure_type <- function(x) {
  if (is_continuous_numeric(x)) {
    return("scale")
  } else if (is_truly_categorical(x)) {
    return("nominal")
  } else {
    return("scale")
  }
}

# Function to set SPSS measure type attribute
set_spss_measure <- function(x, measure_type) {
  attr(x, "measure") <- measure_type
  attr(x, "SPSS.measure") <- measure_type
  return(x)
}

# CRITICAL FIX: Convert factors to numeric with proper value labels
convert_factor_to_spss <- function(x, variable_name) {
  if (!is.factor(x)) return(x)
  
  levels_vec <- levels(x)
  numeric_values <- 1:length(levels_vec)
  value_labels <- setNames(numeric_values, levels_vec)
  numeric_data <- as.numeric(x)
  
  labelled_data <- labelled(numeric_data, labels = value_labels)
  var_label(labelled_data) <- variable_name
  labelled_data <- set_spss_measure(labelled_data, "nominal")
  
  return(labelled_data)
}

# CRITICAL FIX: ALWAYS convert character variables to numeric with value labels
convert_character_to_spss <- function(x, variable_name) {
  if (!is.character(x)) return(x)
  
  unique_vals <- unique(na.omit(x))
  unique_vals <- unique_vals[!is.na(unique_vals) & unique_vals != ""]
  
  # If no unique values, return as is
  if (length(unique_vals) == 0) {
    var_label(x) <- variable_name
    x <- set_spss_measure(x, "nominal")
    return(x)
  }
  
  # ALWAYS convert to numeric with value labels, regardless of length
  numeric_values <- 1:length(unique_vals)
  value_labels <- setNames(numeric_values, unique_vals)
  
  # Create mapping from character to numeric
  char_to_num <- setNames(numeric_values, unique_vals)
  numeric_data <- char_to_num[as.character(x)]
  
  # Handle NAs
  numeric_data[is.na(x)] <- NA
  
  labelled_data <- labelled(numeric_data, labels = value_labels)
  var_label(labelled_data) <- variable_name
  labelled_data <- set_spss_measure(labelled_data, "nominal")
  
  return(labelled_data)
}

# CRITICAL FIX: Handle Stata labelled variables - PRESERVE ORIGINAL NUMERIC VALUES
convert_numeric_to_spss <- function(x, variable_name) {
  if (!is.numeric(x)) return(x)
  
  # Check if it's a Stata labelled variable
  if (inherits(x, "haven_labelled")) {
    # PRESERVE original numeric values and labels from Stata
    value_labels <- attr(x, "labels")
    numeric_values <- as.numeric(x)  # Get the underlying numeric values
    
    if (!is.null(value_labels)) {
      # Use the original numeric values with Stata's value labels
      labelled_data <- labelled(numeric_values, labels = value_labels)
      var_label(labelled_data) <- variable_name
      labelled_data <- set_spss_measure(labelled_data, "nominal")
      return(labelled_data)
    }
  }
  
  # Check for regular value labels attribute
  value_labels <- attr(x, "labels")
  if (!is.null(value_labels)) {
    labelled_data <- labelled(as.numeric(x), labels = value_labels)
    var_label(labelled_data) <- variable_name
    labelled_data <- set_spss_measure(labelled_data, "nominal")
    return(labelled_data)
  }
  
  # Check if it's continuous numeric
  if (is_continuous_numeric(x)) {
    var_label(x) <- variable_name
    x <- set_spss_measure(x, "scale")
    return(x)
  } else {
    # CATEGORICAL NUMERIC: Use original numeric values as labels
    unique_vals <- unique(na.omit(x))
    value_labels <- setNames(unique_vals, as.character(unique_vals))
    
    labelled_data <- labelled(x, labels = value_labels)
    var_label(labelled_data) <- variable_name
    labelled_data <- set_spss_measure(labelled_data, "nominal")
    return(labelled_data)
  }
}

# Function to handle logical variables
convert_logical_to_spss <- function(x, variable_name) {
  if (!is.logical(x)) return(x)
  
  value_labels <- c(`FALSE` = 0, `TRUE` = 1)
  numeric_data <- as.numeric(x)
  labelled_data <- labelled(numeric_data, labels = value_labels)
  var_label(labelled_data) <- variable_name
  labelled_data <- set_spss_measure(labelled_data, "nominal")
  
  return(labelled_data)
}

# CRITICAL FIX: Main function to ensure ALL categorical variables become numeric
convert_to_spss_format <- function(x, variable_name) {
  tryCatch({
    current_label <- attr(x, "label", exact = TRUE)
    var_lbl <- if (is.null(current_label) || current_label == "") variable_name else current_label
    
    # Handle haven_labelled variables from Stata FIRST - PRESERVE NUMERIC VALUES
    if (inherits(x, "haven_labelled")) {
      value_labels <- attr(x, "labels")
      numeric_values <- as.numeric(x)  # Get underlying numeric values
      
      if (!is.null(value_labels)) {
        labelled_data <- labelled(numeric_values, labels = value_labels)
        var_label(labelled_data) <- var_lbl
        labelled_data <- set_spss_measure(labelled_data, "nominal")
        return(labelled_data)
      }
    }
    
    # Handle character variables - ALWAYS convert to numeric
    if (is.character(x)) {
      return(convert_character_to_spss(x, var_lbl))
    }
    
    if (is.factor(x)) {
      return(convert_factor_to_spss(x, var_lbl))
    } else if (is.logical(x)) {
      return(convert_logical_to_spss(x, var_lbl))
    } else if (is.numeric(x)) {
      return(convert_numeric_to_spss(x, var_lbl))
    } else {
      # For any other type, try to convert to character then to numeric
      tryCatch({
        char_x <- as.character(x)
        return(convert_character_to_spss(char_x, var_lbl))
      }, error = function(e) {
        # Fallback: keep as is but add label
        var_label(x) <- var_lbl
        x <- set_spss_measure(x, "nominal")
        return(x)
      })
    }
    
  }, error = function(e) {
    warning(paste("Error processing variable", variable_name, ":", e$message))
    # Fallback: try to convert to character then numeric
    tryCatch({
      return(convert_character_to_spss(as.character(x), variable_name))
    }, error = function(e2) {
      result <- x
      var_label(result) <- variable_name
      return(result)
    })
  })
}

# Robust function to convert to SPSS format
convert_to_spss <- function(data) {
  tryCatch({
    spss_data <- as.data.frame(data)
    original_names <- names(spss_data)
    cleaned_names <- clean_variable_names(original_names, "spss")
    names(spss_data) <- cleaned_names
    
    # First pass: convert all variables
    for (i in 1:ncol(spss_data)) {
      col_name <- names(spss_data)[i]
      original_col_name <- original_names[i]
      col_data <- spss_data[[col_name]]
      
      spss_data[[col_name]] <- convert_to_spss_format(col_data, original_col_name)
    }
    
    # Second pass: ensure all variables are properly labelled
    for (col_name in names(spss_data)) {
      var <- spss_data[[col_name]]
      
      # Ensure variable label is set
      if (is.null(var_label(var))) {
        var_label(spss_data[[col_name]]) <- col_name
      }
      
      # Ensure measure type is set
      if (is.null(attr(var, "SPSS.measure"))) {
        if (is.numeric(var) && is_continuous_numeric(var)) {
          spss_data[[col_name]] <- set_spss_measure(var, "scale")
        } else {
          spss_data[[col_name]] <- set_spss_measure(var, "nominal")
        }
      }
    }
    
    return(spss_data)
    
  }, error = function(e) {
    stop(paste("Error converting to SPSS format:", e$message))
  })
}

# CRITICAL FIX: Proper SPSS export ensuring numeric variables
write_spss_with_attributes <- function(data, file_path) {
  tryCatch({
    spss_export <- as.data.frame(data)
    
    # Apply proper measure types to each variable
    for (col_name in names(spss_export)) {
      var <- spss_export[[col_name]]
      
      # Ensure variable name is valid for SPSS
      valid_name <- clean_variable_names(col_name, "spss")[1]
      if (valid_name != col_name) {
        names(spss_export)[names(spss_export) == col_name] <- valid_name
        col_name <- valid_name
        var <- spss_export[[col_name]]
      }
      
      # Set measure type
      if (!is.null(attr(var, "SPSS.measure"))) {
        measure_type <- attr(var, "SPSS.measure")
        attr(spss_export[[col_name]], "measure") <- measure_type
      } else {
        if (!is.null(val_labels(var))) {
          attr(spss_export[[col_name]], "measure") <- "nominal"
        } else if (is.numeric(var) && is_continuous_numeric(var)) {
          attr(spss_export[[col_name]], "measure") <- "scale"
        } else {
          attr(spss_export[[col_name]], "measure") <- "nominal"
        }
      }
      
      # Set SPSS format
      if (is.numeric(spss_export[[col_name]])) {
        attr(spss_export[[col_name]], "format.spss") <- "F8.2"
      }
    }
    
    # Final validation
    final_names <- names(spss_export)
    if (any(!grepl("^[a-zA-Z][a-zA-Z0-9_]*$", final_names))) {
      stop("Invalid variable names detected after cleaning")
    }
    
    # Export to SPSS
    haven::write_sav(spss_export, file_path)
    
    return(TRUE)
    
  }, error = function(e) {
    tryCatch({
      write_sav(data, file_path)
      return(TRUE)
    }, error = function(e2) {
      stop(paste("Export failed:", e$message, "| Fallback error:", e2$message))
    })
  })
}

# Enhanced data summary function
get_data_summary <- function(data) {
  if (is.null(data)) return(NULL)
  
  summary_df <- data.frame(
    Variable = names(data),
    Type = sapply(data, function(x) {
      if (inherits(x, "haven_labelled")) "Labelled Numeric"
      else paste(class(x), collapse = ", ")
    }),
    Missing = sapply(data, function(x) sum(is.na(x))),
    Unique = sapply(data, function(x) length(unique(na.omit(x)))),
    Is_Continuous = sapply(data, is_continuous_numeric),
    SPSS_Measure = sapply(data, get_spss_measure_type),
    Will_Be_Numeric = sapply(data, function(x) {
      if (is_continuous_numeric(x)) "YES (Scale)"
      else if (is_truly_categorical(x)) "YES (Nominal)"
      else "YES (Nominal)"
    }),
    stringsAsFactors = FALSE
  )
  return(summary_df)
}

# Enhanced SPSS metadata function
get_spss_metadata <- function(spss_data) {
  if (is.null(spss_data)) return(NULL)
  
  metadata <- data.frame(
    Variable = names(spss_data),
    Type = sapply(spss_data, function(x) {
      if (inherits(x, "haven_labelled")) "Labelled Numeric"
      else if (is.numeric(x)) "Numeric" 
      else if (is.character(x)) "String" 
      else class(x)[1]
    }),
    Label = sapply(spss_data, function(x) {
      lbl <- var_label(x)
      if (is.null(lbl) || lbl == "") "No label" else as.character(lbl)
    }),
    SPSS_Measure = sapply(spss_data, function(x) {
      measure <- attr(x, "SPSS.measure")
      if (is.null(measure)) "nominal" else measure
    }),
    Has_Value_Labels = sapply(spss_data, function(x) {
      lbls <- val_labels(x)
      if (is.null(lbls) || length(lbls) == 0) "NO" else "YES"
    }),
    Label_Count = sapply(spss_data, function(x) {
      lbls <- val_labels(x)
      if (is.null(lbls)) 0 else length(lbls)
    }),
    Sample_Labels = sapply(spss_data, function(x) {
      lbls <- val_labels(x)
      if (is.null(lbls) || length(lbls) == 0) {
        "No value labels"
      } else {
        sample_lbls <- head(lbls, 3)
        paste(paste0(names(sample_lbls), "=", sample_lbls), collapse = "; ")
      }
    }),
    stringsAsFactors = FALSE
  )
  return(metadata)
}

# ============================================
# BIDIRECTIONAL CONVERSION FUNCTIONS
# Convert from any loaded data to target format (SPSS to other formats)
# ============================================

# CRITICAL FIX: Enhanced function to clean data for export to any format
clean_data_for_export <- function(data, target_format) {
  export_data <- as.data.frame(data)
  
  # Clean variable names first
  if (target_format %in% c("sas", "sas7bdat", "xpt")) {
    names(export_data) <- clean_variable_names(names(export_data), "sas")
  } else {
    names(export_data) <- clean_variable_names(names(export_data), "spss")
  }
  
  # Process each column
  for (i in seq_along(export_data)) {
    col <- export_data[[i]]
    col_name <- names(export_data)[i]
    
    # Remove any problematic characters from column name for display
    if (target_format %in% c("sas", "sas7bdat", "xpt")) {
      # For SAS, ensure names are valid
      col_name_clean <- gsub("[^a-zA-Z0-9_]", "", col_name)
      if (col_name_clean != col_name) {
        names(export_data)[i] <- col_name_clean
      }
    }
    
    # Handle haven_labelled variables
    if (inherits(col, "haven_labelled")) {
      labels <- attr(col, "labels")
      if (!is.null(labels) && length(labels) > 0) {
        # Convert to factor with labels
        export_data[[i]] <- factor(col, levels = labels, labels = names(labels))
      } else {
        export_data[[i]] <- as.numeric(col)
      }
    }
    # Convert other labelled objects
    else if (!is.null(attr(col, "labels"))) {
      labels <- attr(col, "labels")
      export_data[[i]] <- factor(col, levels = labels, labels = names(labels))
    }
    # Handle factors - ensure levels are valid
    else if (is.factor(col)) {
      # Clean factor levels to remove illegal characters
      clean_levels <- gsub("[^a-zA-Z0-9_]", "", levels(col))
      if (any(clean_levels != levels(col))) {
        export_data[[i]] <- factor(col, levels = levels(col), labels = clean_levels)
      }
    }
    # Handle character - remove illegal characters
    else if (is.character(col)) {
      # Remove control characters and clean
      export_data[[i]] <- gsub("[^\x20-\x7E]", "", col)
    }
    # Keep numeric as is
    else if (is.numeric(col)) {
      export_data[[i]] <- col
    }
    # Default conversion to character with cleaning
    else {
      export_data[[i]] <- as.character(col)
      export_data[[i]] <- gsub("[^\x20-\x7E]", "", export_data[[i]])
    }
  }
  
  return(export_data)
}

# Function to convert loaded data to various formats with robust error handling
convert_to_other_format <- function(data, target_format, file_path) {
  tryCatch({
    # Clean data for export first
    export_data <- clean_data_for_export(data, target_format)
    
    # Export based on target format with error handling
    result <- switch(tolower(target_format),
                     "csv" = write_csv(export_data, file_path),
                     "xlsx" = {
                       # For large datasets, warn but proceed
                       if (nrow(export_data) > 1000000) {
                         warning("Large dataset being written to Excel - may be slow")
                       }
                       write_xlsx(export_data, file_path)
                     },
                     "dta" = write_dta(export_data, file_path),
                     "rdata" = save(export_data, file = file_path),
                     "tsv" = write_tsv(export_data, file_path),
                     "txt" = write.table(export_data, file_path, sep = "\t", row.names = FALSE, quote = FALSE),
                     "sas" = write_sas_with_fallback(export_data, file_path),
                     "sas7bdat" = write_sas_with_fallback(export_data, file_path),
                     "xpt" = write_xpt_with_fallback(export_data, file_path),
                     stop(paste("Unsupported target format:", target_format))
    )
    return(TRUE)
  }, error = function(e) {
    # Last resort fallback: try to write as CSV with cleaned names
    tryCatch({
      fallback_path <- gsub("\\.(sas7bdat|xpt)$", ".csv", file_path)
      write_csv(clean_data_for_export(data, "csv"), fallback_path)
      stop(paste("Original format failed:", e$message, 
                 "\nData was saved as CSV instead at:", fallback_path))
    }, error = function(e2) {
      stop(paste("Error converting to", target_format, ":", e$message))
    })
  })
}

# Write SAS with fallback for problematic data
write_sas_with_fallback <- function(data, file_path) {
  # Ensure .sas7bdat extension
  if (!grepl("\\.sas7bdat$", file_path)) {
    file_path <- paste0(file_path, ".sas7bdat")
  }
  
  # Try to write with original method
  tryCatch({
    haven::write_sas(data, file_path)
    return(TRUE)
  }, error = function(e) {
    # If fails, try with further data cleaning
    tryCatch({
      # Convert all columns to basic types
      fallback_data <- as.data.frame(lapply(data, function(x) {
        if (is.factor(x)) {
          return(as.character(x))
        } else if (is.numeric(x)) {
          return(x)
        } else {
          return(as.character(x))
        }
      }))
      
      # Clean names again
      names(fallback_data) <- clean_variable_names(names(fallback_data), "sas")
      haven::write_sas(fallback_data, file_path)
      return(TRUE)
    }, error = function(e2) {
      # Last resort: save as RDS and provide info
      rds_path <- gsub("\\.sas7bdat$", ".rds", file_path)
      saveRDS(data, rds_path)
      stop(paste("SAS export failed:", e$message, 
                 "\nData saved as RDS instead at:", rds_path))
    })
  })
}

# Write XPT with robust error handling for illegal characters
write_xpt_with_fallback <- function(data, file_path) {
  # Ensure .xpt extension
  if (!grepl("\\.xpt$", file_path)) {
    file_path <- paste0(file_path, ".xpt")
  }
  
  # Clean variable names for SAS XPT (max 32 chars, alphanumeric and underscore)
  cleaned_data <- data
  original_names <- names(cleaned_data)
  
  # Aggressively clean variable names for XPT
  clean_xpt_names <- function(names) {
    cleaned <- sapply(names, function(name) {
      # Remove all non-alphanumeric and underscore
      clean <- gsub("[^a-zA-Z0-9_]", "", name)
      # Must start with letter
      if (!grepl("^[a-zA-Z]", clean)) {
        clean <- paste0("V", clean)
      }
      # Max 32 characters for SAS
      clean <- substr(clean, 1, 32)
      if (clean == "") clean <- "Variable"
      return(clean)
    })
    # Ensure uniqueness
    cleaned <- make.unique(cleaned, sep = "_")
    return(cleaned)
  }
  
  names(cleaned_data) <- clean_xpt_names(names(cleaned_data))
  
  # Clean column data to remove illegal characters
  for (i in seq_along(cleaned_data)) {
    col <- cleaned_data[[i]]
    
    # Handle factors
    if (is.factor(col)) {
      # Clean factor levels
      clean_levels <- gsub("[^a-zA-Z0-9_]", "", levels(col))
      if (any(clean_levels != levels(col))) {
        cleaned_data[[i]] <- factor(col, levels = levels(col), labels = clean_levels)
      }
    }
    # Handle character columns
    else if (is.character(col)) {
      # Remove non-printable characters
      cleaned_data[[i]] <- gsub("[^\x20-\x7E]", "", col)
      # Truncate long strings (XPT has limit)
      max_length <- 200
      cleaned_data[[i]] <- substr(cleaned_data[[i]], 1, max_length)
    }
    # Convert logical to numeric
    else if (is.logical(col)) {
      cleaned_data[[i]] <- as.numeric(col)
    }
    # Ensure all columns are basic types
    else if (!is.numeric(col) && !is.character(col)) {
      cleaned_data[[i]] <- as.character(col)
    }
  }
  
  # Try to write XPT with haven
  tryCatch({
    haven::write_xpt(cleaned_data, file_path, version = 8)
    return(TRUE)
  }, error = function(e) {
    # If still failing, try with even more aggressive cleaning
    tryCatch({
      # Convert everything to character with basic cleaning
      ultimate_fallback <- as.data.frame(lapply(cleaned_data, function(x) {
        as.character(x)
      }))
      
      # Limit to 100 columns max for XPT (safety)
      if (ncol(ultimate_fallback) > 100) {
        ultimate_fallback <- ultimate_fallback[, 1:100]
        warning("Dataset truncated to 100 columns for XPT export")
      }
      
      # Limit rows for XPT (safety for very large datasets)
      if (nrow(ultimate_fallback) > 1000000) {
        ultimate_fallback <- ultimate_fallback[1:1000000, ]
        warning("Dataset truncated to 1,000,000 rows for XPT export")
      }
      
      haven::write_xpt(ultimate_fallback, file_path, version = 8)
      return(TRUE)
    }, error = function(e2) {
      # Last resort: save as CSV with warning
      csv_path <- gsub("\\.xpt$", ".csv", file_path)
      write_csv(cleaned_data, csv_path)
      stop(paste("XPT export failed:", e$message, 
                 "\nData saved as CSV instead at:", csv_path))
    })
  })
}

# UI Definition with improved font sizes, theme toggle, and moving app name
ui <- fluidPage(
  theme = shinytheme("flatly"),
  useShinyjs(),
  title = "Data2SPSS - Professional Data Conversion Tool",
  titleWidth = 350,
  
  
  # JavaScript for theme toggle and upload size removal
  tags$head(
    tags$meta(name = "viewport", content = "width=device-width, initial-scale=1.0"),
    # CRITICAL FIX: Remove upload size limit via JavaScript
    tags$script(HTML("
      // Override Shiny's max upload size
      if (window.Shiny) {
        Shiny.options = Shiny.options || {};
        Shiny.options.maxUploadSize = Infinity;
      }
      
      // Also set HTML file input attribute to remove any browser limits
      $(document).on('shiny:connected', function() {
        // This ensures file input doesn't have size restrictions
        $('input[type=\"file\"]').attr('data-max-size', '0');
      });
    ")),
    tags$script(HTML("
      $(document).ready(function() {
        // Check for saved theme preference
        const savedTheme = localStorage.getItem('data2spss-theme');
        if (savedTheme === 'dark') {
          $('body').addClass('dark-theme');
          $('#themeToggle i').removeClass('fa-moon').addClass('fa-sun');
        }
        
        // Theme toggle functionality
        $('#themeToggle').click(function() {
          $('body').toggleClass('dark-theme');
          
          if ($('body').hasClass('dark-theme')) {
            $('#themeToggle i').removeClass('fa-moon').addClass('fa-sun');
            localStorage.setItem('data2spss-theme', 'dark');
          } else {
            $('#themeToggle i').removeClass('fa-sun').addClass('fa-moon');
            localStorage.setItem('data2spss-theme', 'light');
          }
        });
      });
    ")),
    
    tags$style(HTML("

/* Professional Interface - Subtle Zoom (98% equivalent) */
:root {
  --hopkins-blue: #002D72;
  --hopkins-dark-blue: #001A4D;
  --hopkins-light-blue: #E6F0F8;
  --hopkins-gold: #CFB87C;
  --hopkins-gray: #F8F9FA;
  --hopkins-dark-gray: #2C3E50;
  --hopkins-accent: #68ACE5;
  --text-primary: #333333;
  --text-secondary: #555555;
  --text-light: #777777;
}

/* Dark theme variables */
body.dark-theme {
  --hopkins-blue: #4A90E2;
  --hopkins-dark-blue: #2C3E50;
  --hopkins-light-blue: #1A2530;
  --hopkins-gold: #CFB87C;
  --hopkins-gray: #121212;
  --hopkins-dark-gray: #E0E0E0;
  --hopkins-accent: #68ACE5;
  --text-primary: #E0E0E0;
  --text-secondary: #B0B0B0;
  --text-light: #888888;
  background-color: var(--hopkins-gray) !important;
  color: var(--text-primary) !important;
}

/* Clean base styling - subtle zoom effect */
html, body {
  height: 100%;
  margin: 0;
  padding: 0;
  overflow-x: hidden;
  font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
  font-size: 13.5px;
  line-height: 1.55;
  color: var(--text-primary);
  background-color: var(--hopkins-gray);
  transition: background-color 0.3s ease, color 0.3s ease;
}

/* Subtle container scaling - just a slight zoom-out */
.container-fluid {
  padding-left: 0;
  padding-right: 0;
  max-width: 1380px;
  margin: 0 auto;
  overflow-x: hidden;
  transform: scale(0.985);
  transform-origin: top center;
  width: 101.52%;
}

/* Floating App Name */
.floating-app-name {
  position: fixed;
  top: 15px;
  left: 15px;
  z-index: 1000;
  font-size: 26px;
  font-weight: 700;
  color: var(--hopkins-blue);
  background-color: rgba(255, 255, 255, 0.95);
  padding: 7px 15px;
  border-radius: 8px;
  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
  animation: floatAnimation 6s ease-in-out infinite;
  text-decoration: none;
  border: 2px solid var(--hopkins-gold);
  transition: all 0.3s ease;
  backdrop-filter: blur(4px);
}

.dark-theme .floating-app-name {
  background-color: rgba(30, 30, 30, 0.95);
  color: var(--hopkins-accent);
}

.floating-app-name:hover {
  animation: floatAnimation 3s ease-in-out infinite;
  transform: scale(1.02);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
}

@keyframes floatAnimation {
  0%, 100% {
    transform: translateY(0px) rotate(-0.5deg);
  }
  25% {
    transform: translateY(-6px) rotate(0.5deg);
  }
  50% {
    transform: translateY(0px) rotate(-0.5deg);
  }
  75% {
    transform: translateY(-3px) rotate(0.5deg);
  }
}

/* Theme Toggle Button */
.theme-toggle-btn {
  position: fixed;
  top: 15px;
  right: 15px;
  z-index: 1000;
  background-color: var(--hopkins-blue);
  color: white;
  border: none;
  border-radius: 50%;
  width: 46px;
  height: 46px;
  font-size: 19px;
  cursor: pointer;
  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
  transition: all 0.3s ease;
  display: flex;
  align-items: center;
  justify-content: center;
}

.theme-toggle-btn:hover {
  transform: scale(1.05) rotate(10deg);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
}

/* Blog-inspired header */
.blog-header {
  background: linear-gradient(135deg, var(--hopkins-blue) 0%, var(--hopkins-dark-blue) 100%);
  color: white;
  padding: 3.5rem 0 2.2rem 0;
  margin-bottom: 0;
  border-bottom: 3px solid var(--hopkins-gold);
  box-shadow: 0 2px 8px rgba(0,0,0,0.08);
  transition: all 0.3s ease;
}

.blog-header-content {
  max-width: 1200px;
  margin: 0 auto;
  padding: 0 2rem;
  text-align: left;
}

.app-title-main {
  font-size: 3.5rem;
  font-weight: 700;
  margin: 0;
  letter-spacing: -0.5px;
  line-height: 1.1;
  text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
  color: white;
}

.app-subtitle-main {
  font-size: 1.6rem;
  font-weight: 300;
  opacity: 0.95;
  margin: 0.5rem 0 0 0;
  letter-spacing: 0.5px;
  color: white;
}

.app-tagline {
  font-size: 1.25rem;
  font-weight: 400;
  margin: 0.9rem 0 0 0;
  color: var(--hopkins-light-blue);
  max-width: 600px;
  line-height: 1.4;
}

/* Navigation */
.navbar {
  background-color: white !important;
  border: none;
  border-radius: 0;
  margin-bottom: 1.8rem;
  box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  min-height: 54px;
  transition: background-color 0.3s ease;
}

.dark-theme .navbar {
  background-color: var(--hopkins-light-blue) !important;
}

.navbar-brand {
  color: var(--hopkins-blue) !important;
  font-weight: 600;
  font-size: 1.15rem;
  padding: 15px;
}

.dark-theme .navbar-brand {
  color: var(--hopkins-accent) !important;
}

.nav-tabs {
  border-bottom: 1px solid #dee2e6;
  margin-bottom: 0;
  background: white;
  padding: 0 2rem;
  transition: background-color 0.3s ease;
}

.dark-theme .nav-tabs {
  background: var(--hopkins-light-blue);
  border-bottom-color: #444;
}

.nav-tabs > li > a {
  color: var(--hopkins-dark-gray);
  font-weight: 500;
  border-radius: 4px 4px 0 0;
  margin-right: 2px;
  padding: 13px 24px;
  border: 1px solid transparent;
  font-size: 1.05rem;
  transition: all 0.3s ease;
}

.dark-theme .nav-tabs > li > a {
  color: var(--text-primary);
}

.nav-tabs > li.active > a {
  background-color: white !important;
  color: var(--hopkins-blue) !important;
  border: 1px solid #dee2e6;
  border-bottom-color: transparent;
  font-weight: 600;
  border-top: 2px solid var(--hopkins-blue);
  font-size: 1.05rem;
}

.dark-theme .nav-tabs > li.active > a {
  background-color: var(--hopkins-light-blue) !important;
  color: var(--hopkins-accent) !important;
  border-color: #444;
}

.nav-tabs > li > a:hover {
  background-color: var(--hopkins-light-blue);
  border-color: var(--hopkins-light-blue);
  color: var(--hopkins-dark-blue);
}

.main-content {
  max-width: 1200px;
  margin: 0 auto;
  padding: 0 2rem 1.8rem 2rem;
  width: 100%;
  font-size: 1rem;
  box-sizing: border-box;
}

/* Cards */
.blog-card {
  background-color: white;
  padding: 1.6rem;
  border-radius: 10px;
  box-shadow: 0 2px 12px rgba(0,0,0,0.06);
  margin: 1.3rem 0;
  border-left: none;
  border-top: 1px solid #e9ecef;
  transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.3s ease;
  width: 100%;
  box-sizing: border-box;
  overflow: visible;
}

.dark-theme .blog-card {
  background-color: #1E1E1E;
  border-color: #444;
}

.blog-card h2 {
  font-size: 1.85rem;
  color: var(--hopkins-blue);
  margin-bottom: 0.9rem;
  font-weight: 600;
  word-wrap: break-word;
}

.dark-theme .blog-card h2 {
  color: var(--hopkins-accent);
}

.blog-card h3 {
  font-size: 1.45rem;
  color: var(--hopkins-dark-blue);
  margin: 1.1rem 0 0.7rem 0;
  font-weight: 600;
  word-wrap: break-word;
}

.dark-theme .blog-card h3 {
  color: var(--text-primary);
}

.blog-card h4 {
  font-size: 1.15rem;
  color: var(--hopkins-dark-gray);
  margin: 0.9rem 0 0.6rem 0;
  font-weight: 600;
  word-wrap: break-word;
}

.dark-theme .blog-card h4 {
  color: var(--text-primary);
}

.blog-card p {
  font-size: 1rem;
  color: var(--text-secondary);
  line-height: 1.55;
  margin-bottom: 0.8rem;
  word-wrap: break-word;
}

.blog-card:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 16px rgba(0,0,0,0.1);
}

/* Info Box */
.info-box {
  background-color: white;
  padding: 1.3rem;
  border-radius: 8px;
  border-left: 3px solid var(--hopkins-accent);
  margin-bottom: 1rem;
  box-shadow: 0 1px 4px rgba(0,0,0,0.04);
  font-size: 1rem;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
  overflow: hidden;
  word-wrap: break-word;
}

.dark-theme .info-box {
  background-color: #252525;
}

.info-box p {
  font-size: 1rem;
  color: var(--text-secondary);
  margin-bottom: 0.5rem;
  word-wrap: break-word;
}

/* Step Box */
.step-box {
  background-color: var(--hopkins-light-blue);
  padding: 1.3rem;
  border-radius: 8px;
  margin: 0.9rem 0;
  border-left: 3px solid var(--hopkins-blue);
  font-size: 1rem;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
  overflow: hidden;
}

.dark-theme .step-box {
  background-color: #2A2A2A;
}

.step-box h4 {
  font-size: 1.05rem;
  color: var(--hopkins-blue);
  margin-bottom: 0.5rem;
  word-wrap: break-word;
}

.dark-theme .step-box h4 {
  color: var(--hopkins-accent);
}

.step-box ul {
  font-size: 1rem;
  color: var(--text-secondary);
  padding-left: 1.2rem;
}

.step-box li {
  margin-bottom: 0.35rem;
  line-height: 1.5;
  word-wrap: break-word;
}

/* Buttons */
.btn-hopkins-primary {
  background: linear-gradient(135deg, var(--hopkins-blue) 0%, var(--hopkins-dark-blue) 100%);
  border: none;
  color: white;
  font-weight: 600;
  border-radius: 6px;
  transition: all 0.3s ease;
  padding: 11px 20px;
  font-size: 1rem;
  box-shadow: 0 1px 4px rgba(0,45,114,0.25);
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
  white-space: normal;
  word-wrap: break-word;
  text-align: center;
  display: block;
  margin: 8px 0;
}

.btn-hopkins-primary:hover {
  background: linear-gradient(135deg, var(--hopkins-dark-blue) 0%, var(--hopkins-blue) 100%);
  color: white;
  transform: translateY(-1px);
  box-shadow: 0 2px 8px rgba(0,45,114,0.3);
}

.btn-hopkins-success {
  background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
  border: none;
  color: white;
  border-radius: 6px;
  transition: all 0.3s ease;
  padding: 11px 20px;
  font-weight: 600;
  font-size: 1rem;
  box-shadow: 0 1px 4px rgba(40,167,69,0.25);
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
  white-space: normal;
  word-wrap: break-word;
  text-align: center;
  display: block;
  margin: 8px 0;
}

.btn-hopkins-success:hover {
  background: linear-gradient(135deg, #20c997 0%, #28a745 100%);
  color: white;
  transform: translateY(-1px);
  box-shadow: 0 2px 8px rgba(40,167,69,0.3);
}

/* Download button */
.shiny-download-link {
  width: 100% !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
  margin: 8px 0 !important;
  display: block !important;
  text-align: center !important;
  white-space: normal !important;
  word-wrap: break-word !important;
}

/* Loading spinner */
.loading-spinner {
  border: 4px solid #f3f3f3;
  border-top: 4px solid var(--hopkins-blue);
  border-radius: 50%;
  width: 48px;
  height: 48px;
  animation: spin 1s linear infinite;
  margin: 20px auto;
}

/* Feature grid */
.feature-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(270px, 1fr));
  gap: 1.3rem;
  margin: 1.6rem 0;
}

.feature-item {
  background-color: white;
  padding: 1.6rem;
  border-radius: 8px;
  text-align: center;
  box-shadow: 0 2px 10px rgba(0,0,0,0.06);
  transition: transform 0.2s ease;
  border: 1px solid #e9ecef;
  height: 100%;
  box-sizing: border-box;
}

.dark-theme .feature-item {
  background-color: #252525;
  border-color: #444;
}

.feature-item h4 {
  font-size: 1.15rem;
  color: var(--hopkins-blue);
  margin-bottom: 0.7rem;
  word-wrap: break-word;
}

.feature-item p {
  font-size: 0.95rem;
  color: var(--text-secondary);
  line-height: 1.45;
  word-wrap: break-word;
}

.feature-item:hover {
  transform: translateY(-4px);
  box-shadow: 0 4px 15px rgba(0,0,0,0.1);
}

.feature-icon-large {
  font-size: 45px;
  color: var(--hopkins-blue);
  margin-bottom: 1rem;
}

/* Process steps */
.process-steps {
  counter-reset: step-counter;
  margin: 1.3rem 0;
}

.process-step {
  position: relative;
  padding: 1.3rem 1.3rem 1.3rem 75px;
  margin-bottom: 1rem;
  background: white;
  border-radius: 8px;
  border-left: 3px solid var(--hopkins-gold);
  font-size: 1rem;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
  overflow: hidden;
}

.dark-theme .process-step {
  background: #252525;
}

.process-step h4 {
  font-size: 1.05rem;
  color: var(--hopkins-blue);
  margin-bottom: 0.35rem;
  word-wrap: break-word;
}

.process-step p {
  font-size: 0.95rem;
  color: var(--text-secondary);
  line-height: 1.45;
  word-wrap: break-word;
}

.process-step:before {
  counter-increment: step-counter;
  content: counter(step-counter);
  position: absolute;
  left: 1.3rem;
  top: 50%;
  transform: translateY(-50%);
  width: 48px;
  height: 48px;
  background: linear-gradient(135deg, var(--hopkins-blue) 0%, var(--hopkins-dark-blue) 100%);
  color: white;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: bold;
  font-size: 19px;
  box-shadow: 0 2px 6px rgba(0,0,0,0.15);
}

/* Developer card */
.developer-card {
  background: linear-gradient(135deg, var(--hopkins-blue) 0%, var(--hopkins-dark-blue) 100%);
  color: white;
  padding: 2.2rem;
  border-radius: 10px;
  text-align: center;
  margin: 1.3rem 0;
  box-shadow: 0 4px 15px rgba(0,0,0,0.1);
  width: 100%;
  box-sizing: border-box;
}

.developer-card h2 {
  font-size: 2rem;
  color: white;
  margin-bottom: 0.5rem;
  word-wrap: break-word;
}

.developer-card h4 {
  font-size: 1.15rem;
  color: var(--hopkins-gold);
  margin-bottom: 0.7rem;
  word-wrap: break-word;
}

.developer-card p {
  font-size: 1rem;
  color: rgba(255,255,255,0.9);
  line-height: 1.55;
  word-wrap: break-word;
}

/* Conversion section */
.conversion-section {
  background-color: white;
  padding: 1.6rem;
  border-radius: 8px;
  margin: 1.3rem 0;
  box-shadow: 0 2px 10px rgba(0,0,0,0.06);
  border-top: 3px solid var(--hopkins-accent);
  font-size: 0.95rem;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
  overflow: hidden;
}

/* Data preview container */
.data-preview-container {
  max-height: 550px;
  overflow-y: auto;
  border: 1px solid #e9ecef;
  border-radius: 6px;
  margin: 1rem 0;
  box-shadow: 0 1px 4px rgba(0,0,0,0.04);
  font-size: 0.95rem;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
}

/* Status messages */
.status-success {
  color: #155724;
  background-color: #d4edda;
  border-color: #c3e6cb;
  padding: 1rem;
  border-radius: 6px;
  border-left: 3px solid #28a745;
  margin: 0.8rem 0;
  font-size: 0.95rem;
  width: 100%;
  box-sizing: border-box;
  word-wrap: break-word;
}

.status-error {
  color: #721c24;
  background-color: #f8d7da;
  border-color: #f5c6cb;
  padding: 1rem;
  border-radius: 6px;
  border-left: 3px solid #dc3545;
  margin: 0.8rem 0;
  font-size: 0.95rem;
  width: 100%;
  box-sizing: border-box;
  word-wrap: break-word;
}

.status-waiting {
  color: #856404;
  background-color: #fff3cd;
  border-color: #ffeaa7;
  padding: 1rem;
  border-radius: 6px;
  border-left: 3px solid #ffc107;
  margin: 0.8rem 0;
  font-size: 0.95rem;
  width: 100%;
  box-sizing: border-box;
  word-wrap: break-word;
}

/* DataTables wrapper */
.dataTables_wrapper {
  width: 100% !important;
  margin: 0 auto;
  font-size: 0.95rem;
  box-sizing: border-box;
}

.dataTables_wrapper table {
  font-size: 0.95rem;
  width: 100% !important;
}

.shiny-output-error {
  color: #dc3545;
  padding: 1rem;
  border-radius: 6px;
  background-color: #f8d7da;
  margin: 0.8rem 0;
  font-size: 0.95rem;
  width: 100%;
  box-sizing: border-box;
  word-wrap: break-word;
}

/* Form controls */
.form-control {
  font-size: 1rem;
  padding: 9px 12px;
  background-color: white;
  color: var(--text-primary);
  border: 1px solid #ced4da;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
  max-width: 100%;
}

.dark-theme .form-control {
  background-color: #2A2A2A;
  color: var(--text-primary);
  border-color: #555;
}

.form-group {
  width: 100%;
  box-sizing: border-box;
}

.form-group label {
  font-size: 1rem;
  font-weight: 600;
  color: var(--text-primary);
  margin-bottom: 0.4rem;
  display: block;
  word-wrap: break-word;
}

/* File input */
.shiny-input-container {
  font-size: 1rem;
  width: 100%;
  box-sizing: border-box;
}

/* Copyright footer */
.copyright-footer {
  text-align: center;
  padding: 1.6rem;
  margin-top: 2.2rem;
  background-color: var(--hopkins-dark-gray);
  color: white;
  font-size: 0.95rem;
  border-top: 3px solid var(--hopkins-gold);
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
}

/* Welcome section */
.welcome-section {
  background: linear-gradient(135deg, var(--hopkins-light-blue) 0%, #ffffff 100%);
  padding: 2rem;
  border-radius: 10px;
  margin: 1.3rem 0;
  box-shadow: 0 4px 15px rgba(0,45,114,0.08);
  border-left: 4px solid var(--hopkins-blue);
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
}

.welcome-icon {
  font-size: 3.2rem;
  color: var(--hopkins-blue);
  margin-bottom: 0.9rem;
}

/* Guide section */
.guide-section {
  background-color: white;
  padding: 1.6rem;
  border-radius: 8px;
  margin: 1.1rem 0;
  border: 1px solid #e9ecef;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
}

.guide-step {
  display: flex;
  align-items: flex-start;
  margin-bottom: 1rem;
  padding: 1rem;
  background-color: var(--hopkins-gray);
  border-radius: 6px;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
}

.dark-theme .guide-step {
  background-color: #2A2A2A;
}

.guide-step:hover {
  background-color: #e9ecef;
  transform: translateX(4px);
}

.dark-theme .guide-step:hover {
  background-color: #333;
}

.step-number {
  background-color: var(--hopkins-blue);
  color: white;
  width: 38px;
  height: 38px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: bold;
  margin-right: 1rem;
  flex-shrink: 0;
  font-size: 0.95rem;
}

.step-content {
  flex: 1;
  min-width: 0;
}

.step-content h4 {
  font-size: 1.05rem;
  margin-bottom: 0.35rem;
}

.step-content p {
  font-size: 0.95rem;
  margin: 0;
}

/* Important note */
.important-note {
  background-color: #fff3cd;
  border-left: 4px solid #ffc107;
  padding: 1.3rem;
  border-radius: 6px;
  margin: 1.1rem 0;
  font-size: 1rem;
  transition: all 0.3s ease;
  width: 100%;
  box-sizing: border-box;
  word-wrap: break-word;
}

.dark-theme .important-note {
  background-color: #332701;
}

/* Column layout */
.data-converter-column {
  padding-right: 15px;
  padding-left: 15px;
  box-sizing: border-box;
}

/* Dark theme additional styles */
.dark-theme .dataTables_wrapper .dataTables_length,
.dark-theme .dataTables_wrapper .dataTables_filter,
.dark-theme .dataTables_wrapper .dataTables_info,
.dark-theme .dataTables_wrapper .dataTables_processing,
.dark-theme .dataTables_wrapper .dataTables_paginate {
  color: var(--text-primary) !important;
}

.dark-theme table.dataTable thead th,
.dark-theme table.dataTable thead td {
  border-bottom: 2px solid #555 !important;
  color: var(--text-primary) !important;
}

.dark-theme table.dataTable tbody tr {
  background-color: #252525 !important;
}

.dark-theme table.dataTable tbody tr:hover {
  background-color: #2A2A2A !important;
}

.dark-theme table.dataTable tbody td {
  border-top: 1px solid #444 !important;
  color: var(--text-primary) !important;
}

.dark-theme .selectize-input {
  background-color: #2A2A2A !important;
  color: var(--text-primary) !important;
  border-color: #555 !important;
}

.dark-theme .modal-content {
  background-color: #252525 !important;
  color: var(--text-primary) !important;
  border-color: #444 !important;
}

/* Responsive design */
@media (max-width: 992px) {
  .col-md-3, .col-md-9 {
    width: 100% !important;
    float: none !important;
  }
  
  .data-converter-column {
    padding-right: 0;
    padding-left: 0;
  }
  
  .container-fluid {
    transform: scale(0.99);
    width: 101.01%;
  }
}

@media (max-width: 768px) {
  html, body {
    font-size: 12.5px;
  }
  
  .blog-header {
    padding: 2.2rem 0 1.6rem 0;
  }
  
  .app-title-main {
    font-size: 2.4rem;
  }
  
  .app-subtitle-main {
    font-size: 1.3rem;
  }
  
  .app-tagline {
    font-size: 1.05rem;
  }
  
  .blog-card {
    padding: 1.3rem;
  }
  
  .blog-card h2 {
    font-size: 1.6rem;
  }
  
  .blog-card h3 {
    font-size: 1.3rem;
  }
  
  .feature-grid {
    grid-template-columns: 1fr;
    gap: 1rem;
  }
  
  .nav-tabs > li > a {
    padding: 10px 18px;
    font-size: 0.95rem;
  }
  
  .floating-app-name {
    font-size: 20px;
    padding: 6px 13px;
    top: 12px;
    left: 12px;
  }
  
  .theme-toggle-btn {
    width: 42px;
    height: 42px;
    font-size: 17px;
    top: 12px;
    right: 12px;
  }
  
  .main-content {
    padding: 0 1.2rem 1.2rem 1.2rem;
  }
  
  .container-fluid {
    transform: scale(0.995);
    width: 100.5%;
  }
}

@media (max-width: 480px) {
  html, body {
    font-size: 12px;
  }
  
  .app-title-main {
    font-size: 2rem;
  }
  
  .app-subtitle-main {
    font-size: 1.1rem;
  }
  
  .blog-card {
    padding: 1.1rem;
  }
  
  .btn-hopkins-primary, .btn-hopkins-success {
    font-size: 0.9rem;
    padding: 9px 16px;
  }
}

      
    "))
  ),
  
  # Floating App Name (always visible)
  #tags$a(href = "#", class = "floating-app-name", "Data2SPSS"),
  
  # Theme Toggle Button
  tags$button(
    id = "themeToggle",
    class = "theme-toggle-btn",
    title = "Toggle Dark/Light Theme",
    icon("moon")
  ),
  
  # Blog-inspired header (similar to Johns Hopkins)
  div(class = "blog-header",
      div(class = "blog-header-content",
          h1(class = "app-title-main", "Data2SPSS"),
          p(class = "app-subtitle-main", 
            "Professional Data Conversion Tool"),
          p(class = "app-tagline", 
            "Transform your research data between SPSS, Stata, Excel, CSV, SAS, RData, TSV, and TXT with ease.")
      )
  ),
  
  # Navigation and main content
  div(class = "container-fluid",
      # Tabset with horizontal tabs at the top
      tabsetPanel(
        id = "main_tabs",
        type = "tabs",
        
        # Tab 1: Home - ENHANCED WITH WELCOME SECTION
        tabPanel(
          "Home",
          icon = icon("home"),
          div(class = "main-content",
              
              # NEW: Welcome and Getting Started Section
              div(class = "welcome-section",
                  div(class = "welcome-header",
                      div(class = "welcome-icon", icon("hand-wave")),
                      h2("Welcome to Data2SPSS!"),
                      p("Thank you for using this tool. Before you start converting your data, please take a moment to read this home page to understand how the conversion process works.")
                  ),
                  
                  div(class = "guide-section",
                      h3("📋 How to Use This Tool Effectively"),
                      
                      div(class = "guide-step",
                          div(class = "step-number", "1"),
                          div(class = "step-content",
                              h4("Understand the Conversion Process"),
                              p("This tool automatically detects your variable types and converts them to SPSS-compatible format:"),
                              tags$ul(
                                tags$li("Continuous variables → Scale variables in SPSS"),
                                tags$li("Categorical variables → Nominal variables with value labels"),
                                tags$li("All variable names are cleaned for SPSS compatibility")
                              )
                          )
                      ),
                      
                      div(class = "guide-step",
                          div(class = "step-number", "2"),
                          div(class = "step-content",
                              h4("Prepare Your Data"),
                              p("For best results:"),
                              tags$ul(
                                tags$li("Ensure your data is clean and well-structured"),
                                tags$li("Remove any special characters from variable names"),
                                tags$li("Check that categorical variables have consistent values")
                              )
                          )
                      ),
                      
                      div(class = "guide-step",
                          div(class = "step-number", "3"),
                          div(class = "step-content",
                              h4("Use the Preview Features"),
                              p("Before converting, use the tabs in the Data Converter to:"),
                              tags$ul(
                                tags$li("Preview your original data"),
                                tags$li("Check variable classification"),
                                tags$li("Review SPSS metadata before download")
                              )
                          )
                      ),
                      
                      div(class = "important-note",
                          h4("💡 Important Note About the Tool"),
                          p("Data2SPSS also handles ALL major data formats. You can convert SPSS to Stata, Stata to Excel, SAS to CSV, and any combination. The tool intelligently preserves value labels, variable names, and metadata across all conversions. Additionally, numerical variables converted to 'Scale' type in SPSS will have values and labels, but these are for reference only. You can safely use these variables for inferential analysis as they will produce valid statistical results.")
                      )
                  ),
                  
                  div(style = "text-align: center; margin-top: 2rem;",
                      actionButton("go_to_converter", "🚀 Start Converting Your Data", 
                                   class = "btn-hopkins-primary btn-lg",
                                   icon = icon("arrow-right"))
                  )
              ),
              
              div(class = "blog-card",
                  h2("About Data2SPSS"),
                  p("This professional tool helps researchers and analysts convert their data files between multiple formats including SPSS, Stata, Excel, CSV, SAS, and R files. Designed with academic research in mind, the tool ensures your data is analysis-ready in any environment."),
                  
                  h3("Key Features"),
                  div(class = "feature-grid",
                      div(class = "feature-item",
                          div(class = "feature-icon-large", icon("file-import")),
                          h4("Multi-Format Import"),
                          p("Import from CSV, Excel, Stata, SAS, SPSS, RData, and more with seamless compatibility")
                      ),
                      div(class = "feature-item",
                          div(class = "feature-icon-large", icon("exchange-alt")),
                          h4("Bidirectional Conversion"),
                          p("Convert between any supported formats: SPSS ↔ Stata, Excel ↔ SAS, and all combinations")
                      ),
                      div(class = "feature-item",
                          div(class = "feature-icon-large", icon("magic")),
                          h4("Smart Conversion"),
                          p("Automatic detection of variable types and intelligent value label preservation across all formats")
                      ),
                      div(class = "feature-item",
                          div(class = "feature-icon-large", icon("chart-bar")),
                          h4("Data Analysis"),
                          p("Comprehensive variable analysis and metadata preview before export")
                      )
                  ),
                  
                  h3("Quick Steps for Converting to SPSS"),
                  div(class = "process-steps",
                      div(class = "process-step",
                          h4("Import Your Data"),
                          p("Start by importing your data file in the 'Data Converter' tab. Supported formats include: Stata, Excel, CSV, SAS, RData, and more.")
                      ),
                      
                      div(class = "process-step",
                          h4("Convert to SPSS Format"),
                          p("Click the 'Convert to SPSS Format' button to automatically convert your data with proper variable types and value labels.")
                      ),
                      
                      div(class = "process-step",
                          h4("Download SPSS File"),
                          p("Download the converted data as an SPSS (.sav) file ready for statistical analysis.")
                      ),
                      
                      div(class = "process-step",
                          h4("Finalize in SPSS"),
                          p("Open the downloaded file in SPSS Statistics, go to Variable View, change categorical variables Measure to Nominal/Ordinal where appropriate, and save.")
                      )
                  )
              )
          )
        ),
        
        # Tab 2: Data Converter (SPSS Conversion)
        tabPanel(
          "Data Converter",
          icon = icon("file-export"),
          div(class = "main-content",
              fluidRow(
                column(
                  width = 3,
                  class = "data-converter-column",
                  div(class = "blog-card",
                      h4("📁 Data Import"),
                      # CRITICAL FIX: Added accept attribute for all file types including SAS xpt
                      fileInput("file_input", "Choose Data File",
                                accept = c(".csv", ".xlsx", ".xls", ".dta", ".sav", ".sas7bdat", ".xpt",
                                           ".rdata", ".tsv", ".txt", ".jmp", ".por"),
                                multiple = FALSE,
                                width = "100%"),
                      
                      conditionalPanel(
                        condition = "output.file_loaded",
                        hr(),
                        h4("📊 Data Information"),
                        uiOutput("data_info_box"),
                        
                        hr(),
                        h4("🔄 SPSS Conversion"),
                        textInput("export_filename", "Export Filename", 
                                  value = "converted_data.sav",
                                  width = "100%"),
                        
                        # FIXED: Button with proper width handling
                        div(style = "width: 100%; box-sizing: border-box;",
                            actionButton("convert_btn", "🔧 Convert to SPSS Format", 
                                         class = "btn-hopkins-primary",
                                         style = "width: 100%; margin-bottom: 10px; padding: 12px 10px;")
                        ),
                        
                        conditionalPanel(
                          condition = "output.conversion_complete",
                          div(class = "status-success",
                              style = "width: 100%; box-sizing: border-box;",
                              h5("✅ Conversion Successful!", style = "margin: 0;"),
                              p("Data is ready for download.", style = "margin: 5px 0 0 0; font-size: 14px;")
                          ),
                          # FIXED: Download button with proper width
                          div(style = "width: 100%; box-sizing: border-box;",
                              downloadButton("download_btn", "📥 Download SPSS File", 
                                             class = "btn-hopkins-success shiny-download-link",
                                             style = "width: 100%; margin-top: 10px; padding: 12px 10px;")
                          )
                        ),
                        
                        conditionalPanel(
                          condition = "output.conversion_status == 'error'",
                          div(class = "status-error",
                              style = "width: 100%; box-sizing: border-box;",
                              h5("❌ Conversion Failed"),
                              textOutput("conversion_error")
                          )
                        )
                      )
                  )
                ),
                
                column(
                  width = 9,
                  class = "data-converter-column",
                  conditionalPanel(
                    condition = "!output.file_loaded",
                    div(class = "blog-card",
                        h3("Ready to Convert Your Data"),
                        p("Upload your data file to begin the conversion process. The tool supports multiple file formats and will automatically detect variable types."),
                        div(class = "step-box",
                            h4("Supported Input Formats:"),
                            tags$ul(
                              tags$li("CSV, TSV, TXT files"),
                              tags$li("Excel files (.xlsx, .xls)"),
                              tags$li("Stata files (.dta)"),
                              tags$li("SPSS files (.sav, .por)"),
                              tags$li("SAS files (.sas7bdat, .xpt - SAS JMP compatible)"),
                              tags$li("R Data files (.rdata)")
                            )
                        ),
                        div(class = "info-box",
                            h4("💡 Important:"),
                            p("There is NO FILE SIZE LIMIT. You can upload files of any size, including datasets with trillions of rows (subject to your system's available memory).")
                        )
                    )
                  ),
                  
                  conditionalPanel(
                    condition = "output.file_loaded",
                    div(class = "conversion-section",
                        tabsetPanel(
                          id = "converter_tabs",
                          tabPanel("Data Preview", 
                                   div(class = "data-preview-container",
                                       DTOutput("data_preview")
                                   )),
                          tabPanel("Data Structure", 
                                   DTOutput("data_structure_table")),
                          tabPanel("Variable Analysis", 
                                   h4("Variable Classification"),
                                   p("This shows how each variable will be handled:"),
                                   DTOutput("variable_analysis_table")),
                          tabPanel("SPSS Metadata", 
                                   conditionalPanel(
                                     condition = "output.conversion_complete",
                                     h4("SPSS Export Information"),
                                     p("This shows exactly what will be exported to SPSS:"),
                                     DTOutput("spss_metadata_table")
                                   ),
                                   conditionalPanel(
                                     condition = "!output.conversion_complete",
                                     div(class = "status-waiting",
                                         h4("Conversion Required"),
                                         p("Click the 'Convert to SPSS Format' button first.")
                                     )
                                   ))
                        )
                    )
                  )
                )
              )
          )
        ),
        
        # Tab 3: Export to Other Formats (Bidirectional Conversion) - UPDATED FOR SAS OPTIONS
        tabPanel(
          "Export to Other Formats",
          icon = icon("exchange-alt"),
          div(class = "main-content",
              fluidRow(
                column(
                  width = 3,
                  class = "data-converter-column",
                  div(class = "blog-card",
                      h4("📤 Export Settings"),
                      
                      conditionalPanel(
                        condition = "output.file_loaded",
                        hr(),
                        h4("📊 Loaded Data"),
                        uiOutput("export_data_info_box"),
                        
                        hr(),
                        h4("🔄 Export Options"),
                        # UPDATED: Added SAS export options with clear JMP compatibility note
                        selectInput("target_format", "Select Output Format:",
                                    choices = c(
                                      "CSV" = "csv",
                                      "Excel" = "xlsx", 
                                      "Stata" = "dta",
                                      "R Data" = "rdata",
                                      "TSV" = "tsv",
                                      "Text (Tab)" = "txt",
                                      "SAS (.sas7bdat - Standard SAS)" = "sas7bdat",
                                      "SAS (.xpt - JMP Compatible)" = "xpt"
                                    ),
                                    selected = "csv",
                                    width = "100%"),
                        textInput("export_filename_other", "Export Filename", 
                                  value = "exported_data",
                                  width = "100%"),
                        
                        div(style = "width: 100%; box-sizing: border-box;",
                            actionButton("convert_other_btn", "🔧 Convert & Export", 
                                         class = "btn-hopkins-primary",
                                         style = "width: 100%; margin-bottom: 10px; padding: 12px 10px;")
                        ),
                        
                        conditionalPanel(
                          condition = "output.other_conversion_complete",
                          div(class = "status-success",
                              style = "width: 100%; box-sizing: border-box;",
                              h5("✅ Export Ready!", style = "margin: 0;"),
                              p("File is ready for download.", style = "margin: 5px 0 0 0; font-size: 14px;")
                          ),
                          div(style = "width: 100%; box-sizing: border-box;",
                              downloadButton("download_other_btn", "📥 Download File", 
                                             class = "btn-hopkins-success shiny-download-link",
                                             style = "width: 100%; margin-top: 10px; padding: 12px 10px;")
                          )
                        ),
                        
                        conditionalPanel(
                          condition = "output.other_conversion_status == 'error'",
                          div(class = "status-error",
                              style = "width: 100%; box-sizing: border-box;",
                              h5("❌ Export Failed"),
                              textOutput("other_conversion_error")
                          )
                        )
                      ),
                      
                      conditionalPanel(
                        condition = "!output.file_loaded",
                        div(class = "status-waiting",
                            h4("No Data Loaded"),
                            p("Please go to the Data Converter tab first to load your data file.")
                        )
                      )
                  )
                ),
                
                column(
                  width = 9,
                  class = "data-converter-column",
                  conditionalPanel(
                    condition = "output.file_loaded",
                    div(class = "conversion-section",
                        h3("Export Data to Other Formats"),
                        p("Convert your loaded data (which may be in SPSS, Stata, Excel, or other formats) to any of the supported output formats."),
                        
                        div(class = "info-box",
                            h4("📋 Supported Export Formats:"),
                            tags$ul(
                              tags$li("CSV - Comma separated values (universal format)"),
                              tags$li("Excel - Microsoft Excel (.xlsx)"),
                              tags$li("Stata - Stata dataset (.dta)"),
                              tags$li("R Data - R data file (.rdata)"),
                              tags$li("SAS (.sas7bdat) - Standard SAS dataset"),
                              tags$li("SAS (.xpt) - SAS Transport file (JMP compatible)"),
                              tags$li("TSV/TXT - Tab-separated values")
                            )
                        ),
                        
                        div(class = "step-box",
                            h4("💡 How It Works:"),
                            p("The tool automatically handles labelled variables from SPSS, Stata, and other formats, converting them to appropriate formats (factors with labels) for export. All value labels and variable names are preserved where possible."),
                            p("For SAS export: You have two options - .sas7bdat (standard SAS dataset) or .xpt (SAS Transport file, which is JMP compatible). The .xpt format is recommended for opening in SAS JMP and other SAS software that may not support .sas7bdat directly."),
                            p("If you encounter any errors during export, the tool will automatically try to fix them by cleaning variable names, removing illegal characters, and converting data types as needed. For very large datasets (500,000+ rows), the tool optimizes memory usage to prevent crashes.")
                        ),
                        
                        hr(),
                        h4("Data Preview for Export"),
                        div(class = "data-preview-container",
                            DTOutput("export_data_preview")
                        ),
                        
                        h4("Variable Export Information"),
                        div(class = "data-preview-container",
                            DTOutput("export_variable_info")
                        )
                    )
                  ),
                  
                  conditionalPanel(
                    condition = "!output.file_loaded",
                    div(class = "blog-card",
                        h3("No Data Loaded"),
                        p("Please load your data first using the Data Converter tab."),
                        div(style = "text-align: center; margin-top: 1rem;",
                            actionButton("go_to_converter_from_export", "Go to Data Converter", 
                                         class = "btn-hopkins-primary")
                        )
                    )
                  )
                )
              )
          )
        ),
        
        # Tab 4: Developer Information (LAST TAB)
        tabPanel(
          "Developer",
          icon = icon("user-tie"),
          div(class = "main-content",
              div(class = "developer-card",
                  h2("Mudasir Mohammed Ibrahim"),
                  h4("Open-Source Advocate"),
                  p("Passionate about creating tools that bridge the gap between data and statistical analysis."),
                  
                  br(),
                  div(class = "academic-links",
                      style = "margin-top: 15px; padding-top: 15px; border-top: 1px solid #e0e0e0;",
                      h5("Connect with me on academic networks:", style = "margin-bottom: 15px; color: #1A5276;"),
                      div(style = "display: flex; justify-content: center; gap: 25px; flex-wrap: wrap;",
                          
                          # ResearchGate
                          a(href = "https://www.researchgate.net/profile/Mudasir-Ibrahim", 
                            target = "_blank", 
                            title = "ResearchGate Profile",
                            style = "color: #2d3748; text-decoration: none; display: inline-flex; flex-direction: column; align-items: center; transition: transform 0.3s ease;",
                            onmouseover = "this.style.transform='translateY(-3px)'",
                            onmouseout = "this.style.transform='translateY(0)'",
                            div(style = "background: #00ccbb; width: 50px; height: 50px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 5px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);",
                                icon("researchgate", style = "font-size: 24px; color: white;")
                            ),
                            div(style = "font-size: 11px; font-weight: 500;", "")
                          ),
                          
                          # Google Scholar
                          a(href = "https://scholar.google.com/citations?user=xEFzAvgAAAAJ&hl=en", 
                            target = "_blank", 
                            title = "Google Scholar Profile",
                            style = "color: #2d3748; text-decoration: none; display: inline-flex; flex-direction: column; align-items: center; transition: transform 0.3s ease;",
                            onmouseover = "this.style.transform='translateY(-3px)'",
                            onmouseout = "this.style.transform='translateY(0)'",
                            div(style = "background: #4285f4; width: 50px; height: 50px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 5px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);",
                                icon("graduation-cap", style = "font-size: 24px; color: white;")
                            ),
                            div(style = "font-size: 11px; font-weight: 500;", "")
                          ),
                          
                          # Web of Science
                          a(href = "https://www.webofscience.com/wos/author/record/HPC-2085-2023", 
                            target = "_blank", 
                            title = "Web of Science Profile",
                            style = "color: #2d3748; text-decoration: none; display: inline-flex; flex-direction: column; align-items: center; transition: transform 0.3s ease;",
                            onmouseover = "this.style.transform='translateY(-3px)'",
                            onmouseout = "this.style.transform='translateY(0)'",
                            div(style = "background: #ff6b6b; width: 50px; height: 50px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 5px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);",
                                icon("book", style = "font-size: 24px; color: white;")
                            ),
                            div(style = "font-size: 11px; font-weight: 500;", "")
                          ),
                          
                          # ORCID
                          a(href = "https://orcid.org/0000-0002-9049-8222", 
                            target = "_blank", 
                            title = "ORCID Profile",
                            style = "color: #2d3748; text-decoration: none; display: inline-flex; flex-direction: column; align-items: center; transition: transform 0.3s ease;",
                            onmouseover = "this.style.transform='translateY(-3px)'",
                            onmouseout = "this.style.transform='translateY(0)'",
                            div(style = "background: #a6ce39; width: 50px; height: 50px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 5px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);",
                                icon("orcid", style = "font-size: 24px; color: white;")
                            ),
                            div(style = "font-size: 11px; font-weight: 500;", "")
                          )
                      )
                  )
              ),
              
              # About This Application section - NOW INSIDE THE DEVELOPER TAB
              div(class = "blog-card",
                  h3("About This Application"),
                  p("This Data2SPSS web application was developed to address the common challenges researchers face when preparing data for statistical analysis across different platforms. The tool automates the tedious process of format conversion, variable type detection, and value labeling."),
                  
                  h4("Technical Features"),
                  div(class = "step-box",
                      tags$ul(
                        tags$li("Bidirectional conversion between all major statistical formats"),
                        tags$li("Smart variable type detection (continuous vs categorical)"),
                        tags$li("Automatic value label generation and preservation"),
                        tags$li("Support for multiple input and output formats"),
                        tags$li("Interactive data preview and analysis"),
                        tags$li("Professional academic-inspired interface design"),
                        tags$li("Handles datasets of any size (memory permitting)"),
                        tags$li("Full SAS JMP support with .xpt file export"),
                        tags$li("Robust error handling - automatically fixes illegal characters in variable names and data values"),
                        tags$li("Optimized for large datasets (500,000+ rows) with memory-efficient processing")
                      )
                  ),
                  
                  h4("Contact"),
                  p("I am passionate about creating apps that streamline research processes and enhance data analysis workflows. Feel free to reach out with any questions, suggestions, or collaborations"),
                  
                  div(style = "text-align: center; margin-top: 2rem;",
                      actionButton("contact_btn", "Contact Me", 
                                   class = "btn-hopkins-primary",
                                   icon = icon("envelope"),
                                   onclick = "window.open('mailto:mudassiribrahim30@gmail.com')")
                  )
              )
          )
        )
      )
  ),
  
  # Copyright footer with dynamic year
  div(class = "copyright-footer",
      p(paste("Copyright © 2025-", format(Sys.Date(), "%Y"), "Data2SPSS")),
      p("No rights reserved. This app is free to use and distribute.")
  )
)

# Server Logic with enhanced error handling
# CRITICAL FIX: Set maximum upload size to unlimited
options(shiny.maxRequestSize = Inf)

server <- function(input, output, session) {
  
  # Reactive values
  data_loaded <- reactiveVal(NULL)
  data_summary <- reactiveVal(NULL)
  spss_data <- reactiveVal(NULL)
  conversion_status <- reactiveVal("not_started")
  conversion_error <- reactiveVal("")
  original_names <- reactiveVal(NULL)
  cleaned_names <- reactiveVal(NULL)
  
  # New reactive values for bidirectional conversion
  other_conversion_status <- reactiveVal("not_started")
  other_conversion_error <- reactiveVal("")
  export_ready_data <- reactiveVal(NULL)
  current_target_format <- reactiveVal("csv")
  current_export_filename <- reactiveVal("exported_data")
  
  # Navigate to Data Converter tab when Get Started is clicked
  observeEvent(input$go_to_converter, {
    updateTabsetPanel(session, "main_tabs", selected = "Data Converter")
  })
  
  # Navigate from export tab to data converter
  observeEvent(input$go_to_converter_from_export, {
    updateTabsetPanel(session, "main_tabs", selected = "Data Converter")
  })
  
  # Enhanced data loading with better error handling
  observeEvent(input$file_input, {
    req(input$file_input)
    
    tryCatch({
      showModal(modalDialog(
        div(
          style = "text-align: center;",
          h4("Loading Data..."),
          div(class = "loading-spinner"),
          p("Processing file format..."),
          p("Note: File size may be large. Please wait.")
        ), 
        footer = NULL
      ))
      
      conversion_status("not_started")
      other_conversion_status("not_started")
      spss_data(NULL)
      export_ready_data(NULL)
      
      file_ext <- tolower(tools::file_ext(input$file_input$name))
      data <- load_data_file(input$file_input$datapath, file_ext)
      
      # Validate data was loaded correctly
      if (is.null(data) || nrow(data) == 0 || ncol(data) == 0) {
        stop("The file appears to be empty or could not be read properly")
      }
      
      original_names(names(data))
      cleaned_names(clean_variable_names(names(data), "spss"))
      
      data_loaded(data)
      data_summary(get_data_summary(data))
      
      removeModal()
      
      # Show variable classification summary
      continuous_count <- sum(sapply(data, is_continuous_numeric))
      categorical_count <- ncol(data) - continuous_count
      
      showNotification(
        paste("Successfully loaded data with", format(nrow(data), big.mark = ","), "rows and", ncol(data), "columns"),
        type = "message", duration = 6
      )
      
      showNotification(
        paste("Detected:", continuous_count, "continuous,", categorical_count, "categorical variables"),
        type = "message", duration = 6
      )
      
    }, error = function(e) {
      removeModal()
      conversion_status("error")
      conversion_error(e$message)
      showNotification(paste("Error loading file:", e$message), type = "error")
    })
  })
  
  # Convert to SPSS format with enhanced validation
  observeEvent(input$convert_btn, {
    req(data_loaded())
    
    tryCatch({
      showModal(modalDialog(
        div(
          style = "text-align: center;",
          h4("Converting to SPSS Format..."),
          div(class = "loading-spinner"),
          p("Setting proper measure types (nominal for categorical, scale for continuous)..."),
          p("This may take a moment for large datasets...")
        ), 
        footer = NULL
      ))
      
      converted_data <- convert_to_spss(data_loaded())
      
      # Validate conversion
      if (is.null(converted_data) || ncol(converted_data) == 0) {
        stop("Conversion resulted in empty dataset")
      }
      
      # Validate variable names
      final_names <- names(converted_data)
      if (any(!grepl("^[a-zA-Z][a-zA-Z0-9_]*$", final_names))) {
        warning("Some variable names may not be SPSS-compatible after conversion")
      }
      
      spss_data(converted_data)
      conversion_status("complete")
      conversion_error("")
      
      removeModal()
      
      # Count final results from ACTUAL converted data
      spss_data_val <- spss_data()
      nominal_count <- sum(sapply(spss_data_val, function(x) {
        measure <- attr(x, "SPSS.measure")
        !is.null(measure) && measure == "nominal"
      }))
      scale_count <- ncol(spss_data_val) - nominal_count
      
      showNotification(
        paste("✅ Conversion successful! All variable names validated for SPSS."),
        type = "message", duration = 8
      )
      
      showNotification(
        paste(nominal_count, "variables set as NOMINAL (with labels),", scale_count, "as SCALE (without labels)"),
        type = "message", duration = 6
      )
      
    }, error = function(e) {
      removeModal()
      conversion_status("error")
      conversion_error(e$message)
      showNotification(paste("❌ Conversion failed:", e$message), type = "error")
    })
  })
  
  # NEW: Convert to other format (bidirectional) - UPDATED for SAS options with robust handling
  observeEvent(input$convert_other_btn, {
    req(data_loaded())
    target <- input$target_format
    filename_base <- input$export_filename_other
    
    current_target_format(target)
    current_export_filename(filename_base)
    
    tryCatch({
      showModal(modalDialog(
        div(
          style = "text-align: center;",
          h4(paste("Preparing data for", toupper(target), "export...")),
          div(class = "loading-spinner"),
          p("Cleaning variable names and data values..."),
          p("This may take a moment for large datasets...")
        ), 
        footer = NULL
      ))
      
      # Pre-clean data before storing
      export_data <- clean_data_for_export(data_loaded(), target)
      export_ready_data(export_data)
      other_conversion_status("complete")
      other_conversion_error("")
      
      removeModal()
      
      showNotification(
        paste("✅ Data prepared for export as", toupper(target), "! Click download to save."),
        type = "message", duration = 6
      )
      
    }, error = function(e) {
      removeModal()
      other_conversion_status("error")
      other_conversion_error(e$message)
      showNotification(paste("❌ Export preparation failed:", e$message), type = "error")
    })
  })
  
  # Download handler for other formats - UPDATED for SAS .sas7bdat and .xpt using haven with robust handling
  output$download_other_btn <- downloadHandler(
    filename = function() {
      base <- current_export_filename()
      ext <- current_target_format()
      # Handle SAS extensions properly
      if (ext == "sas7bdat") {
        if (!grepl("\\.sas7bdat$", base)) {
          base <- paste0(base, ".sas7bdat")
        }
      } else if (ext == "xpt") {
        if (!grepl("\\.xpt$", base)) {
          base <- paste0(base, ".xpt")
        }
      } else if (ext == "sas") {
        # Legacy support
        if (!grepl("\\.sas7bdat$", base)) {
          base <- paste0(base, ".sas7bdat")
        }
      } else {
        if (!grepl(paste0("\\.", ext, "$"), base)) {
          base <- paste0(base, ".", ext)
        }
      }
      base
    },
    content = function(file) {
      req(export_ready_data())
      target <- current_target_format()
      
      tryCatch({
        showModal(modalDialog(
          div(
            style = "text-align: center;",
            h4(paste("Saving", toupper(target), "file...")),
            div(class = "loading-spinner"),
            p("Please wait while your data is being exported..."),
            p("For large datasets, this may take several minutes.")
          ), 
          footer = NULL
        ))
        
        # For SAS, we need to ensure proper filename
        if (target == "sas7bdat") {
          # Ensure .sas7bdat extension
          if (!grepl("\\.sas7bdat$", file)) {
            file <- gsub("\\.sas$", ".sas7bdat", file)
            if (!grepl("\\.sas7bdat$", file)) {
              file <- paste0(file, ".sas7bdat")
            }
          }
        } else if (target == "xpt") {
          # Ensure .xpt extension
          if (!grepl("\\.xpt$", file)) {
            file <- paste0(file, ".xpt")
          }
        }
        
        success <- convert_to_other_format(export_ready_data(), target, file)
        
        removeModal()
        showNotification(paste("✅ File downloaded successfully as", toupper(target)), 
                         type = "message", duration = 5)
        
      }, error = function(e) {
        removeModal()
        showNotification(paste("❌ Download failed:", e$message), type = "error", duration = 10)
        # Log error for debugging
        warning(paste("Export error for format", target, ":", e$message))
      })
    }
  )
  
  # Output controls
  output$file_loaded <- reactive({ !is.null(data_loaded()) })
  output$conversion_complete <- reactive({ conversion_status() == "complete" })
  output$conversion_status <- reactive({ conversion_status() })
  output$conversion_error <- renderText({ conversion_error() })
  
  # New output controls for other formats
  output$other_conversion_complete <- reactive({ other_conversion_status() == "complete" })
  output$other_conversion_status <- reactive({ other_conversion_status() })
  output$other_conversion_error <- renderText({ other_conversion_error() })
  
  outputOptions(output, "file_loaded", suspendWhenHidden = FALSE)
  outputOptions(output, "conversion_complete", suspendWhenHidden = FALSE)
  outputOptions(output, "conversion_status", suspendWhenHidden = FALSE)
  outputOptions(output, "other_conversion_complete", suspendWhenHidden = FALSE)
  outputOptions(output, "other_conversion_status", suspendWhenHidden = FALSE)
  
  # Data information box
  output$data_info_box <- renderUI({
    req(data_loaded())
    data <- data_loaded()
    
    # Get accurate counts from actual data
    continuous_count <- sum(sapply(data, is_continuous_numeric))
    categorical_count <- ncol(data) - continuous_count
    
    # Get accurate conversion status from actual converted data
    if (conversion_status() == "complete") {
      spss_data_val <- spss_data()
      nominal_count <- sum(sapply(spss_data_val, function(x) {
        measure <- attr(x, "SPSS.measure")
        !is.null(measure) && measure == "nominal"
      }))
      scale_count <- ncol(spss_data_val) - nominal_count
      
      status_html <- paste0(
        "<p><strong>Status:</strong> <span style='color: green;'>✓ Converted to SPSS</span></p>",
        "<p><strong>Nominal Variables:</strong> ", nominal_count, " (with labels)</p>",
        "<p><strong>Scale Variables:</strong> ", scale_count, " (without labels)</p>"
      )
    } else if (conversion_status() == "error") {
      status_html <- "<p><strong>Status:</strong> <span style='color: red;'>✗ Conversion Failed</span></p>"
    } else {
      status_html <- "<p><strong>Status:</strong> <span style='color: orange;'>⏳ Ready for Conversion</span></p>"
    }
    
    div(
      class = "info-box",
      HTML(paste0(
        "<p><strong>Dataset:</strong> ", input$file_input$name, "</p>",
        "<p><strong>Rows:</strong> ", format(nrow(data), big.mark = ","), "</p>",
        "<p><strong>Columns:</strong> ", ncol(data), "</p>",
        "<p><strong>File Size:</strong> ", format(file.info(input$file_input$datapath)$size, units = "auto", big.mark = ","), "</p>",
        "<p><strong>Continuous Variables:</strong> ", continuous_count, " (will be SCALE - no labels)</p>",
        "<p><strong>Categorical Variables:</strong> ", categorical_count, " (will be NOMINAL - with labels)</p>",
        status_html
      ))
    )
  })
  
  # Export data information box
  output$export_data_info_box <- renderUI({
    req(data_loaded())
    data <- data_loaded()
    
    continuous_count <- sum(sapply(data, is_continuous_numeric))
    categorical_count <- ncol(data) - continuous_count
    
    div(
      class = "info-box",
      HTML(paste0(
        "<p><strong>Dataset:</strong> ", input$file_input$name, "</p>",
        "<p><strong>Rows:</strong> ", format(nrow(data), big.mark = ","), "</p>",
        "<p><strong>Columns:</strong> ", ncol(data), "</p>",
        "<p><strong>Continuous Variables:</strong> ", continuous_count, "</p>",
        "<p><strong>Categorical Variables:</strong> ", categorical_count, "</p>",
        "<p><strong>Ready to Export:</strong> Yes</p>",
        "<p><strong>Note:</strong> For SAS XPT export, variable names will be cleaned to alphanumeric only (max 32 chars)</p>"
      ))
    )
  })
  
  # Data preview - Shows value labels instead of numeric codes
  output$data_preview <- renderDT({
    req(data_loaded())
    
    # Convert labelled variables to show labels
    data_display <- as.data.frame(lapply(data_loaded(), function(x) {
      if (inherits(x, "haven_labelled")) {
        haven::as_factor(x)
      } else {
        x
      }
    }))
    
    datatable(
      data_display,
      options = list(
        scrollX = TRUE, 
        pageLength = 10, 
        dom = 'tip',
        language = list(
          emptyTable = "No data available",
          zeroRecords = "No matching records found"
        )
      ),
      rownames = FALSE,
      escape = FALSE,
      selection = 'none'
    )
  })
  
  # Export data preview - Shows value labels instead of numeric codes
  output$export_data_preview <- renderDT({
    req(data_loaded())
    
    # Show a sample for large datasets
    display_data <- data_loaded()
    if (nrow(display_data) > 10000) {
      display_data <- head(display_data, 10000)
      showNotification("Showing first 10,000 rows of large dataset", type = "warning", duration = 3)
    }
    
    # Convert labelled variables to show labels
    data_display <- as.data.frame(lapply(display_data, function(x) {
      if (inherits(x, "haven_labelled")) {
        haven::as_factor(x)
      } else {
        x
      }
    }))
    
    datatable(
      data_display,
      options = list(
        scrollX = TRUE, 
        pageLength = 8, 
        dom = 'tip',
        language = list(
          emptyTable = "No data available",
          zeroRecords = "No matching records found"
        )
      ),
      rownames = FALSE,
      escape = FALSE,
      selection = 'none'
    )
  })
  
  # Export variable information
  output$export_variable_info <- renderDT({
    req(data_loaded())
    data <- data_loaded()
    
    info_df <- data.frame(
      Variable = names(data),
      Cleaned_Name = clean_variable_names(names(data), ifelse(input$target_format %in% c("sas7bdat", "xpt"), "sas", "spss")),
      Original_Type = sapply(data, function(x) {
        if (inherits(x, "haven_labelled")) "SPSS/Stata Labelled"
        else if (is.factor(x)) "Factor"
        else if (is.character(x)) "Character"
        else if (is.numeric(x)) "Numeric"
        else class(x)[1]
      }),
      Unique_Values = sapply(data, function(x) length(unique(na.omit(x)))),
      Will_Export_As = sapply(data, function(x) {
        if (inherits(x, "haven_labelled") || !is.null(attr(x, "labels"))) {
          "Factor with labels"
        } else if (is.factor(x)) {
          "Factor"
        } else if (is.character(x)) {
          "Character"
        } else if (is.numeric(x)) {
          "Numeric"
        } else {
          "Character (converted)"
        }
      }),
      Has_Labels = sapply(data, function(x) {
        if (inherits(x, "haven_labelled")) "YES"
        else if (!is.null(attr(x, "labels"))) "YES"
        else if (is.factor(x)) "YES (factor levels)"
        else "NO"
      }),
      stringsAsFactors = FALSE
    )
    
    datatable(info_df, 
              options = list(scrollX = TRUE, pageLength = 15, dom = 'tip')) %>%
      formatStyle(
        'Has_Labels',
        backgroundColor = styleEqual(c('YES', 'NO'), c('#d4edda', '#f8d7da'))
      )
  })
  
  # Data structure table
  output$data_structure_table <- renderDT({
    req(data_summary())
    datatable(data_summary(), 
              options = list(scrollX = TRUE, pageLength = 25, dom = 'tip'))
  })
  
  # Variable analysis table
  output$variable_analysis_table <- renderDT({
    req(data_loaded())
    data <- data_loaded()
    
    analysis_df <- data.frame(
      Variable = names(data),
      Type = sapply(data, function(x) paste(class(x), collapse = ", ")),
      Unique_Values = sapply(data, function(x) length(unique(na.omit(x)))),
      Min = sapply(data, function(x) if(is.numeric(x)) min(x, na.rm = TRUE) else NA),
      Max = sapply(data, function(x) if(is.numeric(x)) max(x, na.rm = TRUE) else NA),
      Is_Continuous = sapply(data, is_continuous_numeric),
      SPSS_Measure = sapply(data, get_spss_measure_type),
      Will_Get_Labels = sapply(data, function(x) if(is_continuous_numeric(x)) "NO" else "YES"),
      stringsAsFactors = FALSE
    )
    
    datatable(analysis_df, 
              options = list(scrollX = TRUE, pageLength = 25, dom = 'tip')) %>%
      formatStyle(
        'Will_Get_Labels',
        backgroundColor = styleEqual(c('YES', 'NO'), c('#d4edda', '#f8d7da'))
      )
  })
  
  # SPSS Metadata table
  output$spss_metadata_table <- renderDT({
    req(spss_data())
    metadata <- get_spss_metadata(spss_data())
    datatable(metadata, 
              options = list(scrollX = TRUE, pageLength = 25, dom = 'tip')) %>%
      formatStyle(
        'Has_Value_Labels',
        backgroundColor = styleEqual(c('YES', 'NO'), c('#d4edda', '#f8d7da'))
      ) %>%
      formatStyle(
        'SPSS_Measure',
        backgroundColor = styleEqual(c('nominal', 'scale'), c('#d4edda', '#d1ecf1'))
      )
  })
  
  # Enhanced download handler with robust validation
  output$download_btn <- downloadHandler(
    filename = function() { 
      # Ensure .sav extension
      name <- input$export_filename
      if (!grepl("\\.sav$", name)) {
        name <- paste0(name, ".sav")
      }
      name
    },
    content = function(file) {
      req(spss_data())
      
      tryCatch({
        showModal(modalDialog(
          div(
            style = "text-align: center;",
            h4("Saving SPSS File..."),
            p("Validating variable names and setting proper measure types...")
          ), 
          footer = NULL
        ))
        
        # Final validation before export
        final_data <- spss_data()
        final_names <- names(final_data)
        
        # Ensure all names are valid for SPSS
        valid_names <- clean_variable_names(final_names, "spss")
        if (!identical(final_names, valid_names)) {
          names(final_data) <- valid_names
          showNotification("Some variable names were adjusted for SPSS compatibility", 
                           type = "warning", duration = 4)
        }
        
        write_spss_with_attributes(final_data, file)
        removeModal()
        
        showNotification("✅ SPSS file downloaded successfully! All variables validated.", 
                         type = "message", duration = 5)
        
      }, error = function(e) {
        removeModal()
        showNotification(paste("❌ Download failed:", e$message), type = "error")
      })
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
