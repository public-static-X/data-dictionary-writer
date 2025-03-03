checkNecessaryPackages <- function(required_packages = c("haven", "openxlsx", "svDialogs", "moments")) {
  inst <- installed.packages()[required_packages, "Version"]
  if (!all(required_packages %in% inst)) {
      install.packages(setdiff(required_packages, rownames(inst)))
    }
  aval <- available.packages()[required_packages, "Version"]
  
  pkg_table <- data.frame(Installed = inst, Available = aval, row.names = required_packages)
  
  # Find packages that need updating
  pkg_to_update <- required_packages[inst != aval]
  
  # Print update messages
  if (length(pkg_to_update) > 0) {
    for (pkg in pkg_to_update) {
      message(pkg, " needs to be updated from version ", inst[pkg], " to ", aval[pkg] )
    }
    install_cmd <- paste0("install.packages(c(", paste0("'", pkg_to_update, "'", collapse = ","), "))")
    message("\nRun the following command to update all outdated packages:\n", install_cmd)
  } else {
    message("All packages are up to date.")
  }
  
  return(pkg_table)
}


createDataDictionary <- function(data = NULL, char_limit = 32767) { # 32767 is xlsx cell element limit (likely only being reached by ID variables)
  
  suppressWarnings({
    
    # Define list choices for yes/no prompts with padding for better presentation
    list_choices <- c(paste0("Yes",  paste0(rep(" ", 100), collapse = " ")) ,
                      paste0("No ",  paste0(rep(" ", 100), collapse = " ")))
    
    if (is.null(data)) {
      find_df <- TRUE
    } else {
      find_df <- FALSE
      df <- data
    }
    
    getExtension <- function(file) { 
      ex <- strsplit(basename(file), split = "\\.")[[1]]
      return(ex[-1])
    } 
    
    # Check if required packages are installed
    required_packages <- c("haven", "openxlsx", "svDialogs", "moments")
    inst <- installed.packages()
    if (!all(required_packages %in% inst)) {
      install.packages(setdiff(required_packages, rownames(inst)))
    }
    
    # Load required packages
    for(pack in required_packages) {
      library(pack, character.only = TRUE)
    }
    
    # Read .dta file if find_df is TRUE
    if(find_df) {
      mod <- svDialogs::dlg_open()
      if (identical(mod$res, character(0))) return("process terminated by user")
      path <- mod$res
      print(path)
      workDirectory <- dirname(path)
      if(getExtension(path) == "dta"){
        df <- read_dta(path)
      } else if (getExtension(path) == "sav") {
        df <- read_sav(path)
      } else if(getExtension(path) == "spss") {
        df <- read_spss(path)
      } else if(getExtension(path) == "sas7bdat") {
        df <- read_sas(path)
      } else {
        return("This function is not compatible with your selected file")
      } 
      
    } else {
      mod <- svDialogs::dlg_dir(title = "Choose directory to store your dictionary")
      if (identical(mod$res, character(0))) return("process terminated by user")
      workDirectory <- mod$res
    }
    
    # Create data types
    myList <- sapply(as_factor(df), class)
    Dtypes <- sapply(myList, function(x) paste(x,collapse = " / "))
    
    prop_missing <- function(x) {
      100 * (sum(is.na(x))  / length(x)) |> round(4) 
    }
    
    n_missing <- function(x) {
      paste(sum(is.na(x)), '/', length(x))
    }
    
    Means <- function(x) {
      tryCatch( {
        conf <- t.test(x)$conf.int
        return (paste0(mean(x, na.rm = TRUE) |> round(2),", (",
                       round(conf[1], 2),' , ', round(conf[2], 2),")") ) 
      },
      error = function(e) {
        if(all(is.na(x)) ) {
          return("No data")
        }else {
          return("Data is constant") 
        }
      })
    }
    
    IQR <- function(x) {
      tryCatch({ # The try catch helps with sapply(df, function) in the data.frame() part
        y <- round(quantile(x, na.rm = TRUE), 2)
        med_iqr <- paste0(y[3]," [", y[2]," , ", y[4], "]")
        med_iqr <- ifelse(med_iqr == "NA [NA , NA]",
                          "No data", med_iqr)
        return(med_iqr)
      },
      error = function(e) {
        return("")
      })
    }
    
    SkewKurt <- function(x) {
      tryCatch({
        # Skewness and Kurtosis calculations using moments package
        skew <- round(moments::skewness(x, na.rm = TRUE), 2)
        kurt <- round(moments::kurtosis(x, na.rm = TRUE), 2)
        
        # Skewness description with neighborhood around 0
        skew_desc <- ifelse(skew > 0.5, "Right skewed", 
                            ifelse(skew < -0.5, "Left skewed", "Symmetric"))
        
        # Kurtosis description with neighborhood around 3
        kurt_desc <- ifelse(kurt > 3.5, "Leptokurtic (heavy tails)", 
                            ifelse(kurt < 2.5, "Platykurtic (light tails)", "Mesokurtic (normal tails)"))
        
        # Combine and return
        return(paste0("Skewness: ", skew_desc, " (", skew, "), Kurtosis: ", kurt_desc, " (", kurt, ")"))
      },
      error = function(e) {
        if(all(is.na(x))) {
          return("No data")
        } else {
          return("Data is constant")
        }
      })
    }
    
    
    Table <- function(x) {
      freq_table <- table(x)
      if(toString(freq_table) == "") return("No Data")
      prop_table <- paste0(freq_table, " (", round(100 * prop.table(freq_table) ,1), "%)")
      table_str <- paste(names(freq_table), prop_table, sep = ": ", collapse = ", ")
      table_str <- ifelse(nchar(table_str > char_limit),
                          substr(table_str, 1, char_limit),
                          table_str)
      return(table_str)
    }
    
    Label <- function(x) {
      Labs <- attr(x, 'labels')
      if (is.null(Labs)) return("")
      label_str <- paste(names(Labs), Labs, sep = " = ", collapse = ", ")
      label_str <- ifelse(nchar(label_str > char_limit),
                          substr(label_str, 1, char_limit),
                          label_str)
      return(label_str)
    }
    
    tab <- data.frame (
      Variables = names(df),
      Labels = sapply(sapply(df, attr, 'label'), toString),
      Data_type = Dtypes ,
      Missing_perc = (miss <- sapply(df, prop_missing)),
      Missing_n = sapply(df, n_missing),
      Min_score = ifelse(miss < 100, apply(df,2,min, na.rm = TRUE), "No data" ),
      Max_score = ifelse(miss < 100, apply(df,2,max, na.rm = TRUE), "No data" ),
      med_iqr = ifelse(miss < 100 & Dtypes == "numeric",sapply(df, IQR), "" ),
      means = ifelse(miss < 100 & Dtypes == "numeric",sapply(df, Means), "" ),
      skew_curt = ifelse(miss < 100 & Dtypes == "numeric",sapply(df, SkewKurt), "" ),
      `Value:Counts` = ifelse(Dtypes %in% c("factor","character","logical"), 
                              sapply(as_factor(df), Table), ""),
      `Value:Labels` = ifelse(Dtypes %in% c("factor","character","logical"), 
                              sapply(df, Label), "")
    )
    
    names(tab)[c(3:12)] <- c("Data type","% missing", "Number of missing","Min score","Max score", "Median [IQR]",
                             "Mean (95% LCI, 95% UCI)","Skewness & kurtosis" ,"Value : n (%)", "Value = Labels")
    
    # Write to Excel file
    repeat { # Check if file already exists and will ask if you want to overwrite or not. 
      mod <- svDialogs::dlg_input("Name your data dictionary, e.g. dataDict")
      if (identical(mod$res, character(0))) return("process terminated by user")
      dictName <- mod$res
      
      if(find_df) {
        mod <- svDialogs::dlg_list(choices = list_choices, title = "Create dictionary in the same directory as your data?")
        if (identical(mod$res, character(0))) return("process terminated by user")
        user_choice <- gsub(" ", "", mod$res, fixed = TRUE)
        if (user_choice == "No") {
          mod <- svDialogs::dlg_dir()
          if (identical(mod$res, character(0))) return("process terminated by user")
          output_file <- paste0(mod$res, "/", dictName, ".xlsx")
        } else {
          output_file <- paste0(workDirectory, "/", dictName, ".xlsx")
        }
      } else {
        output_file <- paste0(workDirectory, "/", dictName, ".xlsx")
      }
      
      # Check if file exists
      if (file.exists(output_file)) {
        mod <- svDialogs::dlg_list(choices = list_choices, title = "File already exists. Do you want to overwrite it?")
        if (identical(mod$res, character(0))) return("process terminated by user")
        user_choice <- gsub(" ", "", mod$res, fixed = TRUE)
        if (user_choice == "No") {
          # Ask for a new name
          next
        } else {
          break
        }
      } else {
        break
      }
    }
    
    openxlsx::write.xlsx(tab, output_file, asTable = TRUE)
    svDialogs::dlg_message("Data dictionary successfully created!")
    
  })
}

# Example usage:

# Run checkNecessaryPackages() first if you think you may have outdated packages
# This function will tell you your current versions and available versions
# A message with code install.packages(c()) will be pasted to console. Run this command exactly as shown to update necessary packages

# createDataDictionary() # Prompts you to select a data file on your PC
# createDataDictionary(data = df) # Creates a data dictionary using df already in memory
