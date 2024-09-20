createDataDictionary <- function(data = NULL) {
  
  suppressWarnings({
    
    # Define list choices for yes/no prompts with padding for better presentation
    list_choices <- c(paste0("Yes",  paste0(rep(" ", 60), collapse =" ")) ,
                      paste0("No ",  paste0(rep(" ", 60), collapse =" ")))
    
    if (is.null(data)) {
      find_df <- TRUE
    } else {
      find_df <- FALSE
      df <- data
    }
    
    getExtension <- function(file){ 
      ex <- strsplit(basename(file), split="\\.")[[1]]
      return(ex[-1])
    } 
    
    # Check if required packages are installed
    required_packages <- c("haven", "openxlsx", "svDialogs")
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
    
    # Proportion missing
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
        return("Data is constant")
      })
    }
    
    Table <- function(x) {
      freq_table <- table(x)
      table_str <- paste(names(freq_table), freq_table, sep = ": ", collapse = ", ")
      return(table_str)
    }
    
    Label <- function(x) {
      Labs <- attr(x, 'labels')
      label_str <- paste(names(Labs), Labs, sep = " = ", collapse = ", ")
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
      means = ifelse(miss < 100 & Dtypes=="numeric",sapply(df, Means), "" ),
      `Value:Counts` = ifelse(Dtypes %in% c("factor","character","logical"), 
                              sapply(as_factor(df), Table), ""),
      `Value:Labels` = ifelse(Dtypes %in% c("factor","character","logical"), 
                              sapply(df, Label), "")
    )
    
    names(tab)[c(3:10)] <- c("Data type","% missing", "Number of missing","Min score","Max score",
                             "Mean (95% LCI, 95% UCI)","Value : Counts", "Value = Labels")
    
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
# createDataDictionary() # Prompts you to select a data file on your PC
# createDataDictionary(data = df) # Creates a data dictionary using df already in memory
