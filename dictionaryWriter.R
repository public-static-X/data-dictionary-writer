createDataDictionary <- function(data = NULL) {
  
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
    path <- svDialogs::dlg_open()$res
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
    workDirectory <- dlg_dir(title = "Choose directory to store your dictionary")$res # Use this to select a directory to 
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
  dictName <- svDialogs::dlg_input("Name your data dictionary, e.g. dataDict")$res
  
  if(find_df) {
    sameDir <- svDialogs::dlg_message("Create dictionary in the same directory as your data ? Click No for a different directory",
                                      type = "yesno")$res
  } else {
    sameDir <- "yes"
  }
  
  is_created <- FALSE
  output_file <- ifelse(sameDir == "yes", paste0(workDirectory,"/",dictName,".xlsx"),
                        paste0(svDialogs::dlg_dir()$res,"/",dictName,".xlsx") )
  
  openxlsx::write.xlsx(tab, output_file, asTable = TRUE)
  is_created <- TRUE
  if(is_created) svDialogs::dlg_message("Data dictionary successfully created!")
}

# Example usage:


# createDataDictionary() # This will prompt you to select a data file on your PC
# createDataDictionary(data = df) # This will create a data dictionary using df which is already in memory
