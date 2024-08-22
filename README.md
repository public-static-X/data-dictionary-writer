# data-dictionary-writer
creates a data dictionary with basic descriptives for sav, sas and stata datasets

To use, control + A the script and paste it into Rstudio (or download the script and open in RStudio). Then control + Shift + Enter
run the command createDataDictionary() and the follow the prompts.

If you are just using the R console without RStudio, then just paste the code into the R console and run createDataDictionary()

Alternatively (this is only for R users), if you already have a dataset in Ram that you have been using...
then run createDataDictionary(data = df) with df being your data frame.
