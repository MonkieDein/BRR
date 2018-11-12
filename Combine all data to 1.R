################################################################################
# CLEAR ENVIRONMENT, SET WORKING DIRECTORY, IDENTIFY LIBRARY NEEDED.
################################################################################
rm(list = ls())
setwd("C:/Users/USER/Desktop/R/BRR/2018")
################################################################################
library(XLConnect)
library(reshape)
################################################################################
files = list.files()
unopenable = files[grepl('xlsx$', unlist(as.list(files[1:length(files)])))!= TRUE]
files = files[grepl('xlsx$', unlist(as.list(files[1:length(files)])))]
################################################################################
##################        Extract data from files       ########################
################################################################################
l = length(files)
races = unlist(strsplit(files,".xlsx"))

combined_data = readxl::read_excel(files[1],sheet = 1)
combined_data$race = races[1]
combined_data$'Date (mm/dd/yyyy)' = as.character(unlist(combined_data$'Date (mm/dd/yyyy)'))
combined_data$'Date of Birth' = as.character(unlist(combined_data$'Date of Birth'))
for (i in 2:l){
  datasets = readxl::read_excel(files[i],sheet = 1)  
  if (nrow(datasets)==0){
    next
  }
  datasets$race = races[i]
  if (is.null( datasets$'Date (mm/dd/yyyy)')==FALSE){
    datasets$'Date (mm/dd/yyyy)' = as.character(unlist(datasets$'Date (mm/dd/yyyy)'))
  }
  if (is.null( datasets$'Date of Birth')==FALSE){
    datasets$'Date of Birth' = as.character(unlist(datasets$'Date of Birth'))
  }
  combined_data = merge(combined_data, datasets, all.x = TRUE , all.y = TRUE)
}

comb.fix = combined_data[-which(is.na(combined_data$`First Name`) & is.na(combined_data$`Last Name`) ),]
bad = combined_data[which(is.na(combined_data$`First Name`) & is.na(combined_data$`Last Name`) ),]

write.csv(comb.fix , "combined_data.csv")

# write.csv(bad,"bad_observation.csv") 
# All the data only have bib number without runners data So we ignore those data.
# The only data that might be useful is ther 2018 hot chocolate run because there is a data has donation but has no name.








