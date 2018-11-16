################################################################################
# 
#                       CLEAN THE COMBINED DATASETS
#
# 
################################################################################
################################################################################
# CLEAR ENVIRONMENT, SET WORKING DIRECTORY, IDENTIFY LIBRARY NEEDED.
################################################################################
rm(list = ls())
setwd("C:/Users/USER/Desktop/R/BRR/2018 V2/2018/Summary")
################################################################################
library(XLConnect)
library(reshape)
################################################################################

dataset = read.csv("combined_data.csv")

################################################################################
################################################################################
sum((is.na(dataset$"Donations")==FALSE) & (is.na(dataset$"Donation")==FALSE))
sum(is.na(dataset$"Donations") & is.na(dataset$Donation))
Donation = dataset$Donation  
Donation[is.na(Donation)] =   dataset$Donations[is.na(Donation)] 

################################################################################
colSums(is.na(dataset))[1:5]
dataset[dataset=="?"] = NA

Clean.data = dataset[,c("Last.Name","First.Name","race","City","State","Age","Country","Email","Sex")]

# Clean.data$Donation = Donation

# dataset = dataset[,-c(1,2,3,4,5,6,7,9,195,122,120,89,118)]
################################################################################

Sex = as.factor(ifelse(dataset$Gender=="Female","F","M"))
Sex[is.na(Sex)] = dataset$Sex[is.na(Sex)]
Clean.data$Sex = Sex

################################################################################
# names = colnames(dataset)
# 
# 
# names[grepl('^G', unlist(as.list(names[1:length(names)])))]
# names[grepl('^S', unlist(as.list(names[1:length(names)])))]
# 
# DEL = dataset[,-c(89,118)]
# 
# empty.col = order(colSums(is.na(dataset)))
# 
# empty.col
# 
# 
# colnames(dataset)[empty.col]
# 
# order(-colSums(is.na(dataset)))
# 
# 
# colSums(is.na(DEL))[1:5]
# 
# 
# 
# names[grepl('^D', unlist(as.list(names[1:length(names)])))]
# 
# dataset[,c("Date","Date..mm.dd.yyyy.")]
# 
# sum(is.na(dataset$"Date..mm.dd.yyyy.") & is.na(dataset$Date))
# sum((is.na(dataset$"Date..mm.dd.yyyy.")==FALSE) & (is.na(dataset$Date)==FALSE))
# 


################################################################################
Clean.data$Country[Clean.data$Country =="United States"] = "US"

Clean.data$State[Clean.data$State =="Massachusetts"] = "MA"
Clean.data$State[Clean.data$State =="Greater London"] = "England"



Clean.data
unique(Clean.data$Country)

################################################################################
write.csv(Clean.data , "simpledata.csv",row.names = FALSE)
