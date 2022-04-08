#############################################################
#                       LOAD LIBRARY
#############################################################
rm(list = ls())
library(XML)
library(RCurl)
library(rlist)
library(HelpersMG)
library(installr)
library(openxlsx)
library(reshape)
library(installr)
##########################################################################################################################
#                       SET DIRECTORY
##########################################################################################################################

setwd("C:/Users/USER/Desktop/BRR")
files = list.files()
files = files[grepl('^~', unlist(as.list(files[1:length(files)])))!= TRUE]

dataset = read.xlsx(files,sheet = 1)

wget(url = "https://www.asics.com/us/en-us/contact" ,destfile ="website")

webpage <- readLines("website")
htmlpage <- htmlParse(webpage, asText = TRUE)
email = xpathSApply(htmlpage, "//div/p/a [@href='mailto:aac-pr@asics.com']", xmlValue)


