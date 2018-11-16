################################################################################
# 
#                       COMBINED AND EXPLORE ALL DATASETS
#
# 
################################################################################
################################################################################
# CLEAR ENVIRONMENT, SET WORKING DIRECTORY, IDENTIFY LIBRARY NEEDED.
################################################################################
rm(list = ls())
setwd("C:/Users/USER/Desktop/R/BRR/2018 V2/2018")
################################################################################
library(XLConnect)
library(openxlsx)
library(reshape)
library(installr)
################################################################################
files = list.files()
unopenable = files[grepl('xlsx$', unlist(as.list(files[1:length(files)])))!= TRUE]
unopenable = c(unopenable, files[grepl('BIBS.xlsx$', unlist(as.list(files[1:length(files)])))] )
files = files[grepl('xlsx$', unlist(as.list(files[1:length(files)])))]
files = files[-which(files == files[grepl('BIBS.xlsx$', unlist(as.list(files[1:length(files)])))])]
################################################################################
##################        Extract data from files       ########################
################################################################################
l = length(files)
races = unlist(strsplit(files,".xlsx"))
important_list = c("Last Name","First Name","race","City","State","Age","Country","Email","Sex")

info = as.data.frame(matrix(NA,l,length(important_list)))
colnames(info) = important_list

for (i2 in  1:l){
  data.reading = readxl::read_excel(files[i2],sheet = 1)
  for( i in important_list){
    letter <- idx2col(which(names(data.reading) == i))
    if (is.empty(letter)==FALSE){
      info[i2,i] = letter
    }
  }
}
rownames(info)  = races
info

###########################################################################################
out_dir = "C:/Users/USER/Desktop/R/BRR"
WBXls = paste(out_dir,"All summary.xlsx",sep='/')
unlink(WBXls, recursive = FALSE, force = FALSE)
WB <- loadWorkbook(WBXls,create=TRUE)


createSheet(WB,races[1])
saveWorkbook(WB)
###########################################################################################
################################ Explore Function #########################################
###########################################################################################
summary(as.factor(data.reading$Sex))

createName(WB, name='Sex',formula= Sex.formula)

UNIQUE = unique(data.reading$Sex)
writeNamedRegion(WB, c("Sex",UNIQUE), name = 'Sex',header = FALSE)
saveWorkbook(WB)

Sex.formula = paste0('\'',races[1],'\'!$A$1')

if (is.na(info[1,i])==FALSE){
  formulaXls <- paste0('COUNTIF(','\'',races[1],'\'!',info[1,i],':',info[1,i],",\"",UNIQUE[1],"\")")
}

setCellFormula(WB, sheet=races[1],2,2,formulaXls)
saveWorkbook(WB)

###########################################################################################
################################ Way 2 #########################################
###########################################################################################


out_dir = "C:/Users/USER/Desktop/R/BRR/2018 V2/2018"
WBXls = paste(out_dir,files[1],sep='/')
# unlink(WBXls, recursive = FALSE, force = FALSE)

###########################################################################################
wb = openxlsx::loadWorkbook(xlsxFile = files[1])
sheet = names(wb)  #list worksheets

addWorksheet(wb,"Summary" )
writeData(wb,"Summary" , x = c("Sex",UNIQUE))
v <- paste0('COUNTIF(','\'',sheet[1],'\'!',info[1,i],':',info[1,i],',"=',UNIQUE[1],'")')
# v <- paste0('COUNTIF(','\'Sheet1\'!',info[1,i],':',info[1,i],',"=',UNIQUE[1],'")')

writeFormula(wb,"Summary", x = v, startCol = 2, startRow = 2)
saveWorkbook(wb, WBXls, overwrite = TRUE)

#################################################################################################
# USE R WORKBOOK CAN BE PROBLEMATIC BECAUSE THE ORIGINAL FORMAT OF ALL SHEET WILL BE MESSED UP.
#################################################################################################







