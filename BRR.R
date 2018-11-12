################################################################################
# CLEAR ENVIRONMENT, SET WORKING DIRECTORY, IDENTIFY LIBRARY NEEDED.
################################################################################
# GET AN IDEA OF WHAT I CAN DO WITH THE GIVEN DATA.
################################################################################
rm(list = ls())
################################################################################
setwd("C:/Users/USER/Desktop/R/BRR/BRR")
################################################################################
library(XLConnect)
library(reshape)
##########################################################################################
###################### Go To Working Directory and Read Data #############################
##########################################################################################
datasets = readxl::read_excel("C:/Users/USER/Desktop/R/BRR/2018 Boston 5.25 Miler.xlsx",sheet = 1)
sample <- loadWorkbook("C:/Users/USER/Desktop/R/BRR/2017 El Pelon 5K Registration.xlsx", create = FALSE)
sample = readxl::read_excel("C:/Users/USER/Desktop/R/BRR/2017 El Pelon 5K Registration.xlsx",sheet = 4)
###########################################################################################
############################### Create a File #############################################
###########################################################################################
# output directory
out_dir = "C:/Users/USER/Desktop/R/BRR"
fileXls = paste(out_dir,"Testing.xlsx",sep='/')
unlink(fileXls, recursive = FALSE, force = FALSE)
exc <- loadWorkbook(fileXls, create = TRUE)

createSheet(exc,'2018 Boston 5.25 Miler')
saveWorkbook(exc)

###########################################################################################
out_dir = "C:/Users/USER/Desktop/R/BRR"
WBXls = paste(out_dir,"2018 Boston 5.25 Miler.xlsx",sep='/')
unlink(WBXls, recursive = FALSE, force = FALSE)
WB <- loadWorkbook(WBXls)

createSheet(WB,'Summary')
saveWorkbook(WB)

###########################################################################################
################################ Explore Function #########################################
###########################################################################################

createSheet(exc,'Airquality')
airquality$isCurrent<-NA
createName(exc, name='Airquality',formula='Airquality!$A$1')
writeNamedRegion(exc, airquality, name = 'Airquality', header = TRUE)
saveWorkbook(exc)


colIndex <- which(names(airquality) == 'isCurrent')
letterDay <- idx2col(which(names(airquality) == 'Day'))
letterMonth <- idx2col(which(names(airquality) == 'Month'))
formulaXls <- paste('IF(AND(',
                    letterMonth,
                    2:(nrow(airquality)+1),
                    '=Input!C3,',
                    letterDay,
                    2:(nrow(airquality)+1),
                    '=Input!C2)',
                    ',1,0)',sep='')
setCellFormula(exc, sheet='Airquality',2:(nrow(airquality)+1),colIndex,formulaXls)
saveWorkbook(exc)

createSheet(exc,'Input')
saveWorkbook(exc)
input <- data.frame('inputType'=c('Day','Month'),'inputValue'=c(2,5))
writeWorksheet(exc, input, sheet = "Input", startRow = 1, startCol = 2)
saveWorkbook(exc)
###########################################################################################
############################### Summarize Sex #############################################
###########################################################################################
writeWorksheet(WB, "2018 Boston 5.25 Miler", sheet = 'Summary', startRow = 1, startCol = 1, header=FALSE )

Sex = datasets$Sex
Sex = Sex[(is.na(Sex)==FALSE)] 
Male = sum(Sex=="M")
Female = sum(Sex=="F")
Sex.Population = length(Sex)

# letterSex <- idx2col(which(names(datasets) == 'Sex'))
# Count.Sex.M = paste0('=COUNTIF(\'Race Roster\'!',letterSex,':',letterSex,',"=M")')
# Count.Sex.F = paste0('=COUNTIF(\'Race Roster\'!',letterSex,':',letterSex,',"=F")')
# Sex.summary <- data.frame('Gender'=c('Male','Female'),'Number'=c(Count.Sex.M,Count.Sex.F),'Portion'=c(Male/Sex.Population , Female/Sex.Population))
Sex.summary <- data.frame('Gender'=c('Male','Female','Total'),'Number'=c(Male,Female,Sex.Population),'Portion'=c(Male/Sex.Population , Female/Sex.Population,(Male+Female)/Sex.Population ))

writeWorksheet(WB, Sex.summary, sheet = 'Summary', startRow = 3, startCol = 1)
saveWorkbook(WB)

###########################################################################################
############################### Summarize Age #############################################
###########################################################################################

Age = datasets$Age[(is.na(datasets$Age)==FALSE)] 
Age.mod = Age%/%10
Age.summary <- data.frame('Age'=c(as.character(sample$X__1[8:15]),"NA","Total"),
                          'Number'=rep(0,10))
for (i in 0:6){
  Age.summary$Number[(i+1)] =  sum(Age.mod==i)
}
Age.summary$Number[8] = sum(Age.mod>=7)
Age.summary$Number[9] = sum(is.na(datasets$Age))
Age.summary$Portion = Age.summary$Number/sum(Age.summary$Number)
Age.summary[10,] = c("Total",sum(Age.summary$Number),1)

writeWorksheet(WB, Age.summary, sheet = 'Summary', startRow = 8, startCol = 1)
saveWorkbook(WB)

###########################################################################################
############################### Summarize Age #############################################
###########################################################################################

City.summary <- data.frame(summary(as.factor(datasets$City)))
colnames(City.summary) = c("City")
City.summary$Number = rownames(City.summary)
City.summary[,c(1,2)]=City.summary[,c(2,1)]
City.summary$Portion = City.summary$Number/sum(City.summary$Number)
City.summary = rbind(City.summary,c("Total",sum(City.summary$Number),1))

writeWorksheet(WB, City.summary, sheet = 'Summary', startRow = 22, startCol = 1)
saveWorkbook(WB)

###########################################################################################
############################### Summarize Age #############################################
###########################################################################################


