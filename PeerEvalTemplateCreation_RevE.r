
library("tidyverse")
library("openxlsx")    # library("readxl") # this is the tidyverse installed package
library("scales")
library("lubridate")
library("rstudioapi")
library("jsonlite")

# rm(list=ls())         # remove all global environment variables


filenamebase <- "_151A_F24_MidQPeerEvalTemplate.xlsx"

BaseDir <- setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
DataDir <- paste0(BaseDir,"/TeamTemplates/")
#setwd(DataDir)

teamdata <- read.csv("GroupsWithNames.csv", header = TRUE, stringsAsFactors = FALSE)


TwoPerson <- teamdata %>% mutate(TeamNum = gsub(" ", "",str_extract(teamdata$ProjectName,"^([^-])+"))) %>% filter(NumMembers==2) 
ThreePerson <- teamdata %>% mutate(TeamNum = gsub(" ", "",str_extract(teamdata$ProjectName,"^([^-])+"))) %>% filter(NumMembers==3) 
FourPerson <- teamdata %>% mutate(TeamNum = gsub(" ", "",str_extract(teamdata$ProjectName,"^([^-])+"))) %>% filter(NumMembers==4) 
FivePerson <- teamdata %>% mutate(TeamNum = gsub(" ", "",str_extract(teamdata$ProjectName,"^([^-])+"))) %>% filter(NumMembers==5)
SixPerson <- teamdata %>% mutate(TeamNum = gsub(" ", "",str_extract(teamdata$ProjectName,"^([^-])+"))) %>% filter(NumMembers==6)
SevenPerson <- teamdata %>% mutate(TeamNum = gsub(" ", "",str_extract(teamdata$ProjectName,"^([^-])+"))) %>% filter(NumMembers==7) 

for (i in 1:nrow(TwoPerson)/2) {
  wb <- loadWorkbook("PeerEvaluationTemplate_2Person.xlsx")
  for (j in 1:2) 
    writeData(wb,"PeerRating",TwoPerson$Name[i*2-2+j], startCol = j, startRow = j)
  protectWorksheet(wb, "PeerRating", protect = TRUE, password = "456", lockSelectingLockedCells = TRUE, 
                   lockSelectingUnlockedCells = FALSE, lockObjects = TRUE, lockScenarios = TRUE)
  saveWorkbook(wb,paste(DataDir,"Team",TwoPerson$TeamNum[i*2],filenamebase,sep = ""),overwrite = TRUE)
}

for (i in 1:nrow(ThreePerson)/3) {
  wb <- loadWorkbook("PeerEvaluationTemplate_3Person.xlsx")
  for (j in 1:3) 
    writeData(wb,"PeerRating",ThreePerson$Name[i*3-3+j], startCol = j, startRow = j)
  protectWorksheet(wb, "PeerRating", protect = TRUE, password = "456", lockSelectingLockedCells = TRUE, 
                   lockSelectingUnlockedCells = FALSE, lockObjects = TRUE, lockScenarios = TRUE)
  saveWorkbook(wb,paste(DataDir,"Team",ThreePerson$TeamNum[i*3],filenamebase,sep = ""),overwrite = TRUE)
}

for (i in 1:nrow(FourPerson)/4) {
  wb <- loadWorkbook("PeerEvaluationTemplate_4Person.xlsx")
  for (j in 1:4) 
    writeData(wb,"PeerRating",FourPerson$Name[i*4-4+j], startCol = j, startRow = j)
  protectWorksheet(wb, "PeerRating", protect = TRUE, password = "456", lockSelectingLockedCells = TRUE, 
                   lockSelectingUnlockedCells = FALSE, lockObjects = TRUE, lockScenarios = TRUE)
  saveWorkbook(wb,paste(DataDir,"Team",FourPerson$TeamNum[i*4],filenamebase,sep = ""),overwrite = TRUE)
}

for (i in 1:nrow(FivePerson)/5) {
  wb <- loadWorkbook("PeerEvaluationTemplate_5Person.xlsx")
  for (j in 1:5) 
    writeData(wb,"PeerRating",FivePerson$Name[i*5-5+j], startCol = j, startRow = j)
  protectWorksheet(wb, "PeerRating", protect = TRUE, password = "456", lockSelectingLockedCells = TRUE, 
                   lockSelectingUnlockedCells = FALSE, lockObjects = TRUE, lockScenarios = TRUE)
  saveWorkbook(wb,paste(DataDir,"Team",FivePerson$TeamNum[i*5],filenamebase,sep = ""),overwrite = TRUE)
}

for (i in 1:nrow(SixPerson)/6) {
  wb <- loadWorkbook("PeerEvaluationTemplate_6Person.xlsx")
  for (j in 1:6) 
    writeData(wb,"PeerRating",SixPerson$Name[i*6-6+j], startCol = j, startRow = j)
  protectWorksheet(wb, "PeerRating", protect = TRUE, password = "456", lockSelectingLockedCells = TRUE, 
                   lockSelectingUnlockedCells = FALSE, lockObjects = TRUE, lockScenarios = TRUE)
  saveWorkbook(wb,paste(DataDir,"Team",SixPerson$TeamNum[i*6],filenamebase,sep = ""),overwrite = TRUE)
}

for (i in 1:nrow(SevenPerson)/7) {
  wb <- loadWorkbook("PeerEvaluationTemplate_7Person.xlsx")
  for (j in 1:7) 
    writeData(wb,"PeerRating",SevenPerson$Name[i*7-7+j], startCol = j, startRow = j)
  protectWorksheet(wb, "PeerRating", protect = TRUE, password = "456", lockSelectingLockedCells = TRUE, 
                   lockSelectingUnlockedCells = FALSE, lockObjects = TRUE, lockScenarios = TRUE)
  saveWorkbook(wb,paste(DataDir,"Team",SevenPerson$TeamNum[i*7],filenamebase,sep = ""),overwrite = TRUE)
}























# Used this to fix the silly error that I was getting in my templates
wb <- loadWorkbook("PeerEvaluationTemplate_2Person.xlsx")
saveWorkbook(wb, "test2.xlsx", overwrite = TRUE)
wb <- loadWorkbook("PeerEvaluationTemplate_3Person.xlsx")
saveWorkbook(wb, "test3.xlsx", overwrite = TRUE)
wb <- loadWorkbook("PeerEvaluationTemplate_4Person.xlsx")
saveWorkbook(wb, "test4.xlsx", overwrite = TRUE)




# Example from https://www.rdocumentation.org/packages/openxlsx/versions/4.2.5/topics/writeData

wb2 <- createWorkbook()

## Add worksheets
addWorksheet(wb2, "Cars")
addWorksheet(wb2, "Formula")


x <- mtcars[1:6, ]
writeData(wb2, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)


#####################################################################################
## Bordering

writeData(wb2, "Cars", x,
          rowNames = TRUE, startCol = "O", startRow = 3,
          borders = "surrounding", borderColour = "black"
) ## black border

writeData(wb2, "Cars", x,
          rowNames = TRUE,
          startCol = 2, startRow = 12, borders = "columns"
)

writeData(wb2, "Cars", x,
          rowNames = TRUE,
          startCol = "O", startRow = 12, borders = "rows"
)


#####################################################################################
## Header Styles

hs1 <- createStyle(
  fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "italic",
  border = "Bottom"
)

writeData(wb2, "Cars", x,
          colNames = TRUE, rowNames = TRUE, startCol = "B",
          startRow = 23, borders = "rows", headerStyle = hs1, borderStyle = "dashed"
)


hs2 <- createStyle(
  fontColour = "#ffffff", fgFill = "#4F80BD",
  halign = "center", valign = "center", textDecoration = "bold",
  border = "TopBottomLeftRight"
)

writeData(wb2, "Cars", x,
          colNames = TRUE, rowNames = TRUE,
          startCol = "O", startRow = 23, borders = "columns", headerStyle = hs2
)


## Save workbook
## Open in excel without saving file: openXL(wb)
# }
# NOT RUN {
saveWorkbook(wb2, "writeDataExample.xlsx", overwrite = TRUE)
# }

