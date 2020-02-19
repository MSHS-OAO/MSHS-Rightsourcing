###----------------------MSH Rightsourcing

#Rightsource file needs to be "eport table" format. Can include all sites
rightsourcing <- function(dateRSfilename,JCfilename){
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSH")
  #read current job code list
  jc <- read.csv("Rightsource Job Code.csv",stringsAsFactors = F,header=F)
  colnames(jc) <- c("JobTitle","JobCode")
  #Read raw excel file from rightsourcing
  right <- read.csv(file.choose(), fileEncoding = "UTF-16LE",sep='\t',header=T,stringsAsFactors = F)
  colnames(right) <- c("Quarter",	"Months",	"Service Line",	"Worker",	"Manager",	"Supplier",
                       "JobTitle",	"Job Category",	"Job Class",	"Client Bill Rate",	"Dept",
                       "Dept Lvl 1","New Fill?","Dept Lvl 2",	"Location",	"Hours",	"Spend",	"Earnings E/D",
                       "All Other Spend",	"OT Spend")
  #replace N/A job titles with unkown job title
  for(i in 1:nrow(right)){
    if(right$JobTitle[i] == "#N/A"){
      right$JobTitle[i] <- "Unknown"
    }
  }
  library(anytime)
  right$`Earnings E/D` <- anytime(right$`Earnings E/D`)
  #filter on MSH and dates after end of last upload date
  right1 <<- right[right$Location == "MSH - Mount Sinai Hospital" & right$`Earnings E/D` >= anytime(date),]
  #replace workers name with just their last name
  right1$Worker <- gsub(",.*$","",right1$Worker)
  #add "Rightsource" to the the job title
  right1$JobTitle <- paste("Rightsource",right1$JobTitle,sep=" ")
  #subset all rows with new job codes
  newright <- subset(right1,!(right1$JobTitle %in% jc$JobTitle))
  #append new jobs only if there are new jobs
  if(nrow(newright) != 0){
    #create vector of unique new jobs
    newjobs <- unique(newright$JobTitle)
    #append the new jobs to the jobcode list
    #i from job list length plus 1 to length of new jobs
    a <- nrow(jc)
    for(i in a+1:length(newjobs)){
      #append with job title and job code for all new jobs
      #if job is over the 99th job, adjust the coding scheme
      jc[i,] <- list(newjobs[i-a],if(i>99){
        paste("B00",i,sep="")
      }else{
        paste("B000",i,sep="")
      })
    }
  }
  #append job code to prepare fo jobcode dictionary
  newrightjc <<- merge(newright,jc,by="JobTitle")
  #only create job code dictionary if there are new job codes
  if(nrow(newrightjc) != 0){
    #Creat rightsourcing jobcode dictionary from newrightjc
    jcdict <- data.frame(partner="729805", hospital="NY0014",dep=substr(newrightjc$Dept,start=1,stop=8),
                         jc=newrightjc$JobCode,title=newrightjc$JobTitle)
    #Only take unique rows from jcdict
    jcdict2 <- jcdict[!duplicated(jcdict),]
  }
  #merge the rightsourcing jobcode dictionary and the rightsourcing file
  #essentially assigns jobcode based on job title
  right2 <<- merge(right1,jc,by="JobTitle")
  #round the hours and spend columns
  right2$Hours <- round(right2$Hours,2)
  right2$Spend <- round(right2$Spend,2)
  #Create Premier formatted export
  export <-  data.frame(partner="729805",hospital="NY0014",home=substr(right2$Dept,start=1,stop=8),
                        hosp="NY0014",work=substr(right2$Dept,start=1,stop=8),start=right2$`Earnings E/D`-6,
                        end=right2$`Earnings E/D`,EmpCode=paste0(substr(right2$Worker,start=1,stop=12),gsub("\\..*","",right2$Hours)),
                        name=right2$Worker,budget="0",JobCode=right2$JobCode,paycode="AG1",
                        hours=right2$Hours,spend=right2$Spend)
  export$start <- paste(substr(export$start,start=6,stop=7),"/",substr(export$start,start=9,stop=10),
                        "/",substr(export$start,start=1,stop=4),sep="")
  export$end <- paste(substr(export$end,start=6,stop=7),"/",substr(export$end,start=9,stop=10),
                      "/",substr(export$end,start=1,stop=4),sep="")
  export <<- export
  jc <<- jc
  newrightjc <<- newrightjc
}

#function to save the rightsource and jobcode exports in home directory and jc list in rightsourcing directory
save <- function(RSfilename,JCfilename){
  #----------------------------------------------------------------------------------------------------------------
  #Save new export
  setwd("~/")
  write.table(export,file=RSfilename,sep=",",col.names=F,row.names=F)
  #only save jc dictionary if there are new jobs
  if(exists("jcdict2")){
    #save job code dictionary
    write.table(jcdict2,file=JCfilename,sep=",",col.names=F,row.names=F)
  }
  #Overwrit Jobcode table
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSH")
  write.table(jc,file="Rightsource Job Code.csv",sep=",",col.names=F,row.names=F)
}

#date should be the day after the last upload date for rightsourcing
#RSfilename is the filename for rightsourcing upload file
#JCfilename is the filename for the jcbcode dictionary upload file
rightsourcing(date = "12/01/2019")

#check the "export", "jc" and "jcdict2". If all look correct then run save function
save(RSfilename="MSH_Rightsourcing_01DEC2019 to 28DEC2019.csv",
     JCfilename = "MSH_Rightsource Job Code Dict_01DEC2019 to 28DEC2019.csv")