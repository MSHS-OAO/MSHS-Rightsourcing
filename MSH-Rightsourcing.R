###----------------------MSHS Rightsourcing

#Read raw excel file from rightsourcing - file needs to be "export table" format
message("Select most recent raw file")
right <- read.csv(file.choose(), fileEncoding = "UTF-16LE",sep='\t',header=T,stringsAsFactors = F)

##################################################################################################

#Site based function to create site based payroll, zero and jc dictionary
rightsourcing <- function(Site){
  
  library(anytime)
  #Read in Site's previous month zero file
  message("Select Site's most recent zero file")
  previous_zero <- read.csv(file.choose(),stringsAsFactors = F, header = F)
  previous_zero$V6 <- anytime(previous_zero$V6)
  #Read in Site's previous 2 month upload
  message("Select Site's most recent Rightsourcing upload")
  previous_site <- read.csv(file.choose(),stringsAsFactors = F, header = F)
  previous_site$V6 <- anytime(previous_site$V6)
  
  #if statement to tell code which site we are evaluating
  if(Site == "MSH"){
    i <-  1
  } else if(Site == "MSBI"){
    i <-  2
  }
  #set location and Hospital ID based on Site input
  Location <- c("MSH - Mount Sinai Hospital", "MSBI - Mount Sinai Beth Israel")
  Hospital <- c("NY0014", "630571")
  Loc <<- Location[i]
  Hosp <<- Hospital[i]
  
  
  library(dplyr)
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor")
  #read current job code list
  jc <- read.csv("Rightsource Job Code.csv",stringsAsFactors = F,header=F)
  colnames(jc) <- c("JobTitle","JobCode")
  #Column names for raw file
  colnames(right) <- c("Quarter",	"Months",	"ServiceLine",	"Worker",	"Manager",	"Supplier",
                       "JobTitle",	"JobCategory",	"JobClass",	"ClientBill Rate",	"Dept",
                       "DeptLvl1","NewFill?","DeptLvl 2",	"Location",	"Hours",	"Spend",	"EarningsE/D",
                       "AllOtherSpend",	"OTSpend")
  
  #replace N/A job titles with unkown job title
  right <- tibble::as_tibble(right)
  right <- right %>% mutate(JobTitle = replace(JobTitle,JobTitle == "#N/A","Unknown"))
  
  #Create variable for max date in previous months zero and system file
  library(anytime)
  max_zero <<- max(anytime(previous_zero[,7]))
  max_site <<- max(anytime(previous_site[,7]))
  
  ###########Create site level job code dictionary
  right$`EarningsE/D` <- anytime(right$`EarningsE/D`)
  right <- right[right$`EarningsE/D` > max_zero & right$Location == Loc,]
  
  #replace workers name with just their last name
  #right$Worker <- gsub(",.*$","",right$Worker)
  
  #add "Rightsource" to the the job title
  right$JobTitle <- paste("Rightsourcing",right$JobTitle,sep=" ")
  #subset all rows with new job codes
  newright <- subset(right,!(right$JobTitle %in% jc$JobTitle))
  #append new jobs only if there are new jobs
  if(nrow(newright) != 0){
    #create vector of unique new jobs
    newjobs <- unique(newright$JobTitle)
    #append the new jobs to the jobcode list
    #i from job list length plus 1 to length of new jobs
    a <- nrow(jc)
    for(j in a+1:length(newjobs)){
      #append with job title and job code for all new jobs
      #if job is over the 99th job, adjust the coding scheme
      jc[j,] <- list(newjobs[j-a],if(j>99){
        paste("R00",j,sep="")
      }else{
        paste("R000",j,sep="")
      })
    }
    #Overwrit Jobcode table
    setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/")
    write.table(jc,file="Rightsource Job Code.csv",sep=",",col.names=F,row.names=F)
  }
  #append job code to prepare fo jobcode dictionary
  newrightjc <<- merge(newright,jc,by="JobTitle")
  #only create job code dictionary if there are new job codes
  if(nrow(newrightjc) != 0){
    if(i == 1){
      #Creat rightsourcing jobcode dictionary from newrightjc
      jcdict1 <- data.frame(partner="729805", hospital=Hosp,dep=substr(newrightjc$Dept,start=1,stop=8),
                            jc=newrightjc$JobCode,title=newrightjc$JobTitle)
      #Only take unique rows from jcdict
      jcdict1 <- jcdict1[!duplicated(jcdict1),]
      jcdict1 <<- jcdict1
    } else if(i == 2){
      #Adjust all Beth Israel Cost centers 
      library(stringr)
      newrightjc$Dept[str_length(newrightjc$Dept)==30] <- str_c(
        str_sub(newrightjc$Dept[str_length(newrightjc$Dept)==30], 1, 4),
        str_sub(newrightjc$Dept[str_length(newrightjc$Dept)==30], 13, 14),
        str_sub(newrightjc$Dept[str_length(newrightjc$Dept)==30], 16, 19))
      
      newrightjc$Dept[str_length(newrightjc$Dept)==32] <- str_c(
        str_sub(newrightjc$Dept[str_length(newrightjc$Dept)==32], 1, 4),
        str_sub(newrightjc$Dept[str_length(newrightjc$Dept)==32], 14, 15),
        str_sub(newrightjc$Dept[str_length(newrightjc$Dept)==32], 17, 20))
      
      #Creat rightsourcing jobcode dictionary from newrightjc
      jcdict2 <- data.frame(partner="729805", hospital=Hosp,dep=newrightjc$Dept,
                            jc=newrightjc$JobCode,title=newrightjc$JobTitle)
      #Only take unique rows from jcdict
      jcdict2 <- jcdict2[!duplicated(jcdict2),]
      jcdict2 <<- jcdict2
    }
  }
  #merge the rightsourcing jobcode dictionary and the rightsourcing file
  #essentially assigns jobcode based on job title
  right <- merge(right,jc,by="JobTitle")
  
  ###########Create site level job code dictionary
  if(i == 1){
    export1 <-  data.frame(partner="729805",hospital=Hosp,home="01010101",
                           hosp=Hosp,work=substr(right$Dept,start=1,stop=8),start=as.Date(right$`EarningsE/D`)-6,
                           end=as.Date(right$`EarningsE/D`),EmpCode=paste0(substr(right$Worker,start=1,stop=12),gsub("//..*","",right$Hours)),
                           name=right$Worker,budget="0",JobCode=right$JobCode,paycode="AG1",
                           hours=right$Hours,spend=right$Spend) 
    export1$start <- paste(substr(export1$start,start=6,stop=7),"/",substr(export1$start,start=9,stop=10),
                           "/",substr(export1$start,start=1,stop=4),sep="")
    export1$end <- paste(substr(export1$end,start=6,stop=7),"/",substr(export1$end,start=9,stop=10),
                         "/",substr(export1$end,start=1,stop=4),sep="")
    export1 <<- export1
  } else if(i == 2){
    #Adjust all Beth Israel Cost centers 
    library(stringr)
    right$Dept[str_length(right$Dept)==30] <- str_c(
      str_sub(right$Dept[str_length(right$Dept)==30], 1, 4),
      str_sub(right$Dept[str_length(right$Dept)==30], 13, 14),
      str_sub(right$Dept[str_length(right$Dept)==30], 16, 19))
    
    right$Dept[str_length(right$Dept)==32] <- str_c(
      str_sub(right$Dept[str_length(right$Dept)==32], 1, 4),
      str_sub(right$Dept[str_length(right$Dept)==32], 14, 15),
      str_sub(right$Dept[str_length(right$Dept)==32], 17, 20))
    
    export2 <-  data.frame(partner="729805",hospital=Hosp,home="1010101010",
                           hosp=Hosp,work=right$Dept,start=as.Date(right$`EarningsE/D`)-6,
                           end=as.Date(right$`EarningsE/D`),EmpCode=paste0(substr(right$Worker,start=1,stop=12),gsub("//..*","",right$Hours)),
                           name=right$Worker,budget="0",JobCode=right$JobCode,paycode="AG1",
                           hours=right$Hours,spend=right$Spend)
    export2$start <- paste(substr(export2$start,start=6,stop=7),"/",substr(export2$start,start=9,stop=10),
                           "/",substr(export2$start,start=1,stop=4),sep="")
    export2$end <- paste(substr(export2$end,start=6,stop=7),"/",substr(export2$end,start=9,stop=10),
                         "/",substr(export2$end,start=1,stop=4),sep="")
    export2 <<- export2
  }
  
  ###########Create site level zero file
  if(i == 1){
    zero1 <- previous_site[previous_site$V6 > max_zero,]
    zero1[,13:14] <- 0
    zero1$V6 <- paste(substr(zero1$V6,start=6,stop=7),"/",substr(zero1$V6,start=9,stop=10),
                      "/",substr(zero1$V6,start=1,stop=4),sep="")
    zero1 <<- zero1
  } else if(i == 2){
    zero2 <- previous_site[previous_site$V6 > max_zero,]
    zero2[,13:14] <- 0
    zero2$V6 <- paste(substr(zero2$V6,start=6,stop=7),"/",substr(zero2$V6,start=9,stop=10),
                      "/",substr(zero2$V6,start=1,stop=4),sep="")
    zero2 <<- zero2
  }
}
#system function combines site based payroll, zero and jc dictionary exports in system level
system <- function(){
  if(exists("export1") & exists("export2")){
    sys_export <<- rbind(export1,export2)
  }
  if(exists("jcdict1") & exists("jcdict2")){
    sys_jcdict <<- rbind(jcdict1,jcdict2)
  }
  if(exists("zero1") & exists("zero2")){
    sys_zero <<- rbind(zero1,zero2)
  }
}
##Create save function for system and site exports if they exist
save <- function(){
  library(anytime)
  #save system and site exports if they exist
  if(exists("sys_export")){
    start <- min(anytime(sys_export$start))
    end <- max(anytime(sys_export$end))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSHS/",
                   "MSHS_Rightsourcing_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(sys_export,file=name,sep=",",col.names=F,row.names=F)
  }
  if(exists("export1")){
    start <- min(anytime(export1$start))
    end <- max(anytime(export1$end))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSH/",
                   "MSH_Rightsourcing_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(export1,file=name,sep=",",col.names=F,row.names=F)
  } 
  if(exists("export2")){
    start <- min(anytime(export2$start))
    end <- max(anytime(export2$end))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSBI/",
                   "MSBI_Rightsourcing_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(export2,file=name,sep=",",col.names=F,row.names=F)
  }
  #save system and site zero if they exist
  if(exists("sys_zero")){
    start <- min(anytime(sys_zero$V6))
    end <- max(anytime(sys_zero$V7))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSHS/Zero/",
                   "MSHS_Rightsourcing Zero_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(sys_zero,file=name,sep=",",col.names=F,row.names=F)
  }
  if(exists("zero1")){
    start <- min(anytime(zero1$V6))
    end <- max(anytime(zero1$V7))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSH/Zero/",
                   "MSH_Rightsourcing Zero_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(zero1,file=name,sep=",",col.names=F,row.names=F)
  } 
  if(exists("zero2")){
    start <- min(anytime(zero2$V6))
    end <- max(anytime(zero2$V7))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSBI/Zero/",
                   "MSBI_Rightsourcing Zero_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(zero2,file=name,sep=",",col.names=F,row.names=F)
  }
  #save system and site jcdict if they exist
  if(exists("sys_jcdict")){
    date <- as.Date(Sys.Date(), format = "%Y-%m-%d")
    library(lubridate)
    month <- toupper(month.abb[month(date)])
    day <- format(as.Date(date, formate = "%Y-%m-%d"), format="%d")
    year <- substr(date, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSHS/JCdict/",
                   "MSHS_Job Code Dictionary_",day,month,year,".csv")
    write.table(sys_jcdict,file=name,sep=",",col.names=F,row.names=F)
  }
  if(exists("jcdict1")){
    date <- as.Date(Sys.Date(), format = "%Y-%m-%d")
    library(lubridate)
    month <- toupper(month.abb[month(date)])
    day <- format(as.Date(date, formate = "%Y-%m-%d"), format="%d")
    year <- substr(date, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSH/JCdict/",
                   "MSH_Job Code Dictionary_",day,month,year,".csv")
    write.table(jcdict1,file=name,sep=",",col.names=F,row.names=F)
  }
  if(exists("jcdict2")){
    date <- as.Date(Sys.Date(), format = "%Y-%m-%d")
    library(lubridate)
    month <- toupper(month.abb[month(date)])
    day <- format(as.Date(date, formate = "%Y-%m-%d"), format="%d")
    year <- substr(date, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSBI/JCdict/",
                   "MSBI_Job Code Dictionary_",day,month,year,".csv")
    write.table(jcdict2,file=name,sep=",",col.names=F,row.names=F)
  }
}

##################################################################################################

#Create Site based payroll, zero and jc dictionary
rightsourcing("MSH")
rightsourcing("MSBI")

#if rightsourcing() was run for multiple sites then create system payroll, zero and jc dictionary
system()

##Check all exports
#Saves all site and system exports if they exist
save()
