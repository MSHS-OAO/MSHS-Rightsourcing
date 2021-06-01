###----------------------MSHS Rightsourcing

library(anytime)
library(tidyr)

def_dir <- "J:\\deans\\Presidents\\SixSigma\\MSHS Productivity\\Productivity\\Labor - Data\\Rightsourcing Labor\\Source Data"

#Read raw excel file from rightsourcing - file needs to be "export table" format
message("Select most recent raw file")
right <- read.csv(choose.files(default = def_dir), fileEncoding = "UTF-16LE",sep='\t',header=T,stringsAsFactors = F)
#right <- read.csv(file.choose(),header = T,stringsAsFactors = F)

##################################################################################################
#Create empty list for all dictionaries
dict <- vector(mode = "list", length = 4)
names(dict) <- c("MSH","MSBI","MSQ","MSB")
#Create empty list for all rightsourcing exports
export <- vector(mode = "list", length = 4)
names(export) <- c("MSH","MSBI","MSQ","MSB")
#Create empty list for all zero tables
zero <- vector(mode = "list", length = 4)
names(zero) <- c("MSH","MSBI","MSQ","MSB")

#Site based function to create site based payroll, zero and jc dictionary
rightsourcing <- function(Site){
  
  #Read in Site's previous month zero file
  message("Select Site's most recent zero file")
  previous_zero <- read.csv(choose.files(default = def_dir),stringsAsFactors = F, header = F)
  previous_zero$V6 <- anytime(previous_zero$V6)
  #Read in Site's previous 2 month upload
  message("Select Site's most recent Rightsourcing upload")
  previous_site <- read.csv(choose.files(default = def_dir),stringsAsFactors = F, header = F)
  previous_site$V6 <- anytime(previous_site$V6)
  
  #if statement to tell code which site we are evaluating
  if(Site == "MSH"){
    i <-  1
    conversion <- read.csv("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSHQ Code Conversion.csv",colClasses = c("character","character"))
  } else if(Site == "MSBI"){
    i <-  2
    conversion <- read.csv("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSBIB Code Conversion.csv",colClasses = c("character","character"))
  } else if(Site == "MSQ"){
    i <- 3
    conversion <- read.csv("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSHQ Code Conversion.csv",colClasses = c("character","character"))
  } else if(Site == "MSB"){
    i <- 4
    conversion <- read.csv("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSBIB Code Conversion.csv",colClasses = c("character","character"))
  }
  
  #set location and Hospital ID based on Site input
  Location <- c("MSH - Mount Sinai Hospital", "MSBI - Mount Sinai Beth Israel", "MSQ - Mount Sinai Queens", "MSB - Mount Sinai Brooklyn")
  Hospital <- c("NY0014", "630571", "NY0014", "630571")
  Loc <<- Location[i]
  Hosp <<- Hospital[i]
  
  
  library(dplyr)
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor")
  #read current job code list
  jc <- read.csv("Rightsource Job Code.csv",stringsAsFactors = F,header=F)
  colnames(jc) <- c("JobTitle","JobCode")
  #Column names for raw file
  colnames(right) <- c("Quarter",	"Months",	"Location",	"Dept",	"Earnings.E.D",	"ServiceLine",
                       "JobTitle",	"JobCategory",	"JobClass",	"Hours",	"Spend",
                       "Supplier","Worker","Manager",	"ClientBill Rate",	"DeptLvl1",	"NewFill?",	"DeptLvl 2",
                       "AllOtherSpend",	"OTSpend","COVID19 Worker", "COVID19 Spend", "Year")
  
  #replace N/A job titles with unkown job title
  right <- tibble::as_tibble(right)
  right <- right %>% mutate(JobTitle = replace(JobTitle,JobTitle == "#N/A","Unknown"))
  
  #Create variable for max date in previous months zero and system file
  library(anytime)
  max_zero <<- max(anytime(previous_zero[,7]))
  max_site <<- max(anytime(previous_site[,7]))
  
  ###########filter raw file on proper dates
  right <- right %>% 
    mutate(`Earnings.E.D` = anytime(`Earnings.E.D`)) %>%
    filter(`Earnings.E.D` > max_zero,
           Location == Loc)
  
  ###########Create site level job code dictionary
  #replace workers name with just their last name
  right$Worker <- gsub(",.*$","",right$Worker)
  #add "Rightsource" to the the job title
  right$JobTitle <- paste("Rightsourcing",right$JobTitle,sep=" ")
  #subset all rows with new job codes
  newright <- subset(right,!(right$JobTitle %in% jc$JobTitle))
  #create jobcode dictionary if there are new jobcodes
  if(nrow(newright) != 0){
    if(i == 1){
      #MSH
      jcdict1 <<- newright %>% 
        mutate(SYSTEM = 729805, HOSP = "NY0014", Dept = substr(Dept,1,8), JC.Description = "") %>%
        select(SYSTEM,HOSP,Dept,JobTitle) %>%
        distinct()
      jcdict1 <-  jcdict1 %>%
        left_join(conversion,by=c("Dept"="Dept")) %>%
        mutate(Oracle = case_when(
          is.na(Oracle) ~ "101010101010101",
          Oracle == "DELETE" ~ "101010101010101",
          TRUE ~ Oracle)) %>%
        select(SYSTEM,HOSP,Oracle,JobTitle) %>%
        rename(Dept = Oracle)
      dict[[i]] <- jcdict1
      dict <<- dict
    } else if(i == 2){
      #MSBI
      library(stringr)
      newright$Dept[str_length(newright$Dept)==30] <- str_c(
        str_sub(newright$Dept[str_length(newright$Dept)==30], 1, 4),
        str_sub(newright$Dept[str_length(newright$Dept)==30], 13, 14),
        str_sub(newright$Dept[str_length(newright$Dept)==30], 16, 19))
      
      newright$Dept[str_length(newright$Dept)==32] <- str_c(
        str_sub(newright$Dept[str_length(newright$Dept)==32], 1, 4),
        str_sub(newright$Dept[str_length(newright$Dept)==32], 14, 15),
        str_sub(newright$Dept[str_length(newright$Dept)==32], 17, 20))
      jcdict2 <<- newright %>% 
        mutate(SYSTEM = 729805, HOSP = "630571") %>%
        select(SYSTEM,HOSP,Dept,JobTitle) %>%
        distinct()
      jcdict2 <-  jcdict2 %>%
        left_join(conversion,by=c("Dept"="Dept")) %>%
        mutate(Oracle = case_when(
          is.na(Oracle) ~ "900000040490000",
          Oracle == "DELETE" ~ "900000040490000",
          TRUE ~ Oracle)) %>%
        select(SYSTEM,HOSP,Oracle,JobTitle) %>%
        rename(Dept = Oracle)
      dict[[i]] <- jcdict2
      dict <<- dict
    } else if(i == 3){
      #MSQ
      jcdict3 <<- newright %>% 
        mutate(SYSTEM = 729805, HOSP = "NY0014", Dept = substr(Dept,1,8), JC.Description = "") %>%
        select(SYSTEM,HOSP,Dept,JobTitle) %>%
        distinct()
      jcdict3 <-  jcdict3 %>%
        left_join(conversion,by=c("Dept"="Dept")) %>%
        mutate(Oracle = case_when(
          is.na(Oracle) ~ "101010101010102",
          Oracle == "DELETE" ~ "101010101010102",
          TRUE ~ Oracle)) %>%
        select(SYSTEM,HOSP,Oracle,JobTitle) %>%
        rename(Dept = Oracle)
      dict[[i]] <- jcdict3
      dict <<- dict
    } else if(i == 4){
      #MSB
      library(stringr)
      newright$Dept[str_length(newright$Dept)==30] <- str_c(
        str_sub(newright$Dept[str_length(newright$Dept)==30], 1, 4),
        str_sub(newright$Dept[str_length(newright$Dept)==30], 13, 14),
        str_sub(newright$Dept[str_length(newright$Dept)==30], 16, 19))
      
      newright$Dept[str_length(newright$Dept)==32] <- str_c(
        str_sub(newright$Dept[str_length(newright$Dept)==32], 1, 4),
        str_sub(newright$Dept[str_length(newright$Dept)==32], 14, 15),
        str_sub(newright$Dept[str_length(newright$Dept)==32], 17, 20))
      jcdict4 <<- newright %>% 
        mutate(SYSTEM = 729805, HOSP = "630571") %>%
        select(SYSTEM,HOSP,Dept,JobTitle) %>%
        distinct() 
      jcdict4<-  jcdict4 %>%
        left_join(conversion,by=c("Dept"="Dept")) %>%
        mutate(Oracle = case_when(
          is.na(Oracle) ~ "900000040790000",
          Oracle == "DELETE" ~ "900000040790000",
          TRUE ~ Oracle)) %>%
        select(SYSTEM,HOSP,Oracle,JobTitle) %>%
        rename(Dept = Oracle)
      dict[[i]] <- jcdict4
      dict <<- dict
    }
  }
  
  #merge the rightsourcing jobcode dictionary and the rightsourcing file
  #essentially assigns jobcode based on job title
  right <- left_join(right,jc,by = c("JobTitle"))
  
  ###########Create Site level upload
  if(i == 1){
    #MSH
    library(stringr)
    export1 <-  data.frame(partner="729805",hospital=Hosp,home="101010101010101",
                           hosp=Hosp,work=substr(right$Dept,start=1,stop=8),start=as.Date(right$`Earnings.E.D`)-6,
                           end=as.Date(right$`Earnings.E.D`),EmpCode=paste0(substr(right$Worker,start=1,stop=12),str_extract(right$Hours,"[^.]+")),
                           name=right$Worker,budget="0",JobCode=right$JobCode,paycode="AG1",
                           hours=right$Hours,spend=right$Spend,JobTitle=right$JobTitle)
    export1$EmpCode <- substr(export1$EmpCode,start=1,stop=15)
    export1$start <- paste(substr(export1$start,start=6,stop=7),"/",substr(export1$start,start=9,stop=10),
                           "/",substr(export1$start,start=1,stop=4),sep="")
    export1$end <- paste(substr(export1$end,start=6,stop=7),"/",substr(export1$end,start=9,stop=10),
                         "/",substr(export1$end,start=1,stop=4),sep="")
    export1$work <- as.character(export1$work)
    test <- export1 %>%
      left_join(conversion,by=c("work" = "Dept")) %>%
      select(c(1:4,16,6:15)) %>%
      rename(work = Oracle) %>%
      mutate(work = case_when(
        is.na(work) ~ "101010101010101",
        work == "DELETE" ~ "101010101010101",
        TRUE ~ work))
    export1 <<- export1
    export[[i]] <- export1
    export <<- export
  } else if(i == 2){
    #MSBI
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
    
    export2 <-  data.frame(partner="729805",hospital=Hosp,home="900000040490000",
                           hosp=Hosp,work=right$Dept,start=as.Date(right$`Earnings.E.D`)-6,
                           end=as.Date(right$`Earnings.E.D`),EmpCode=paste0(substr(right$Worker,start=1,stop=12),str_extract(right$Hours,"[^.]+")),
                           name=right$Worker,budget="0",JobCode=right$JobCode,paycode="AG1",
                           hours=right$Hours,spend=right$Spend,JobTitle=right$JobTitle)
    export2$EmpCode <- substr(export2$EmpCode,start=1,stop=15)
    export2$start <- paste(substr(export2$start,start=6,stop=7),"/",substr(export2$start,start=9,stop=10),
                           "/",substr(export2$start,start=1,stop=4),sep="")
    export2$end <- paste(substr(export2$end,start=6,stop=7),"/",substr(export2$end,start=9,stop=10),
                         "/",substr(export2$end,start=1,stop=4),sep="")
    export2 <- export2 %>%
      left_join(conversion,by=c("work" = "Dept")) %>%
      select(c(1:4,16,6:15)) %>%
      rename(work = Oracle) %>%
      mutate(work = case_when(
        is.na(work) ~ "900000040490000",
        work == "DELETE" ~ "900000040490000",
        TRUE ~ work))
    export2 <<- export2
    export[[i]] <- export2
    export <<- export
  } else if(i==3){
    #MSQ
    library(stringr)
    export3 <-  data.frame(partner="729805",hospital=Hosp,home="101010101010102",
                           hosp=Hosp,work=substr(right$Dept,start=1,stop=8),start=as.Date(right$`Earnings.E.D`)-6,
                           end=as.Date(right$`Earnings.E.D`),EmpCode=paste0(substr(right$Worker,start=1,stop=12),str_extract(right$Hours,"[^.]+")),
                           name=right$Worker,budget="0",JobCode=right$JobCode,paycode="AG1",
                           hours=right$Hours,spend=right$Spend,JobTitle=right$JobTitle)
    export3$EmpCode <- substr(export3$EmpCode,start=1,stop=15)
    export3$start <- paste(substr(export3$start,start=6,stop=7),"/",substr(export3$start,start=9,stop=10),
                           "/",substr(export3$start,start=1,stop=4),sep="")
    export3$end <- paste(substr(export3$end,start=6,stop=7),"/",substr(export3$end,start=9,stop=10),
                         "/",substr(export3$end,start=1,stop=4),sep="")
    export3$work <- as.character(export3$work)
    export3 <- export3 %>%
      left_join(conversion,by=c("work" = "Dept")) %>%
      select(c(1:4,16,6:15)) %>%
      rename(work = Oracle) %>%
      mutate(work = case_when(
        is.na(work) ~ "101010101010102",
        work == "DELETE" ~ "101010101010102",
        TRUE ~ work))
    export3 <<- export3
    export[[i]] <- export3
    export <<- export
  } else if(i ==4){
    #MSB
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
    
    export4 <-  data.frame(partner="729805",hospital=Hosp,home="900000040790000",
                           hosp=Hosp,work=right$Dept,start=as.Date(right$`Earnings.E.D`)-6,
                           end=as.Date(right$`Earnings.E.D`),EmpCode=paste0(substr(right$Worker,start=1,stop=12),str_extract(right$Hours,"[^.]+")),
                           name=right$Worker,budget="0",JobCode=right$JobCode,paycode="AG1",
                           hours=right$Hours,spend=right$Spend,JobTitle=right$JobTitle)
    export4$EmpCode <- substr(export4$EmpCode,start=1,stop=15)
    export4$start <- paste(substr(export4$start,start=6,stop=7),"/",substr(export4$start,start=9,stop=10),
                           "/",substr(export4$start,start=1,stop=4),sep="")
    export4$end <- paste(substr(export4$end,start=6,stop=7),"/",substr(export4$end,start=9,stop=10),
                         "/",substr(export4$end,start=1,stop=4),sep="")
    export4 <- export4 %>%
      left_join(conversion,by=c("work" = "Dept")) %>%
      select(c(1:4,16,6:15)) %>%
      rename(work = Oracle) %>%
      mutate(work = case_when(
        is.na(work) ~ "900000040790000",
        work == "DELETE" ~ "900000040790000",
        TRUE ~ work))
    export4 <<- export4
    export[[i]] <- export4
    export <<- export
  }
  
  ###########Create site level zero file
  if(i == 1){
    #MSH
    zero1 <- previous_site[previous_site$V6 > max_zero,]
    zero1[,13:14] <- 0
    zero1$V6 <- paste(substr(zero1$V6,start=6,stop=7),"/",substr(zero1$V6,start=9,stop=10),
                      "/",substr(zero1$V6,start=1,stop=4),sep="")
    zero1 <<- zero1
    zero[[i]] <- zero1
    zero <<- zero
  } else if(i == 2){
    #MSBI
    zero2 <- previous_site[previous_site$V6 > max_zero,]
    zero2[,13:14] <- 0
    zero2$V6 <- paste(substr(zero2$V6,start=6,stop=7),"/",substr(zero2$V6,start=9,stop=10),
                      "/",substr(zero2$V6,start=1,stop=4),sep="")
    zero2 <<- zero2
    zero[[i]] <- zero2
    zero <<- zero
  } else if(i == 3){
    #MSQ
    zero3 <- previous_site[previous_site$V6 > max_zero,]
    zero3[,13:14] <- 0
    zero3$V6 <- paste(substr(zero3$V6,start=6,stop=7),"/",substr(zero3$V6,start=9,stop=10),
                      "/",substr(zero3$V6,start=1,stop=4),sep="")
    zero3 <<- zero3
    zero[[i]] <- zero3
    zero <<- zero
  } else if(i == 4){
    #MSB
    zero4 <- previous_site[previous_site$V6 > max_zero,]
    zero4[,13:14] <- 0
    zero4$V6 <- paste(substr(zero4$V6,start=6,stop=7),"/",substr(zero4$V6,start=9,stop=10),
                      "/",substr(zero4$V6,start=1,stop=4),sep="")
    zero4 <<- zero4
    zero[[i]] <- zero4
    zero <<- zero
    }
}
#function to compare all new job titles and properly append job list/create jobcode dictionaries
dictionary <- function(){
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/")
  jc <- read.csv("Rightsource Job Code.csv",stringsAsFactors = F,header=F)
  colnames(jc) <- c("JobTitle","JobCode")
  sys_dict = do.call("rbind",dict)
  if(!is.null(sys_dict)){
    sys_jc <- sys_dict %>% 
      select(JobTitle) %>%
      distinct()
    if(nrow(sys_jc) != 0){
      newjobs <- as.vector(sys_jc$JobTitle)
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
      # #Overwrite Jobcode table
      # setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/")
      # write.table(jc,file="Rightsource Job Code.csv",sep=",",col.names=F,row.names=F)
    }
    for(x in 1:length(dict)){
      if(!is.null(dict[[x]])){
        dict[[x]] <- left_join(dict[[x]],jc,by=c("JobTitle")) %>%
          select(SYSTEM,HOSP,Dept,JobCode,JobTitle)
        dict <<- dict
      }
    }
  }
  for(y in 1:length(export)){
    if(!is.null(export[[y]])){
      export[[y]] <- left_join(export[[y]],jc,by=c("JobTitle")) %>%
        select(partner,hospital,home,hosp,work,start,end,EmpCode,name,budget,16,paycode,hours,spend)
      export <<- export
    }
  }
  return(jc)
}
#system function combines site based payroll, zero and jc dictionary exports in system level
system <- function(){
  if(length(export) != 0){
    sys_export = do.call("rbind",export)
    sys_export <<- sys_export
  }
  if(length(dict) != 0){
    sys_jcdict = do.call("rbind",dict)
    sys_jcdict <<- sys_jcdict
  }
  if(length(zero) != 0){
    sys_zero = do.call("rbind",zero)
    sys_zero <<- sys_zero
  }
}
##Create save function for system and site exports if they exist
save <- function(){
  library(anytime)
  #save new rightsourcing jobcode list
  write.table(jc,file="J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/Rightsource Job Code.csv",sep=",",col.names=F,row.names=F)
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
  if(is.data.frame(export[[1]])){
    start <- min(anytime(export[[1]]$start))
    end <- max(anytime(export[[1]]$end))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSH/",
                   "MSH_Rightsourcing_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(export[[1]],file=name,sep=",",col.names=F,row.names=F)
  } 
  if(is.data.frame(export[[2]])){
    start <- min(anytime(export[[2]]$start))
    end <- max(anytime(export[[2]]$end))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSBI/",
                   "MSBI_Rightsourcing_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(export[[2]],file=name,sep=",",col.names=F,row.names=F)
  }
  if(is.data.frame(export[[3]])){
    start <- min(anytime(export[[3]]$start))
    end <- max(anytime(export[[3]]$end))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSQ/",
                   "MSQ_Rightsourcing_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(export[[3]],file=name,sep=",",col.names=F,row.names=F)
  }
  if(is.data.frame(export[[4]])){
    start <- min(anytime(export[[4]]$start))
    end <- max(anytime(export[[4]]$end))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSB/",
                   "MSB_Rightsourcing_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(export[[4]],file=name,sep=",",col.names=F,row.names=F)
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
  if(is.data.frame(zero[[1]])){
    start <- min(anytime(zero[[1]]$V6))
    end <- max(anytime(zero[[1]]$V7))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSH/Zero/",
                   "MSH_Rightsourcing Zero_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(zero[[1]],file=name,sep=",",col.names=F,row.names=F)
  } 
  if(is.data.frame(zero[[2]])){
    start <- min(anytime(zero[[2]]$V6))
    end <- max(anytime(zero[[2]]$V7))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSBI/Zero/",
                   "MSBI_Rightsourcing Zero_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(zero[[2]],file=name,sep=",",col.names=F,row.names=F)
  }
  if(is.data.frame(zero[[3]])){
    start <- min(anytime(zero[[3]]$V6))
    end <- max(anytime(zero[[3]]$V7))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSQ/Zero/",
                   "MSQ_Rightsourcing Zero_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(zero[[3]],file=name,sep=",",col.names=F,row.names=F)
  }
  if(is.data.frame(zero[[4]])){
    start <- min(anytime(zero[[4]]$V6))
    end <- max(anytime(zero[[4]]$V7))
    library(lubridate)
    smonth <- toupper(month.abb[month(start)])
    emonth <- toupper(month.abb[month(end)])
    sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
    eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
    syear <- substr(start, start=1, stop=4)
    eyear <- substr(end, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSB/Zero/",
                   "MSB_Rightsourcing Zero_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
    write.table(zero[[4]],file=name,sep=",",col.names=F,row.names=F)
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
  if(is.data.frame(dict[[1]])){
    date <- as.Date(Sys.Date(), format = "%Y-%m-%d")
    library(lubridate)
    month <- toupper(month.abb[month(date)])
    day <- format(as.Date(date, formate = "%Y-%m-%d"), format="%d")
    year <- substr(date, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSH/JCdict/",
                   "MSH_Job Code Dictionary_",day,month,year,".csv")
    write.table(dict[[1]],file=name,sep=",",col.names=F,row.names=F)
  }
  if(is.data.frame(dict[[2]])){
    date <- as.Date(Sys.Date(), format = "%Y-%m-%d")
    library(lubridate)
    month <- toupper(month.abb[month(date)])
    day <- format(as.Date(date, formate = "%Y-%m-%d"), format="%d")
    year <- substr(date, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSBI/JCdict/",
                   "MSBI_Job Code Dictionary_",day,month,year,".csv")
    write.table(dict[[2]],file=name,sep=",",col.names=F,row.names=F)
  }
  if(is.data.frame(dict[[3]])){
    date <- as.Date(Sys.Date(), format = "%Y-%m-%d")
    library(lubridate)
    month <- toupper(month.abb[month(date)])
    day <- format(as.Date(date, formate = "%Y-%m-%d"), format="%d")
    year <- substr(date, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSQ/JCdict/",
                   "MSQ_Job Code Dictionary_",day,month,year,".csv")
    write.table(dict[[3]],file=name,sep=",",col.names=F,row.names=F)
  }
  if(is.data.frame(dict[[4]])){
    date <- as.Date(Sys.Date(), format = "%Y-%m-%d")
    library(lubridate)
    month <- toupper(month.abb[month(date)])
    day <- format(as.Date(date, formate = "%Y-%m-%d"), format="%d")
    year <- substr(date, start=1, stop=4)
    name <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/Rightsourcing Labor/MSB/JCdict/",
                   "MSB_Job Code Dictionary_",day,month,year,".csv")
    write.table(dict[[4]],file=name,sep=",",col.names=F,row.names=F)
  }
}

##################################################################################################

#Create Site based payroll, zero and jc dictionary
rightsourcing("MSH")
rightsourcing("MSBI")
rightsourcing("MSQ")
rightsourcing("MSB")

#Create new job codes and incorporate to dictionaries and uploads
jc <- dictionary()

#if rightsourcing() was run for multiple sites then create system payroll, zero and jc dictionary
system()

##Check all exports
#Saves all site and system exports if they exist
save()
