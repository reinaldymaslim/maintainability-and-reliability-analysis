#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Extracts breakdown data from a single raw breakdown data file

# OUTLINE
# =======
# This function extracts the Equipment id, Breakdown time, RFU time, Problem, Breakdown duration, Breakdown HM, and RFU HM from a single .xlsx file containing raw breakdown events

# INPUT
# =====
#1. title = the full title of the file, e.g "KPI Week #27 (30 Jun - 06 Jul 2018).xlsx"
#2. path = the location of the file in local disk. for Windows 10, backslashes(\) should be replaced by (\\) and ending with (\\)
#          e.g "C:\\Users\\Reinaldy\\Desktop\\ITB\\MS\\KP\\MATERI\\"

# OUTPUT
# ======
# a dataframe containing {"Equipment","Breakdown Time","RFU Time","Problem","Actual.Duration.BD","HM.BD","HM.RFU"} columns of equipments breakdown data in the .xlsx file

#---------
# PROGRAM
#---------
weeklyBD<-function(title,path){
  #WeeklyBD is the basic function. This function reads one sheet from excel file of equipment breakdown and returns the cleaned dataframe. incorrect BD date is corrected by assuming RFU Date is always correct.
  #load packages required
  if (!require("pacman")) install.packages("pacman")
  pacman::p_load(readxl, tidyr, ggplot2, dplyr, magrittr)
  data<-data.frame(read_excel(paste0(path,title),sheet=6))
  data<-data[complete.cases(data),]#remove rows with NA data
  data<-data%>%
    separate("Breakdown.Time",c("year1","Breakdown.Hour"),sep=" ")%>%
    separate("RFU.Time",c("year2","RFU.Hour"),sep=" ")%>%
    unite("Breakdown Time","Breakdown.Date","Breakdown.Hour",sep=" ")%>%
    unite("RFU Time","RFU.Date","RFU.Hour",sep=" ")
  #change variables to keep here.
  colkeep1<-c("Equipment","Breakdown Time","RFU Time","Actual.Duration.BD","Problem","HM.BD","HM.RFU")
  data<-data[,colkeep1]
  data["RFU Time"]<-as.POSIXct(data[,"RFU Time"],format="%d.%m.%Y %H:%M:%S")
  data["Breakdown Time"]<-as.POSIXct(data[,"Breakdown Time"],format="%d.%m.%Y %H:%M:%S")
  #get difference btw. RFU Time and BD Time
  data["Diff1"]<-round(difftime(data[,"RFU Time"],data[,"Breakdown Time"],units="hours"),2) 
  data["Diff1">"Actual.Duration.BD","Breakdown Time"]<-24*floor(data["Diff1">"Actual.Duration.BD","Diff1"]/24)+data["Diff1">"Actual.Duration.BD","Breakdown Time"]#*24 becaues automatically converted to hours.
  #Diff2 to check
  data["Diff2"]<-round(difftime(data[,"RFU Time"],data[,"Breakdown Time"],units="hours"),2) 
  #change variables to keep here.
  colkeep2<-c("Equipment","Breakdown Time","RFU Time","Problem","Actual.Duration.BD","HM.BD","HM.RFU")
  #keep rows with correct dates
  data<-data["Actual.Duration.BD"!="Diff2",colkeep2]
}

#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Exctracts breakdown data from multiple raw breakdown data file

# OUTLINE
# =======
# This function binds dataframes returned by weeklyBD function and returns the combined dataframe.

# INPUT
# =====
#1. span = the span of weeks seperated by colon(:), e.g. 23:31.
#2. path = the location of the file in local disk. for Windows 10, backslashes(\) should be replaced by (\\) and ending with (\\)
#          e.g "C:\\Users\\Reinaldy\\Desktop\\ITB\\MS\\KP\\MATERI\\"

# OUTPUT
# ======
# a combined dataframe containing {"Equipment","Breakdown Time","RFU Time","Problem","Actual.Duration.BD","HM.BD","HM.RFU"} columns of multiple equipments breakdown data in the .xlsx file

#---------
# PROGRAM
#---------
binder<-function(span,path){
  #binder function binds dataframes returned by weeklyBD function and returns the combined dataframe.
  #initialize new dataframe.
  df<-data.frame() 
  for (i in span){
    nam<-paste0("KPI Week #",i)
    for (j in list.files()){if (grepl(nam,j) & !startsWith(j,"~")){df<-rbind(df,weeklyBD(j,path))}}
  }
  columnnames<-c("Equipment","BD.Time","RFU.Time","Problem","TTR","HM.BD","HM.RFU")
  colnames(df)<-columnnames
  #get unscheduled Breakdowns.
  df<-filter(df,Problem!="SCHEDULED/PLANNED")
  #discard duplicate rows.
  df<-unique(df[,columnnames])
}

#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Classifies breakdown data based on equipment types: {"HD","EXCT","EXKM"}.

# OUTLINE
# =======
# This function returns a list of dataframes based on equipment type.

# INPUT
# =====
#1. span = the span of weeks seperated by colon(:), e.g. 23:31.

# OUTPUT
# ======
# a list of dataframes containing breakdown data based on equipment types: {"HD","EXCT","EXKM"}.

#---------
# PROGRAM
#---------
equipmentBD1<-function(span=23:32){
  #equipmentBD1(builder1) function returns a list of dataframes based on equipment type.
  #Put path to folder containing the files here.
  path<-"C:\\Users\\Reinaldy\\Desktop\\ITB\\MS\\KP\\MATERI\\" 
  setwd(path)
  dflist<-list()
  df<-binder(span,path)#calls binder
  equipments<-c("HD","EXCT","EXKM")#add or subtract equipment from list
  for (i in seq_along(equipments)){
    dflist[[i]]<-subset(df,grepl(equipments[i],Equipment))#grab each row with unique equipment type to form dflist
  }
  dflist
}

#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Classifies breakdown data based on equipment id.

# OUTLINE
# =======
# This function returns a list of dataframes based on equipment id.

# INPUT
# =====
#1. span = the span of weeks seperated by colon(:), e.g. 23:32.

# OUTPUT
# ======
# a list of dataframes containing breakdown data based on equipment id.

#---------
# PROGRAM
#---------
equipmentBD2<-function(span=23:32){
  #equipmentBD2(builder2) function creates a list of dataframes based on each equipment.
  path<-"C:\\Users\\Reinaldy\\Desktop\\ITB\\MS\\KP\\MATERI\\" #Put path to folder containing the files here
  setwd(path)
  dflist<-list()
  df<-binder(span,path)#calls binder
  equipments<-unique(df$Equipment)
  for (i in seq_along(equipments)){
    dflist[[i]]<-subset(df,grepl(equipments[i],Equipment))#grab each row with unique equipment from the dataframe from binder to form dflist
    dflist[[i]]<-subset(dflist[[i]],!duplicated(dflist[[i]][,"BD.Time"]))
    dflist[[i]]<-dflist[[i]][order(dflist[[i]]["BD.Time"]),]#sorts the incorrect oder of Breakdown Time
    
  }
  dflist
}

#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Creates a pdf file containing pareto chart for the equipments specified in the input: {"HD","EXCT","EXKM"}.

# OUTLINE
# =======
# This function returns a saves pareto chart to a pdf file to the specified filepath.

# INPUT
# =====
#1. span = the span of weeks seperated by colon(:), e.g. 23:32.
#2. filename = the filename input as string, e.g. "HD".
#3. equipment = the equipment type with default types are: {"HD","EXCT","EXKM"}.

# OUTPUT
# ======
# a pdf file containing pareto chart for the equipments specified in the input: {"HD","EXCT","EXKM"}.

#---------
# PROGRAM
#---------
pareto<-function(span=23:32,filename,equipment=c("HD","EXCT","EXKM")){
  #this function saves pareto chart to a pdf file
  #input filename as string
  equipments<-vector(mode="list",length=3)
  names(equipments)<-c("HD","EXCT","EXKM")
  for (i in seq_along(equipments)){#builds a dictionary of equipments
    equipments[[i]]<-i
  }
  #specify filepath
  filepath<-"C:\\Users\\Reinaldy\\Desktop\\ITB\\MS\\KP\\MATERI"
  filepath<-paste0(filepath,"\\","pareto_",filename,".pdf")
  pdf(file=filepath)
  data<-equipmentBD1(span)#calls builder1
  for (i in seq_along(equipment)){
    chart_title<-paste0("Failure Cause Pareto of ",equipment[i]," Breakdown ")
    dat<-data.frame(count(data[[equipments[[equipment[i]]]]],Problem))
    dat<-dat[order(dat$n,decreasing=TRUE),]
    dat$Problem<-factor(dat$Problem,levels=dat$Problem)
    dat$cum<-cumsum(dat$n)
    print(ggplot(dat,aes(x=Problem))+
            geom_bar(aes(y=n,fill=n),stat="identity")+
            geom_text(stat="count", aes(label=n), vjust=-1)+
            geom_point(aes(y=cum))+
            geom_path(aes(y=cum,group=1))+
            labs(title = chart_title)+
            theme(plot.title = element_text(hjust = 0.5))+
            theme(axis.text.x=element_text(angle=45,hjust=1)))
    
  }
  dev.off()
}

#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Classifies breakdown data based on equipment id and sorts breakdown data based on breakdown counts.

# OUTLINE
# =======
# This function returns a list of dataframes based on equipment id sorted based on breakdown counts.

# INPUT
# =====
#1. span = the span of weeks seperated by colon(:), e.g. 23:32.

# OUTPUT
# ======
# a list of dataframes containing breakdown data based on equipment id sorted based on breakdown counts.

#---------
# PROGRAM
#---------
sortEquipmentBD<-function(span=23:32){
  #sortEquipmentBD returns a sorted list of dataframes containing Breakdown of each Equipment
  orderedList<-list()
  #calls equipmentBD2
  dflist<-equipmentBD2(span)
  #get a vector that contains nrow of each dataframe in dflist
  n<-sapply(dflist,nrow)
  #get the order to rearrange dflist based on decreasing nrow
  ord<-order(-n)
  for (i in seq_along(ord)){
    #order list of dataframe decreasing order of num of rows
    orderedList[[i]]<-dflist[[ord[i]]]
  }
  orderedList
}

#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Classifies breakdown data based on equipment id, sorts breakdown data based on breakdown counts and calculates Cumulative TTR, TBF and Cumulative TBF columns.

# OUTLINE
# =======
# This function returns a list of dataframes based on equipment id sorted based on breakdown counts with Cumulative TTR, TBF and Cumulative TBF columns.

# INPUT
# =====
#1. span = the span of weeks seperated by colon(:), e.g. 23:32.

# OUTPUT
# ======
# a list of dataframes containing breakdown data based on equipment id sorted based on breakdown counts with Cumulative TTR, TBF and Cumulative TBF columns.

#---------
# PROGRAM
#---------
equipmentTBF<-function(span){
  #calculate TBF, cumulative TTR, cumulative TBF
  #calls sortEquipmentBD
  dflist<-sortEquipmentBD(span)
  col_order<-c("Equipment","BD.Time","RFU.Time","Problem","HM.BD","HM.RFU","TTR", "Cumulative.TTR","TBF","Cumulative.TBF")
  for (i in seq_along(dflist)){
    numrow<-nrow(dflist[[i]])
    if (numrow>=2){
      #create TBF column
      dflist[[i]][2:numrow,"TBF"]<-as.numeric(round(difftime(dflist[[i]][2:numrow,"BD.Time"],dflist[[i]][1:numrow-1,"RFU.Time"],units="hours"),2))
      #create cumulative TTR column
      dflist[[i]]["Cumulative.TTR"]<-cumsum(dflist[[i]]["TTR"])
      #assign NA to new column "Cumulative TBF (H)"
      dflist[[i]]["Cumulative.TBF"]<-NA
      #create cumulative TBF column
      dflist[[i]]["Cumulative.TBF"][!is.na(dflist[[i]]["TBF"])]<-cumsum(dflist[[i]]["TBF"][!is.na(dflist[[i]]["TBF"])])
      #reorder columns
      dflist[[i]]<-dflist[[i]][,col_order]
      #modify cumulative TBF column as age of machine measured from the first timestamp in the first week of the data recorded in which case was 02.06.2018 06:00:00
      dflist[[i]]["Cumulative.TBF"]<-dflist[[i]]["Cumulative.TBF"]+as.numeric(round(difftime(dflist[[i]][1,"BD.Time"],as.POSIXct("02.06.2018 06:00:00",format="%d.%m.%Y %H:%M:%S"),units="hours"),2))
    }
    else{next}
  }
  dflist
}

compilerTBF<-function(dim,filename,equipment="HD"){
  #dim is the number of equipments required for analysis, starting from the most problematic one
  if (!require("pacman")) install.packages("pacman")
  pacman::p_load(xlsx)
  dflist<-equipmentTBF(23:32)#calls equipmentTBF
  df<-data.frame()#define new empty dataframe
  #define new function, cbind,fill to bind columns of matrix
  cbind.fill <- function(...){
    nm <- list(...) 
    nm <- lapply(nm, as.matrix)
    n <- max(sapply(nm, nrow)) 
    do.call(cbind, lapply(nm, function (x) 
      rbind(x, matrix(, n-nrow(x), ncol(x))))) 
  }
  count<-0
  names<-list()
  for (i in seq(dflist)){
    if (grepl(equipment,dflist[[i]][1,1])){
      df<-cbind.fill(df,dflist[[i]]["Cumulative.TBF"])#take the TBF of the equipment
      count<-count+1#counter for column names
      names[count]<-dflist[[i]][1,1]#fill names for each column
    }
    if (count==dim){break}
  }
  df<-as.data.frame(df)#convert df from matrix back to dataframe
  colnames(df)<-names#change the names of dataframe columns
  write.xlsx(df,filename)#write dataframe to .xlsx file
}

#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Calculates summary statistics for TTR data classified based on types of equipment: {"HD","EXCT","EXKM"}

# OUTLINE
# =======
# This function returns a list of dataframes containing summary statistics of TTR data classified based on equipment type.

# INPUT
# =====
#1. span = the span of weeks seperated by colon(:), e.g. 23:32.
#2. equipment = the equipment type with default types are: {"HD","EXCT","EXKM"}.

# OUTPUT
# ======
# a list of dataframes containing summary statistics of TTR data classified based on equipment type.
#---------
# PROGRAM
#---------
equipmentTTR<-function(span=23:32,equipment=c("HD","EXCT","EXKM")){
  #returns a list of dataframes containing the summary statistics of TTR based on equipment types: {"HD","EXCT","EXKM"}
  if (!require("pacman")) install.packages("pacman")
  pacman::p_load(data.table)
  #calls equipmentBD1
  df_list<-equipmentBD1(span)
  #builds a dictionary of equipments
  equipments<-vector(mode="list",length=3)
  names(equipments)<-c("HD","EXCT","EXKM")
  for (i in seq_along(equipments)){
    equipments[[i]]<-i
  }
  #builds a table containing the summary of TTR
  for (i in seq_along(equipment)){
    df_list[[equipments[[equipment[i]]]]]<-
      data.table(df_list[[equipments[[equipment[i]]]]])[,.(avg.TTR=mean(TTR),sd.TTR=sd(TTR),count.TTR=.N,sum.TTR=sum(TTR),max.TTR=max(TTR),min.TTR=min(TTR),mdn.TTR=median(TTR)),by=Problem][order(-avg.TTR)]
  }
  df_list
}

#------------
# DESCRIPTION
#------------
# PURPOSE
# =======
# Write summary statistics for TTR data classified based on types of equipment: {"HD","EXCT","EXKM"} to an .xlsx file.

# OUTLINE
# =======
# This function returns a .xlsx file containing summary statistics for TTR data classified based on types of equipment: {"HD","EXCT","EXKM"}.

# INPUT
# =====
#1. span = the span of weeks seperated by colon(:), e.g. 23:32.
#2. filename = the filename input as string, e.g. "TTRSummary".
#3. equipment = the equipment type with default types are: {"HD","EXCT","EXKM"}.

# OUTPUT
# ======
# a .xlsx file containing summary statistics for TTR data classified based on types of equipment: {"HD","EXCT","EXKM"}.
#---------
# PROGRAM
#---------
summaryTTR<-function(span=23:32,filename,equipment=c("HD","EXCT","EXKM")){
  if (!require("pacman")) install.packages("pacman")
  pacman::p_load(tidyr,xlsx)
  pdfpath<-"C:\\Users\\Reinaldy\\Desktop\\ITB\\MS\\KP\\MATERI"
  pdfpath<-paste0(pdfpath,"\\",filename,".pdf")
  pdf(file=pdfpath)
  filename<-paste0(filename,".xlsx")
  df_list<-equipmentTTR(span,equipment)
  for (i in seq_along(equipment)){
    df<-df_list[[i]]
    chart_title<-paste0("Repair Frequency and Average TTR of ",equipment[i]," Breakdown ")
    df1<-gather(df,event,values,c("avg.TTR","count.TTR"))
    print(ggplot(df1,aes(Problem,values,fill=event))+
            geom_bar(stat="identity",position="dodge")+
            labs(title = chart_title)+
            theme(plot.title = element_text(hjust = 0.5))+
            geom_text(aes(label=round(values,1)),position=position_dodge(0.9),color="black",size=2.5)+
            theme(axis.text.x=element_text(angle=45,hjust=1)))
    
    n<-df$Problem
    df2<-as.data.frame(t(df[,-1]))
    colnames(df2)<-n
    write.xlsx(df2,filename,sheetName=equipment[i],append=T)
  }
  dev.off()
}
