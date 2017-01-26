# WEEKLY PERFORMANCE REPORT. CLEANED.


# set the beginning of last week
start <- as.Date(Sys.time()-777600)
start
library(dplyr)
library(xlsx)
library(lubridate)

setwd("C:/Programs/gtc_tasks/DriverBirthdays_weekly")
wd<-getwd()

filename <- paste(as.Date(Sys.time()- 259200), "_", "DriverBirthdays.xlsx", sep = '')
filename<- paste("spreadsheets/",filename,sep="")
# load package for sql
library(DBI)
library(RODBC)
library(dplyr)
# connect to database
odbcChannel <- odbcConnect('echo_core', uid='Daria Alekseeva', pwd='Welcome01')
#odbcChannel <- odbcConnect('Dr SQL', uid='Daria Alekseeva', pwd='Welcome01')

birthdays <- sqlQuery( odbcChannel,"
                  
             select c.name,i.fullName,i.date_birth, 
		DATEPART(MM,I.date_birth) as 'month',
		DATEPART(DD,I.date_birth) as 'day',
		d.quick_base_id,d.employee_id as 'driver_id',dc.companyName,d.excludeFromAutoAllocation
from echo_core_prod..drivers d
left join echo_core_prod..individuals i on i.id = d.employee_id
left join echo_core_prod..callsigns c on c.driver_id=d.employee_id
left join echo_core_prod..drivers_company dc on dc.id =d.drivers_company_id
where date_left is null"
)

DriverJobCount <-sqlQuery( odbcChannel,
     "



     
     select AllJobs.driver_id, AllJobs.name, AllJobs.quick_base_id, sum(totalCharge) 'totalCharge', sum(totalPrice) 'totalPrice', weekNo, count(Alljobs.id) 'jobCount',AllJobs.companyName
     from
     (select j.id, j.driver_id, c.name,d.quick_base_id, j.jobDate, j.jobStatus, j.mopStatus, j.totalPrice, j.totalCharge, datepart(ISO_WEEK,jobdate) 'weekNo',dc.companyName
     from echo_core_prod..jobs j
     left join echo_core_prod..callsigns c on c.driver_id=j.driver_id
     left join echo_core_prod..drivers d on d.employee_id=j.driver_id
     left join echo_core_prod..drivers_company dc on dc.id=d.drivers_company_id
     where j.jobDate>'2016-01-01' and
     datepart(ISO_WEEK,j.jobdate) = datepart(ISO_WEEK,getdate()) - 1
     --and c.name not like '%DD%'
     and j.jobStatus in (7,10)) AllJobs
     group by AllJobs.driver_id, name, quick_base_id,weekNo,companyName
                      
                      ")

odbcClose(odbcChannel)

library(lubridate)


birthdays$thisYear<-as.Date(with(birthdays,paste("2016",month,day,sep="-")),"%Y-%m-%d")

birthdays$thisYear<-as.POSIXct(birthdays$thisYear)


today<-as.Date(Sys.time())-1
today<-as.POSIXct(today)
lastweek<-as.POSIXct(today-8*60*60*24)

birthdayslastweek<-birthdays[birthdays$thisYear<=today ,]
                             
birthdayslastweek<-birthdayslastweek[birthdayslastweek$thisYear>lastweek ,]

birthdayslastweek<-birthdayslastweek[!is.na(birthdayslastweek$thisYear),]

ActiveDriverBirthday<-merge(DriverJobCount,birthdayslastweek,by="driver_id",all.x=FALSE)

ActiveDriverBirthday<-ActiveDriverBirthday[,c(1:3,7,8,10,11,17)]

#save file
write.xlsx2(ActiveDriverBirthday,filename,row.names = FALSE)

#send file
library(RDCOMClient)
OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)
outMail[["subject"]] = "Happy Birthday to these guys "
#outMail[["To"]] = "Daria.Alekseeva@greentomatocars.com"
#outMail[["To"]] = "antony.carolan@greentomatocars.com;Haider.Variava@greentomatocars.com;Tyrone.Hunte@greentomatocars.com;paul.middleton@greentomatocars.com;moses.adegoroye@greentomatocars.com;Andrew.middleton@greentomatocars.com;tim.stone@greentomatocars.com"
outMail[["To"]] = "antony.carolan@greentomatocars.com;Haider.Variava@greentomatocars.com;Tyrone.Hunte@greentomatocars.com;Sophie.Jacobsen@greentomatocars.com"
outMail[["body"]] = "Happy Birthday to these drivers who worked for us last week"
outMail[["Attachments"]]$Add(paste(wd,filename,sep="/"))
outMail$Send()
rm(list = c("OutApp","outMail"))




