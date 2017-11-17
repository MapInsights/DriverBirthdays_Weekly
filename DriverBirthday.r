# WEEKLY PERFORMANCE REPORT. CLEANED.


# set the beginning of last week
start <- as.Date(Sys.time())
start
library(dplyr)
library(xlsx)
library(lubridate)

setwd("C:/Programs/gtc_tasks/DriverBirthdays_weekly")

unlink("spreadsheets", recursive = TRUE, force = FALSE)

# Create a directory for spreadsheets otherwise R having a heart attack 
dir.create("spreadsheets",showWarnings = F)

wd<-getwd()

filename <- paste(as.Date(Sys.time()), "_", "DriverBirthdays.xlsx", sep = '')
filename<- paste("spreadsheets/",filename,sep="")
# load package for sql
library(DBI)
library(RODBC)
library(dplyr)
# connect to database
odbcChannel <- odbcConnect('echo_core', uid='Daria Alekseeva', pwd='Welcome01')
#odbcChannel <- odbcConnect('Dr SQL', uid='Daria Alekseeva', pwd='Welcome01')

birthdays <- sqlQuery( odbcChannel,"
                       
                      select DISTINCT d.employee_id as 'driver_id',c.name,i.fullName,i.date_birth, 
		                  DATEPART(MM,I.date_birth) as 'month',
                       DATEPART(DD,I.date_birth) as 'day',
                       d.quick_base_id,dc.companyName,d.excludeFromAutoAllocation,ie.email,ip.phone_number
                       from echo_core_prod..drivers d
                       left join echo_core_prod..individuals i on i.id = d.employee_id
                       left join echo_core_prod..individual_emails ie on i.id = ie.individual_id
                       left join echo_core_prod..individual_phones ip on i.id = ip.individual_id
                       left join echo_core_prod..callsigns c on c.driver_id=d.employee_id
                       left join echo_core_prod..drivers_company dc on dc.id =d.drivers_company_id
                       where date_left is null and c.name is not NULL and defaultPhone =1 and defaultEmail =1"
)

DriverJobCount <-sqlQuery( odbcChannel,
                           "
                           
                           
                           
                           
                           select AllJobs.driver_id, AllJobs.name, AllJobs.quick_base_id, sum(totalCharge) 'totalCharge', sum(totalPrice) 'totalPrice', weekNo, count(Alljobs.id) 'jobCount',AllJobs.companyName
                           from
                           (select j.id, j.driver_id, c.name,d.quick_base_id, j.jobDate, j.jobStatus, j.mopStatus, j.totalPrice, j.totalCharge, datepart(ISO_WEEK,jobdate) +2 'weekNo',dc.companyName
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


birthdays$thisYear<-as.Date(with(birthdays,paste("2017",month,day,sep="-")),"%Y-%m-%d")

birthdays$thisYear<-as.POSIXct(birthdays$thisYear)


today<-as.Date(Sys.time())
today<-as.POSIXct(today)
lastweek<-as.POSIXct(today-7*60*60*24)
nextweek<-as.POSIXct(today+7*60*60*24)

birthdayslastweek<-birthdays[birthdays$thisYear<=today ,]

birthdayslastweek<-birthdayslastweek[birthdayslastweek$thisYear>lastweek ,]

birthdayslastweek<-birthdayslastweek[!is.na(birthdayslastweek$thisYear),]

ActiveDriverBirthday<-merge(DriverJobCount,birthdayslastweek,by="driver_id",all.x=FALSE)

ActiveDriverBirthday1<-ActiveDriverBirthday[,c(1:3,8,10,11,17,18,19)]#Removing col 7 as Job count not needed-Tyrone
ActiveDriverBirthday1<-ActiveDriverBirthday1[!duplicated(ActiveDriverBirthday1$email),]

#Next Week Logic

birthdaysnextweek<-birthdays[birthdays$thisYear<=nextweek ,]

birthdaysnextweek<-birthdaysnextweek[birthdaysnextweek$thisYear>=today ,]

birthdaysnextweek<-birthdaysnextweek[!is.na(birthdaysnextweek$thisYear),]

ActiveDriverBirthday_nw<-merge(DriverJobCount,birthdaysnextweek,by="driver_id",all.x=FALSE)

ActiveDriverBirthday_nw1<-ActiveDriverBirthday_nw[,c(1:3,8,10,11,17,18,19)]#Removing col 7 as Job count not needed-Tyrone
ActiveDriverBirthday_nw1<-ActiveDriverBirthday_nw1[!duplicated(ActiveDriverBirthday_nw1$email),]

#save file
write.xlsx(ActiveDriverBirthday1,filename,sheetName="LastWeek",row.names = FALSE)
write.xlsx(ActiveDriverBirthday_nw1,filename,sheetName="NextWeek",row.names = FALSE, append =TRUE)

#send file
library(RDCOMClient)
OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)
outMail[["subject"]] = "Happy Birthday to these guys "
#outMail[["To"]] = "muaaz.sarfaraz@greentomatocars.com"
#outMail[["To"]] = "antony.carolan@greentomatocars.com;Haider.Variava@greentomatocars.com;Tyrone.Hunte@greentomatocars.com;paul.middleton@greentomatocars.com;moses.adegoroye@greentomatocars.com;Andrew.middleton@greentomatocars.com;tim.stone@greentomatocars.com"
outMail[["To"]] = "jonny.goldstone@greentomatocars.com;olivesupport@greentomatocars.com;limesupport@greentomatocars.com;mintsupport@greentomatocars.com;Daria.Alekseeva@greentomatocars.com;antony.carolan@greentomatocars.com;Haider.Variava@greentomatocars.com;Tyrone.Hunte@greentomatocars.com;Sophie.Jacobsen@greentomatocars.com;Yinka.Ogunniyi@greentomatocars.com;muaaz.sarfaraz@greentomatocars.com"
outMail[["body"]] = "Happy Birthday to these drivers"
outMail[["Attachments"]]$Add(paste(wd,filename,sep="/"))
outMail$Send()
rm(list = c("OutApp","outMail"))




