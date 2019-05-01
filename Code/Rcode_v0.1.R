#####
# Calculate Incentive Payments
#   - This script calculates incentive payments for communte trip reduction
#   - Built in R Version 3.5.0 - a copy of the installer is stored with this script
#   - Requires ODBC connection to the IdentityManagment database currently on Server: HQSQLIDMgtP (Name Data source: IdentityManagment)
#   - All packages used are stored in the folder with this script and will not be updated
#   
#####


#####
# Setup
#####

#Set working directory
WorkingDirectory <- "G:/QuarterlyIncentivePayments"
setwd(WorkingDirectory)

#Libraries
library(tidyverse)
library(RODBC)
library(RDCOMClient)


#Variables that may change  
IncentivePaymentAmount <- .75 #per trip
EligibleModes <- c("Bike", "Walk", "Carpool")
MinEligiblePrecent <- .6 #Remove all trip with < this amount.
EtcEmailAddress <- "SmithPa@wsdot.wa.gov; wesleyi@wsdot.wa.gov; NguyenAm@wsdot.wa.gov; leducb@wsdot.wa.gov"

CountiesNotSujbectTo60PercentRule <- c("King")
ModesNotSujbectTo60PercentRule <- c("Bus", "Train", "Light Rail", "Streetcar")

# Technical
dsnName <- "IdentityManagment"

#####
# Read data
#####

# Read Rideshare Online Data
InputData <- read.csv("./Input/UsersTripLogDetails.csv", stringsAsFactors = FALSE)

InputData$eMail <- tolower(InputData$eMail)

InputData <- InputData %>% filter(Purpose == "Commute")

# Identity Managment Data
conn <- odbcConnect(dsnName)

#Download Email Data
IM_EmailDataSqlQry <- "SELECT [EmailAddressId]
      ,[PersonId]
      ,[EmailAddress]
      ,[PrimaryEmailAddressFlag]
  FROM [IdentityManagement].[dbo].[EmailAddress]
  where [EmailAddress] is not NULL"

IM_EmailData <- sqlQuery(conn, IM_EmailDataSqlQry)
IM_EmailData$EmailAddress <- tolower(IM_EmailData$EmailAddress)

#Download Person Data
IM_PersonSqlQry <- "SELECT [PersonId]
      ,[FirstName]
      ,[MiddleName]
      ,[LastName]
      ,[PersonnelId]
      ,[ActiveFlag]
	    ,[JobPositionNumber]
      ,[WorkScheduleId]
      ,[JobClassificationCode]
      ,[EmployeeOfficeCity]
      ,[EmployeeTerminationDate]
    FROM [IdentityManagement].[dbo].[Person]
    WHERE (EmployeeTerminationDate IS NULL OR EmployeeTerminationDate >= DATEADD(m, -3, GETDATE())) AND PersonnelId IS NOT NULL AND [JobPositionNumber] IS NOT NULL"

IM_PersonData <- sqlQuery(conn, IM_PersonSqlQry)

close(conn)

#####
# Identify non-WSDOT Email Addresses
#####

BadEmailList <- InputData %>% 
  select(eMail) %>% 
  unique() %>% 
  left_join(IM_EmailData, by = c("eMail" = "EmailAddress")) %>% 
  filter(is.na(EmailAddressId))

SaveName <- paste0("./Output/MissingEmail/EmailAddressesNeedingCorrections.csv")
#SaveName <- paste0("./Output/MissingEmail/EmailAddressesNeedingCorrections-", Sys.Date(), ".csv")
write.csv(BadEmailList, SaveName, row.names = FALSE)

#####
# Deal with ferry tips
#####
FerryTripDays <- InputData %>% 
  filter(Mode == "Passenger Ferry") %>% 
  select(eMail, TripDate) %>% 
  unique() 

EligibleTripsOnFerryDays <- FerryTripDays %>% 
  left_join(InputData, by = c("eMail" = "eMail", "TripDate" = "TripDate"))

# Remove ferry trip days from Input Data
InputData_noFerry <- InputData %>% anti_join(EligibleTripsOnFerryDays, by = c("eMail" = "eMail", "TripDate" = "TripDate"))

#####
# Remove King County trips Not Sujbect To 60 Percent Rule
#####
CityCounty <- read.csv("G:/QuarterlyIncentivePayments/Code/GeographicCodes/geographic_codes.csv", stringsAsFactors = FALSE)

KingCountyCities <- CityCounty %>% 
  filter(COUNTY_NAME %in% CountiesNotSujbectTo60PercentRule) %>% 
  select(PLACE_NAME) %>% 
  unlist()

KingCountyPeopleEmailAddresses <- IM_PersonData %>% 
  filter(EmployeeOfficeCity %in% KingCountyCities) %>% 
  inner_join(IM_EmailData, by = "PersonId") %>% 
  select(EmailAddress) %>% 
  unlist()

unique(InputData$Mode)
test <- InputData %>% filter(eMail %in% KingCountyPeopleEmailAddresses & Mode %in% ModesNotSujbectTo60PercentRule)

#####
# Apply Min. Eligible trip mode precent rule
#####

# Determine  if at least X% of the trip distance is in an eligiable mode 
EligibleModeMilesPerDay <- InputData_noFerry %>% 
  filter(Mode %in% EligibleModes) %>% 
  group_by(eMail, TripDate) %>% 
  summarise(EligibleDistance = sum(CommuteDistance))

NotEligibleModeMilesPerDay <- InputData_noFerry %>% 
  filter(!(Mode %in% EligibleModes)) %>% 
  group_by(eMail, TripDate) %>% 
  summarise(NotEligibleDistance = sum(CommuteDistance))

MilesPerDayByEligiblity <- EligibleModeMilesPerDay %>% full_join(NotEligibleModeMilesPerDay, 
                                                                 by = c("eMail" = "eMail", "TripDate" = "TripDate"))

MilesPerDayByEligiblity$NotEligibleDistance[is.na(MilesPerDayByEligiblity$NotEligibleDistance)] <- 0
MilesPerDayByEligiblity$EligibleDistance[is.na(MilesPerDayByEligiblity$EligibleDistance)]       <- 0

MilesPerDayByEligiblity$PercentEligible <- 
  MilesPerDayByEligiblity$EligibleDistance / 
  (MilesPerDayByEligiblity$NotEligibleDistance + MilesPerDayByEligiblity$EligibleDistance)

NotEligibleDays <- MilesPerDayByEligiblity %>% 
  filter(PercentEligible < MinEligiblePrecent)

# Remove trips on days that are not Eligible Days for all users

EligibleTrips <- InputData_noFerry %>% anti_join(NotEligibleDays, by = c("eMail" = "eMail", "TripDate" = "TripDate"))

EligibleTrips <- rbind(EligibleTrips, EligibleTripsOnFerryDays)

#####
# Calculate eligible trips per day and daily payment amount
#####

CountOfEligibleTrips <- EligibleTrips %>% 
  filter(Mode %in% EligibleModes) %>% 
  group_by(eMail, TripDate) %>% 
  summarise(NumberEligibleTrips = n())

CountOfEligibleTrips$NumberEligibleTrips[CountOfEligibleTrips$NumberEligibleTrips > 2] <- 2

FinalPaymentAmount <- CountOfEligibleTrips %>% 
  group_by(eMail) %>% 
  summarise(NumberEligibleTrips = sum(NumberEligibleTrips)) %>% 
  mutate(IncentiveAmount = NumberEligibleTrips * IncentivePaymentAmount)


#####
# Merge with HR data and format for payroll
#####

IM_Data <- IM_PersonData %>% inner_join(IM_EmailData, by = "PersonId")

AllData <- FinalPaymentAmount %>% 
  left_join(IM_Data, by = c("eMail" = "EmailAddress")) 

OutPut <- data.frame("Email.Address" = AllData$eMail,
                      'Full name (Last, First)' = paste(AllData$LastName, AllData$FirstName, sep = ", "),
                     'x' = "",
                     'Employee Number' = AllData$PersonnelId,
                     'WT' = 1145,
                     'Sum of Total Incentive Due for Quarter' = AllData$IncentiveAmount,
                     'Payment Date' = "",
                     'Wage Type' = "Commute Incentive")

  
RecordCountCheck <- AllData %>% group_by(eMail) %>% summarise(Count = n())

SaveName <- paste0("./Output/Incentives/OutPut.csv")
#SaveName <- paste0("./Output/Data/OutPut-", Sys.Date(), ".csv")
write.csv(OutPut, SaveName, row.names = FALSE)

  
visData <- OutPut %>%  select(-Email.Address)
visData$Quarter <- "This Quarter" 

files <- data.frame(Path = list.files("./Archive/Output", full.names = TRUE), stringsAsFactors = FALSE)
files$Date <- unlist(lapply(files$Path, function(x) file.info(x)$ctime))
files %>% arrange(desc(Date)) %>% top_n(4)
  
for(i in 1:nrow(files)){
  data <- read.csv(files$Path[i])
  data$Quarter <- paste0(i , " Quarter(s) Ago")
  colnames(data) <- colnames(visData)
  visData <- rbind(visData, data)
}


userVisData <- visData %>% group_by(Quarter) %>% summarise(PeopleRecivingIncentive = n())

UserCountPlot <- ggplot(visData, aes(x = Quarter, fill = Quarter)) +
  geom_bar() +
  ggtitle("Number of People Recieving Incentive Payments")
#UserCountPlot
ggsave("./Output/Plots/UserCountPlot.png",
       plot = UserCountPlot,
       device = "png",
       width = 6,
       height = 4,
       units = "in")
  
IncentivePaymentAmount <- ggplot(visData, aes(x = Quarter, y = Sum.of.Total.Incentive.Due.for.Quarter, fill = Quarter)) +
  geom_boxplot() +
  ylab("Incentive Payment Ammount") +
  ggtitle("Summary of Incentive Payments")
#IncentivePaymentAmount 
ggsave("./Output/Plots/IncentivePaymentAmount.png",
       plot = IncentivePaymentAmount,
       device = "png",
       width = 6,
       height = 4,
       units = "in")


MyHTML <- paste0("<html><h2>Your report has finished running!</h2>
                  <p> There were <b>", nrow(BadEmailList), " email addresses that could not be matched </b> with internal data. 
                  <a href='", "file://hqolymfl01/groupi$/631020/QuarterlyIncentivePayments", "/Output/MissingEmail/EmailAddressesNeedingCorrections.csv"  ,"'> Please view them here</a>
                  <br>
                  <p> The incentive report is ready for review and formatted for payroll. <a href='", "file://hqolymfl01/groupi$/631020/QuarterlyIncentivePayments", "/Output/Incentives/OutPut.csv"  ,"'> Please view it here</a>  
                  </p>
                  <h2> Summary of quarterly incentive payments</h2>",
                  "<img src='", "file://hqolymfl01/groupi$/631020/QuarterlyIncentivePayments", "/Output/Plots/UserCountPlot.png' height='400' width='600'> <br>
                  <img src='", "file://hqolymfl01/groupi$/631020/QuarterlyIncentivePayments", "/Output/Plots/IncentivePaymentAmount.png' height='400' width='600'>")


OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)
outMail[["To"]] = EtcEmailAddress
outMail[["subject"]] = "Your Incentive Payment Report"
outMail[["HTMLbody"]] =  MyHTML                  
outMail$Send()  

