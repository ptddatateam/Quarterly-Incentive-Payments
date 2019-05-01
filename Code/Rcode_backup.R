#####
# Calculate Incentive Payments
#   - This script calculates incentive payments for communte trip reduction
#   - Built in R Version 3.4.4 - a copy of the windows installer is stored with this script
#   - Requires ODBC connection to the IdentityManagment database currently on Server: HQSQLIDMgtP (Name Data source: IdentityManagment)
#   - All packages used in this process are stored in the folder with this script and will not be updated
#   - Ian Wesley 6/21/18
#
#   - Other Items:
#       - This script is called from a scheduled job that runs every 5 minuites on my computer.
#       - The scheduled job first checks to see if a text file says "Run".  If it does then this script is run.  
#       - There is a vbs file that a person can click on to change to value of the file to "Run". 
#       - This enables the end user to triger the report which is then run on my computer
#####

##### 
# variables that may change 
#####

# Non-technical 
EtcEmailAddress        <- "wesleyi@wsdot.wa.gov; SmithPa@wsdot.wa.gov;" #NguyenAm@wsdot.wa.gov; leducb@wsdot.wa.gov"
MinEligiblePrecent     <- .6 # Remove all trip with less this amount in an eligible mode.
EligibleModes          <- c("Bike", "Walk", "Carpool")
IncentivePaymentAmount <- .75 # per trip
MaxPaymentPerMonth     <- 33  
MaxTripsPerDay         <- 2
TechnicalMainancePersonName <- "Ian" 

# The following three variables are historical and can be used to exempt modes from the 60% rule.
ModesNotSubjectTo60PercentRule <- c() # any mode added here will simple be removed from the data set and therefore the 60% rule will not apply to this mode 

CountiesNotSujbectTo60PercentRule <- c() # People with a worksite in this county will have the following modes exmpted from the 60% rule
ModesInCountiesNotSujbectTo60PercentRule <- c() 

# Technical
dsnName <- "IdentityManagment"
WorkingDirectory <- "//hqolymfl01/groupi$/631020/QuarterlyIncentivePayments"

#####
# Setup
#####

# Email functions
GenerateEmailError <- function(MyErrMsg, err, myTechnicalMainancePersonName = TechnicalMainancePersonName){
  myHTML <- paste0("<html><h2>There was an error with your incentive report</h2>",
                   "<p> There is no need to do anything. ", myTechnicalMainancePersonName, " has been notified of the error and will work to correct it.",
                   "<br><br> Error Message: <br>", MyErrMsg, "<br>", err)
  myHTML
}

SendEmailReport <- function(myHTML, isErr = FALSE, myEtcEmailAddress = EtcEmailAddress) {
  subject <- "Your Incentive Payment Report"
  
  if (isErr) {
    subject <- "Error with incentive payment report"
  }
  
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  outMail[["To"]] = myEtcEmailAddress
  outMail[["subject"]] = subject
  outMail[["HTMLbody"]] =  myHTML                  
  outMail$Send()  
}


#Set working directory
tryCatch(setwd(WorkingDirectory), error= function(cond) {
  errMsg <- "Could not find working directory."
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
})

#Load libraries
LibLoc <- paste0(WorkingDirectory, "/Code/Packages")

library(tidyverse, lib.loc = LibLoc)
library(RODBC, lib.loc = LibLoc)
library(RDCOMClient, lib.loc = LibLoc)

#####
# Read data
#####

# Read Rideshare Online Data
GetInputData <- function() {
  output <- tryCatch({
    
    lfiles <- list.files(paste0(WorkingDirectory,"/Input/"), full.names = TRUE)
    lfiles <- lfiles[grepl(".csv", lfiles)]
    
    if (length(lfiles) > 1){
      errMsg =  "There is more than one csv file in the input folder.  Please be sure there is only one csv file in the input folder and run agian."
      SendEmailReport(GenerateEmailError(errMsg, "File input error"), isErr = TRUE)
      stop()
    }
    
    con <- file(lfiles[1])
    
    InputDataHeader <- readLines(con, n = -1)
    
    if(InputDataHeader[1] != "TripNumber,UserName,ScreenName,eMail,TripDate,Purpose,Mode,IsVerified,CommuteDistance,OriginCity,Worksite,DestinationCity,GasSaved,CarbonDioxideSaved"){
      errMsg =  "Input data is not formated correctly"
      SendEmailReport(GenerateEmailError(errMsg, "File input error"), isErr = TRUE)
      stop()
    }
    
    InputDataHeader[1] <- "TripNumber,UserName,ScreenName,eMail,TripDate,Purpose,Mode,IsVerified,CommuteDistance,OriginCity,Worksite,DestinationCity,GasSaved,CarbonDioxideSaved,"
    tempCon <- tempfile(pattern = "file", tmpdir = tempdir(), fileext = ".csv")
    writeLines(InputDataHeader, tempCon)
    
    InputData <- read.csv(tempCon, stringsAsFactors = FALSE)
    InputData$X <- NULL
    
    # Validate input file is formated correctly
    if(paste(colnames(InputData),collapse=",") != "TripNumber,UserName,ScreenName,eMail,TripDate,Purpose,Mode,IsVerified,CommuteDistance,OriginCity,Worksite,DestinationCity,GasSaved,CarbonDioxideSaved"){
      errMsg =  "Input data is not formated correctly"
      SendEmailReport(GenerateEmailError(errMsg, "File input error"), isErr = TRUE)
      stop()
    }
    
    InputData$eMail <- tolower(InputData$eMail)
    InputData       <- InputData %>% filter(Purpose == "Commute")
    
    close(con)
    return(InputData)
    
  }, error = function(err) {
    errMsg <- "Error loading data from rideshare online."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

# Download Email Data from Identity Managment 
GetIM_EmailData <- function(dsnName){
  output <- tryCatch({
    
    conn <- odbcConnect(dsnName)
    # SQL Querry
    IM_EmailDataSqlQry <- "SELECT [EmailAddressId]
      ,[PersonId]
      ,[EmailAddress]
      ,[PrimaryEmailAddressFlag]
      FROM [IdentityManagement].[dbo].[EmailAddress]
      where [EmailAddress] is not NULL"
    IM_EmailData <- sqlQuery(conn, IM_EmailDataSqlQry)
    
    IM_EmailData$EmailAddress <- tolower(IM_EmailData$EmailAddress)
    close(conn)
    
    return(IM_EmailData)
    
  }, error = function(err) {
    errMsg <- "Error Connecting to Identity Managment - Email Table"
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

# Download person data from Identity Managment 
GetIM_PersonData <- function(dsnName) {
  output <- tryCatch({
    
    conn <- odbcConnect(dsnName)
        # SQL Query
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
    
    return(IM_PersonData)
    
  }, error = function(err) {
    errMsg <- "Error Connecting to Identity Managment - Person Table"
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

#####
# Identify non-WSDOT Email Addresses
#####

GetNonWsdotEmailAddresses <- function(InputDat, myWorkingDirectory = WorkingDirectory, myIM_EmailData = IM_EmailData) {
  output <- tryCatch({
    
    InvalidEmailList <- InputData %>% 
      select(eMail) %>% 
      unique() %>% 
      left_join(myIM_EmailData, by = c("eMail" = "EmailAddress")) %>% 
      filter(is.na(EmailAddressId))
    
    SaveName <- paste0(myWorkingDirectory, "/Output/MissingEmail/EmailAddressesNeedingCorrections.csv")
    write.csv(InvalidEmailList, SaveName, row.names = FALSE)
   
    return(InvalidEmailList)
    
  }, error = function(err) {
    errMsg <- "Error generating list of non-WSDOT email address"
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

#####
# Remove special circumstance trips
#####

GetCitiesNotSubjectTo60PercentRule <- function(myCountiesNotSujbectTo60PercentRule = CountiesNotSujbectTo60PercentRule){
  output <- tryCatch({
    
    CityCounty <- read.csv("G:/QuarterlyIncentivePayments/Code/GeographicCodes/geographic_codes.csv", stringsAsFactors = FALSE)
    Cities <- CityCounty %>% 
      filter(COUNTY_NAME %in% myCountiesNotSujbectTo60PercentRule) %>% 
      select(PLACE_NAME) %>% 
      unlist()
    
    return(Cities)
    
  }, error = function(err) {
    errMsg <- "Error generating list of cities not subject to 60 percent Rule"
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

# Returns vector of email address of people that have a work site in a county not subject to the 60% rule
GetPersonsInCountiesNotSujbectTo60PercentRule <- function(myIM_PersonData = IM_PersonData, 
                                                          myIM_EmailData = IM_EmailData,
                                                          Cities = GetCitiesNotSubjectTo60PercentRule()){
  
  
  # Vector of email address
  PersonsInCountiesNotSujbectTo60PercentRule <- myIM_PersonData %>% 
    filter(EmployeeOfficeCity %in% Cities) %>% 
    inner_join(myIM_EmailData, by = "PersonId") %>% 
    select(EmailAddress) %>% 
    unlist()
  
  PersonsInCountiesNotSujbectTo60PercentRule
} 


RemoveTripsNotSubjectTo60PercentRule <- function(InputData, 
                                                 myModesNotSubjectTo60PercentRule = ModesNotSubjectTo60PercentRule,
                                                 PersonsInCountiesNotSujbectTo60PercentRule = GetPersonsInCountiesNotSujbectTo60PercentRule(),
                                                 myModesInCountiesNotSujbectTo60PercentRule = ModesInCountiesNotSujbectTo60PercentRule) {
  output <- tryCatch({
    
    NoTripsNotSubjectTo60PercentRule <- InputData %>% 
      filter(!(Mode %in% myModesNotSubjectTo60PercentRule)) %>% 
      filter(!(eMail %in% PersonsInCountiesNotSujbectTo60PercentRule & Mode %in% myModesInCountiesNotSujbectTo60PercentRule))
    
    return(NoTripsNotSubjectTo60PercentRule)
    
  }, error = function(err) {
    errMsg <- "Error removing trips not subject to 60 percent rule or Getting persons in counties not subject to the 60 percent rule."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

#####
# Remove all trips on days where a person did not meet the MinEligiblePrecent threshold
#####

RemoveTripsOnDaysNotMeetingThreshold <- function(InputData, myEligibleModes = EligibleModes) {
  output <- tryCatch({
    # Calculate elegible trip distance
    EligibleModeMilesPerDay <- InputData %>% 
      filter(Mode %in% EligibleModes) %>% 
      group_by(eMail, TripDate) %>% 
      summarise(EligibleDistance = sum(CommuteDistance))
    
    # Calculate Not eligible trips
    NotEligibleModeMilesPerDay <- InputData %>% 
      filter(!(Mode %in% EligibleModes)) %>% 
      group_by(eMail, TripDate) %>% 
      summarise(NotEligibleDistance = sum(CommuteDistance))
    
    # Join together
    MilesPerDayByEligiblity <- EligibleModeMilesPerDay %>% 
      full_join(NotEligibleModeMilesPerDay, by = c("eMail" = "eMail", "TripDate" = "TripDate"))
    
    MilesPerDayByEligiblity[is.na(MilesPerDayByEligiblity)] <- 0
    
    # Calculate percent
    MilesPerDayByEligiblity$PercentEligible <- 
      MilesPerDayByEligiblity$EligibleDistance / 
      (MilesPerDayByEligiblity$NotEligibleDistance + MilesPerDayByEligiblity$EligibleDistance)
    
    # Remove Trips
    NotEligibleDays <- MilesPerDayByEligiblity %>% 
      filter(PercentEligible < MinEligiblePrecent)
    EligibleTrips <- InputData %>% 
      anti_join(NotEligibleDays, by = c("eMail" = "eMail", "TripDate" = "TripDate"))
    
    return(EligibleTrips)
  }, error = function(err) {
    errMsg <- "Error removing trips that did not meet 60 percent threshold."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

#####
# Calculate payment amount
#####

CalculatePaymentAmount <- function(InputData, 
                                   myEligibleModes = EligibleModes, 
                                   myIncentivePaymentAmount = IncentivePaymentAmount, 
                                   myMaxTripsPerDay = MaxTripsPerDay, 
                                   myMaxPaymentPerMonth = MaxPaymentPerMonth) {
  output <- tryCatch({
    
    # Count trips per day and set max trips per day
    CountOfEligibleTrips <- InputData %>% 
      filter(Mode %in% EligibleModes) %>% 
      group_by(eMail, TripDate) %>% 
      summarise(NumberEligibleTrips = n())
    
    CountOfEligibleTrips$NumberEligibleTrips[CountOfEligibleTrips$NumberEligibleTrips > MaxTripsPerDay] <- MaxTripsPerDay
    
    # Calculate incentive payment per month and cap at max per month
    CountOfEligibleTrips$DailyPaymentAmount <- CountOfEligibleTrips$NumberEligibleTrips * IncentivePaymentAmount
    
    PaymentPerMonth <- CountOfEligibleTrips %>% 
      mutate(PaymentMonth = months(as.Date(TripDate))) %>% 
      group_by(eMail, PaymentMonth) %>% 
      summarise(MonthlyPayment = sum(DailyPaymentAmount))
    
    PaymentPerMonth$MonthlyPayment[PaymentPerMonth$MonthlyPayment > MaxPaymentPerMonth] <- MaxPaymentPerMonth
    
    # Calculate total payment
    TotalPayment <- PaymentPerMonth %>% 
      group_by(eMail) %>% 
      summarise(IncentiveAmount = sum(MonthlyPayment))
    
    return(TotalPayment)
  }, error = function(err) {
    errMsg <- "Error Calculating payment amount."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

#####
# Merge with HR data and format for payroll
#####

MergeWithIdentityManagementData <- function(InputData, myIM_PersonData = IM_PersonData, myIM_EmailData = IM_EmailData) {
  output <- tryCatch({
    
    IM_Data <- myIM_PersonData %>% inner_join(myIM_EmailData, by = "PersonId")
    
    AllData <- InputData %>% 
      left_join(IM_Data, by = c("eMail" = "EmailAddress")) 
    
    return(AllData)
  }, error = function(err) {
    errMsg <- "Error Merging with Identity Management data."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

FormatForPayroll <- function(InputData) {
  output <- tryCatch({
    
    outPut <- data.frame("Email.Address" = InputData$eMail,
                         'Full name (Last, First)' = paste(InputData$LastName, InputData$FirstName, sep = ", "),
                         'x' = "",
                         'Employee Number' = InputData$PersonnelId,
                         'WT' = 1145,
                         'Sum of Total Incentive Due for Quarter' = InputData$IncentiveAmount,
                         'Payment Date' = "",
                         'Wage Type' = "Commute Incentive")
    return(outPut)
  }, error = function(err) {
    errMsg <- "Error formating for payroll."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

SaveOutputAsCsv <- function(InputData, myWorkingDirectory = WorkingDirectory){
  output <- tryCatch({
    
    SaveName <- paste0(myWorkingDirectory, "/Output/Incentives/Output.csv")
    write.csv(InputData, SaveName, row.names = FALSE)
  }, error = function(err) {
    errMsg <- "Error saving output to csv."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}


#####
# Generate Email Report
#####

GetPastFourQuartersData <- function(Input, myWorkingDirectory = WorkingDirectory) {
  output <- tryCatch({
    
    visData <- Input %>% select(-Email.Address) # First historical record does not have email address
    visData$Quarter <- "This Quarter" 
    
    archiveLoc <- paste0(myWorkingDirectory, "/Archive/Output")
    
    files <- data.frame(Path = list.files(archiveLoc, full.names = TRUE), stringsAsFactors = FALSE)
    files$Date <- unlist(lapply(files$Path, function(x) file.info(x)$ctime))
    files %>% arrange(desc(Date)) %>% top_n(4)
    
    for(i in 1:nrow(files)){
      data <- read.csv(files$Path[i]) 
      
      if("Email.Address" %in% colnames(data)){
        data <- data %>% select(-Email.Address) # First historical record does not have email address
      }
      data$Quarter <- paste0(i , " Quarter(s) Ago")
      
      colnames(data) <- colnames(visData)
      visData <- rbind(visData, data)
    }
    
    return(visData)
  }, error = function(err) {
    errMsg <- "Error getting historical data for email report."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
  return(output)
}

GeneratePlotsForEmailReport <- function(visData, myWorkingDirectory = WorkingDirectory){
  tryCatch({
  # UserCountPlot
  visData
  UserCountPlot <- ggplot(visData, aes(x = Quarter, fill = Quarter)) +
    geom_bar() +
    ggtitle("Number of People Recieving Incentive Payments") +
    ylab("Number of People") 
  ggsave(paste0(myWorkingDirectory, "/Output/Plots/UserCountPlot.png"),
         plot = UserCountPlot,
         device = "png",
         width = 6,
         height = 4,
         units = "in")
  
  # IncentivePaymentAmount 
  IncentivePaymentAmount <- ggplot(visData, aes(x = Quarter, y = Sum.of.Total.Incentive.Due.for.Quarter, fill = Quarter)) +
    geom_boxplot() +
    ylab("Incentive Payment Ammount") +
    ggtitle("Summary of Incentive Payments")
  ggsave(paste0(myWorkingDirectory, "/Output/Plots/IncentivePaymentAmount.png"),
         plot = IncentivePaymentAmount,
         device = "png",
         width = 6,
         height = 4,
         units = "in")
  }, error = function(err) {
    errMsg <- "Error creating plots for email report."
    myHTML <- GenerateEmailError(errMsg, err)
    SendEmailReport(myHTML, isErr = TRUE)
    stop()
  })
}

GenerateEmailReport <- function(InvalidEmailListCount = nrow(InvalidEmailList), myWorkingDirectory = WorkingDirectory) {
  MyHTML <- paste0("<html><h2>Your report has finished running!</h2>
                  <p> There were <b>", InvalidEmailListCount, " email addresses that could not be matched </b> with internal data. 
                   <a href='file://", myWorkingDirectory, "/Output/MissingEmail/EmailAddressesNeedingCorrections.csv"  ,"'> Please view them here</a>
                   <br>
                   <p> The incentive report is ready for review and formated for payroll. <a href='file://", myWorkingDirectory, "/Output/Incentives/Output.csv"  ,"'> Please view it here</a>  
                   </p>
                   <h2> Summary of quarterly incentive payments</h2>",
                   "<img src='file://", myWorkingDirectory, "/Output/Plots/UserCountPlot.png' height='400' width='600'> <br>
                   <img src='file://", myWorkingDirectory, "/Output/Plots/IncentivePaymentAmount.png' height='400' width='600'>")

  MyHTML
}

#####
# Run Process
#####

InputData     <- GetInputData()
IM_EmailData  <- GetIM_EmailData(dsnName)
IM_PersonData <- GetIM_PersonData(dsnName)

InvalidEmailList <- GetNonWsdotEmailAddresses(InputData)

Output <- InputData %>% 
  RemoveTripsNotSubjectTo60PercentRule() %>% 
  RemoveTripsOnDaysNotMeetingThreshold() %>% 
  CalculatePaymentAmount() %>% 
  MergeWithIdentityManagementData() %>% 
  FormatForPayroll()

SaveOutputAsCsv(Output)

visData <- GetPastFourQuartersData(Output)

GeneratePlotsForEmailReport(visData)

SendEmailReport(GenerateEmailReport())
