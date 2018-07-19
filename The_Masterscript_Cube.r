###PLEASE READ!!!
### Assignment given to Robert Picard - Compile all R files into one masterscript detailing the fiscal year data from 2017 to 2018 - July 19th, 2018
### My assignment is to make sure the code is all readable and succinct, and to add comments where I feel there are better ways to write the code
### The order of R scipts goes like this:
#1 - Masterscript_JM - This is the masterscript that the previous co-op had not finished, but was working on
#2 - Summary_by_Business_Unit_v3 - This is the R code that details the spending per month from April to June of 2018 per business unit
#3 - 50K_Bin - This seperates purchase orders between micro, small, and large
#4 - highlvl - shows the platforms through which POs are ordered on
#5 - SpendByPlatformSub1000 - Shows the count and spend by vendors per platform
#Note: My comments will be after the other comments in a "#Note:" format.



#1 - Masterscript_JM 


# PO side of Executive Overview data pull for Dan Maloney
# Author: Jansen Manahan (comprised of scripts for Scott Taing, Kojiro So, Jenna Goldberg, and Jansen Manahan)
# Created 6/26/18
# Last Updated; 6/26/18
# Does not include platform breakdowns, vendor info, or timeline of req to dispatch

# Set up --------------------------------------------------------------------------------------------------------------------------------------------
# Read libraries
library(janitor)
library(tidyverse)
library(readxl)
library(kableExtra)
library(knitr)
library(lubridate)

# Load data: FMIS data for POs (including DASH_E_ERROR...), FMIS data for Reqs, Business Center data (all) Note#: We should be pulling data from July 1 to June 30th
Raw_PO <- read_excel("O:/Codebase/0_FMIS_Source_Data/July1-2017_June15-2018_Pull/REQ_TO_PO_SPILT_PO_SIDE_PUBLIC_July1_June15.xlsx", skip = 1)

# Necessary Variables -------------------------------------------------------------------------------------------------------------------------------
# Most recent date in data
Today <- max(Raw_PO$`Last Dttm`) %>% as_date()

# Beginning of one month ago
Last_Month <- Today - months(1)

# Beginning of two months ago
Two_Months_Ago <- Today - months(2)

# Beginning of FY
FY_Day_1 <- if_else( # MANUAL override of this variable may be merited in July
  month(Today) >= 7,
  str_glue("{year(Today)}-07-01") %>% ymd(),
  str_glue("{year(Today - years(1))}-07-01") %>% ymd()
)

# Buyer classifications: SE, Inventory, Non-Inventory
Sourcing_Execs <- c("AFLYNN", "MBARRY", "TDIONNE", "MBERNSTEIN") # MANUAL
Not_Sourcing_Execs <- Raw_PO[!Raw_PO$Buyer %in% Sourcing_Execs,]$Buyer %>% as.tibble() %>% distinct()
Non_Inv_Buyers <- c("JKIDD", "CFRANCIS", "JLEBBOSSIERE", "DMALONEY", "KHALL", "TTOUSSAINT") # MANUAL
Inv_Buyers <- c("DMARTINOS", "TSULLIVAN1", "PHONG", "AKNOBEL") # MANUAL

Taxonomy_Duplicates <- # Some taxonomy is not MECE, so we need to clean the data for it
  Raw_PO %>% 
  distinct(.data$`Level 1`, .data$`Level 2`) %>% 
  count(.data$`Level 2`) %>%
  filter(n == 2) %>%
  select(`Level 2`)

# Clean all data ------------------------------------------------------------------------------------------------------------------------------------
Clean_PO <- 
  Raw_PO %>% 
  filter(.data$Status != c("O", "PA", "PX"), .data$`Sum Amount` >= 0) %>% 
  mutate(
    Is_This_Month = if_else(.data$`PO Date` > Last_Month, T, F),
    Is_Previous_Month = if_else(.data$`PO Date` > Two_Months_Ago, T, F),
    Explicit_Taxonomy_2 = if_else(
      .data$`Level 2` %in% Taxonomy_Duplicates$`Level 2`,
      str_glue("{.data$`Level 2`}_{.data$`Level 1`}") %>% as.character(), 
      .data$`Level 2`
    ),
    Buyer_Classification = case_when(
      .data$Buyer %in% Sourcing_Execs     ~ "Sourcing Executive",
      .data$Buyer %in% Non_Inv_Buyers     ~ "Non-Inventory Buyer",
      .data$Buyer %in% Inv_Buyers         ~ "Inventory Buyer",
      .data$Buyer %in% Not_Sourcing_Execs ~ "Buyer Outside P&L",
      TRUE                                ~ "ERROR" # Is this an error or are there legitimate reasons for a buyer to be NA? Legal?
    )#,
    ###   Bid_Category = ?
  )

Clean_PO <- 
  Clean_PO %>% 
  group_by(.data$Bussiness_Unit, .data$`PO No.`) %>% 
  summarize(PO_Total = sum(.data$`Sum Amount`)) %>% 
  right_join(Clean_PO, by = c("Bussiness_Unit", "PO No.")) # RIGHT JOIN is the pipable LEFT JOIN

Clean_PO <- 
  Clean_PO %>% 
  group_by(.data$Bussiness_Unit, .data$`PO No.`, .data$Buyer_Classification) %>% 
  summarize(PO_Total = n())

#Note: Some of JM's comments of his own masterscript
# Rename columns
# Filter out certain statuses, non-positive spend PO and req lines
# Dedup(?) -- better to mutate a boolean field (make sure there is always a L1)
# Mutate variables for data integrity: a key of BU-PO?, HasPO for reqs
# Handle NAs
# Mutate variables for filtering: Is within one month, Is within two months, Req is in Biz Center data (outer join?), Spend bin per BU-PO (grouping)_
# Is denied ever?
# Mutate for output summarization: Req origin, PO buyer classification (incl "SWC Buyer"), Bid Category (Single, Sole...), unique taxonomy_
# Denial age bin, Req age bin
# Mutate req, solicitation, and PO spans


# Calculate output ----------------------------------------------------------------------------------------------------------------------------------
# META: group by ones that need hdr and ones that need lines
# META: Write to different sheets of one XL file
# POs by spend bin and time period
# POs by buyer (rolled up to buyer classification) and time period
# POs by bid category
# POs by taxonomy and (origin code | BU | buyer)
##### POs by platform, time period, and spend bin (>|<) $1000
##### Denials by (age bin | size | level | user)
##### Procurement timeline by buyer classification and timeline
##### Req backlog by month (3 trailing) and current output
##### Reqs by Age, BU, Buyer, and Is on hold (and roll-ups)
##### Vendors by (platform | direct voucher vs PO) ranked by (Count of POs | Sum of MBTA spend to them)

# Visualization -------------------------------------------------------------------------------------------------------------------------------------

# Output to folders ---------------------------------------------------------------------------------------------------------------------------------




#2 - Summary_by_Business_Unit_v3

#Note: When we finalize the master script, we don't need to load the library multiple times
library(readxl)
library(tidyverse)
library(knitr)
library(kableExtra)
library(lubridate)
library(stringr)
library(janitor)

## Data is from July1_2017_June15_2018 #Note: We should be going from July 1st to June 30th
setwd("O:/Codebase/0_FMIS_Source_Data/July1-2017_June15-2018_Pull")
raw_bizunit <- read_excel("REQ_TO_PO_SPILT_PO_SIDE_PUBLIC_July1_June15.xlsx", skip = 1)
setwd("O:/Codebase/1_All_PO_count_and_spend_bizunit")

##Handle renames at the FMIS level in the future
raw_bizunit <- raw_bizunit %>% rename("PO_No" = "PO No.", "PO_Date" = "PO Date", "Sum_Amount" = "Sum Amount", "Req_ID" = "Req ID", "Bid_Category" = "Bid Category")

##Clean
clean_raw <- raw_bizunit %>% filter(.$Status != c("O","PA", "PX"))

## Past 2 Months #Note: This seems misleading because the objects and the comment says the past two months, but the dttms are only an interval of one month.
twomonth_clean <- clean_raw %>% filter(.data$PO_Date >= ymd("2018-04-01") & .data$PO_Date < ymd("2018-05-01"))
twomonth_amt <- twomonth_clean %>% group_by(.data$Bussiness_Unit) %>% summarize(twomonth_spend = sum(.data$Sum_Amount))
twomonth_cnt <- twomonth_clean %>% distinct(.data$PO_No, .keep_all = T) %>% group_by(.data$Bussiness_Unit) %>% summarize(twomonth_cnt = n())
twomonth <- twomonth_amt %>% left_join(twomonth_cnt)

## Past 1 Month
onemonth_clean <- clean_raw %>% filter(.data$PO_Date >= ymd("2018-05-01") & .data$PO_Date < ymd("2018-06-01"))
onemonth_amt <- onemonth_clean %>% group_by(.data$Bussiness_Unit) %>% summarize(onemonth_spend = sum(.data$Sum_Amount))
onemonth_cnt <- onemonth_clean %>% distinct(.data$PO_No, .keep_all = T) %>% group_by(.data$Bussiness_Unit) %>% summarize(onemonth_cnt = n())
onemonth <- onemonth_amt %>% left_join(onemonth_cnt)

##YTD #Note: I think this is suppose to be short for yearly total spend.
YTD_amt <- clean_raw %>% group_by(.data$Bussiness_Unit) %>% summarize(ytd_spend = sum(.data$Sum_Amount))
YTD_cnt <- clean_raw %>% distinct(.data$PO_No, .keep_all = T) %>% group_by(.data$Bussiness_Unit) %>% summarize(ytd_cnt = n())
YTD <- YTD_amt %>% left_join(YTD_cnt)

##Final #Note:Summaries by business unit over the past two months.
summary_biz_unit <- YTD %>% left_join(twomonth)
summary_biz_unit <- summary_biz_unit %>% left_join(onemonth) %>% replace(is.na(.),0) %>%  arrange(desc(.data$ytd_spend)) %>% adorn_totals()

##Renaming and formatting
summary_biz_unit <- summary_biz_unit %>% rename("Total Spend" = "ytd_spend",
                                                "Total No. of POs" = "ytd_cnt",
                                                "Spend in April" = "twomonth_spend",
                                                "No. of POs in April" = "twomonth_cnt",
                                                "Spend in May" = "onemonth_spend",
                                                "No. of POs in May" = "onemonth_cnt")

summary_biz_unit$`Total Spend` <- paste("$",format(round(summary_biz_unit$`Total Spend`,0), big.mark=","),sep="")
summary_biz_unit$`Spend in April` <- paste("$",format(round(summary_biz_unit$`Spend in April`,0), big.mark=","),sep="")
summary_biz_unit$`Spend in May` <- paste("$",format(round(summary_biz_unit$`Spend in May`,0), big.mark=","),sep="")
summary_biz_unit$`Total No. of POs` <- format(round(summary_biz_unit$`Total No. of POs`,0), big.mark=",")
summary_biz_unit$`No. of POs in April` <- format(round(summary_biz_unit$`No. of POs in April`,0), big.mark=",")
summary_biz_unit$`No. of POs in May` <- format(round(summary_biz_unit$`No. of POs in May`,0), big.mark=",")

write_csv(summary_biz_unit, "Summary_by_Business_Unit.csv")




#3 - 50K_Bin


# Last updated 6/28/18 by Jansen Manahan to adjust threshold of "SMALL"

library(readxl,lubridate)
library(tidyverse)

#Set Enviroment and Read in Raw excel #Note: This needs to be updated to include the right date and time
#Data is from June1_2017-May30_2018--YTD
setwd("O:/Codebase/0_FMIS_Source_Data/July1-2017_June15-2018_Pull")
po_raw <- read_excel("REQ_TO_PO_SPILT_PO_SIDE_PUBLIC_July1_June15.xlsx", skip = 1)
setwd("O:/Codebase/5_YTD_Count_and Spend_of_POs_sent to Vendors")

#Bin by 50K Sum(Amount) and Cnt and Spend
cleaned_po <- po_raw %>% filter(.data$`Sum Amount` >= 0 & !is.na(.data$`PO No.`) & .data$Status != c("O","PA", "PX"))
PO_and_Amt <- cleaned_po %>% group_by(.data$`PO No.`) %>% summarise(PO_Total_Amt = sum(.data$`Sum Amount`))

#50K Bins only
Bin_50k <- PO_and_Amt %>% mutate(Under_Equal_50K = .data$PO_Total_Amt >= 50000)
final_50K <- Bin_50k %>% group_by(.data$Under_Equal_50K) %>% summarise(Cnt = n(), Spend = sum(.data$PO_Total_Amt))

#3 Bins (0 < Amt <= 3,500 ; 3,500 < Amt <= 50,00; > 50,000)
Bin_3_levels <- PO_and_Amt %>% mutate(Bin_level = ifelse(.data$PO_Total_Amt <= 3500, "Micro", ifelse(.data$PO_Total_Amt > 3500 & .data$PO_Total_Amt < 50000, "Small", ifelse(.data$PO_Total_Amt >= 50000,"Large","Other"))))
Bin_3_Levels_Output <- Bin_3_levels %>% group_by(.data$Bin_level) %>% summarise(Cnt = n(), Spend = sum(.data$PO_Total_Amt))

#Add Percentage and make Pretty
Bin_3_Levels_Output <-  Bin_3_Levels_Output %>% mutate(Pct_Cnt = Cnt/sum(Cnt)*100, Pct_Spend = Spend/sum(Spend)*100)

#Factor Micro,Small,Large
Spend_threshhold_levels <- c("Micro","Small","Large")
Bin_3_Levels_Output$Bin_level <- factor(Bin_3_Levels_Output$Bin_level, levels = Spend_threshhold_levels)
Bin_3_Levels_Output <- Bin_3_Levels_Output %>% arrange(Bin_level)
Bin_3_Levels_Output <- Bin_3_Levels_Output %>% select(Bin_level,Cnt,Pct_Cnt,Spend, Pct_Spend)

#Write to CSV
write.csv(Bin_3_Levels_Output, file = "50K_Bin_Output.csv")




#4 - highlvl


# Written by Jenna Goldberg

library(tidyverse)
library(readxl)
setwd("O:/Codebase/0_FMIS_Source_Data")
raw <- read_excel("July1-2017_June15-2018_Pull/REQ_TO_PO_SPILT_PO_SIDE_PUBLIC_July1_June15.xlsx", skip = 1) #Note:Needs to be changed to the correct query (July 1st to June 30th)
Bids2018 <- read_excel("Bids2018_BizCenter.xlsx", col_names = FALSE, 
                       na = "N/A")
Bids2017 <- read_excel("Bids2017_BizCenter.xlsx", col_names=FALSE, na = "N/A")
setwd("O:/Codebase/2_PO Count and Spend by Platform/High Level Overview")

#Aggregating PO Data 
colnames(raw)[colnames(raw) == "PO No."] <- "POID"
colnames(raw)[colnames(raw) == "Req ID"] <- "REQID"
Clean <- raw %>%  group_by(.data$`POID`) %>% summarise(Amt = sum(.data$`Sum Amount`))
Clean <- left_join(raw, Clean, key = POID) %>% select(POID, Amt, Buyer, Bussiness_Unit, Origin, REQID, Vendor_Name, QuoteLink, `PO Date`) 
Clean <- Clean %>% distinct()

#Dealing with Biz Center Data - match via Req IDs from Bid Log Sheets. Will need a sheet of all req.ids from these. 
#Note that until 2018, these are split between Capital and Operating.  
All_Bids <- dplyr::bind_rows(Bids2017, Bids2018)
All_Bids <- All_Bids %>% distinct()
Bids_List <- c(All_Bids$X__1)

#Identify FairMarkIt by existence of quotelink, Origin = SWC: Platform = Commbuys, Business Center by matching Req IDs
# Not using info directly from FairMarkIt since the existence of an FM bid does not determine that the best bid was made there
Clean <- Clean %>% mutate(Platform = if_else(.data$Origin == "SWC", "Commbuys",  
                                             if_else(.data$REQID %in% Bids_List, "Business Center",
                                                     if_else(str_detect(.data$QuoteLink, "^h") == TRUE, "FairMarkIt"
                                                             , "Uncategorized"))))
#NA = Uncategorized
Clean$Platform [is.na(Clean$Platform)] <- "Uncategorized"
#remove duplicate PO#s caused by multiple matched Req IDs
Clean <- Clean %>% distinct(POID, .keep_all = TRUE)

#By Platform
POs_by_Platform <- Clean %>% group_by(Platform) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))

#Breakdown of Uncategorized, how many <50k?
Uncategorized <- Clean %>% filter(.data$Platform == "Uncategorized")
Uncategorized <- Uncategorized %>% mutate(Size = if_else(.data$Amt <50000, "<50,000", "50,000+"))
Size_of_Uncategorized <- Uncategorized %>% group_by(Size) %>% summarise("QTY of POs" = n())

#Buyer Types 
SE_Buyers <- c("AFLYNN","MBARRY","TDIONNE","HREIL")
SE_short <- Clean %>%  filter(.data$Buyer%in% SE_Buyers)
NON_SE <- Clean[!Clean$Buyer %in% SE_Buyers,]

Non_Inv <- c("JKIDD", "KAHERN", "CFRANCIS", "EKAJTAZI", "TDIONNE", "JLEBBOSSIERE", "KMCHUGH2","DMALONEY", "EGAGNON", "HREIL", "MBARRY", "KHALL", "TTOUSSAINT", "AFLYNN","MBERNSTEIN")
Inv <- c("DMARTINOS", "TSULLIVAN1", "SMORRISSEY", "PHONG", "RGIRADO", "KLOVE", "AKNOBEL")
Non_Inv_Short <- Clean %>% filter(.data$Buyer %in% Non_Inv)
Inv_Short <- Clean %>% filter(.data$Buyer %in% Inv)

Clean <- Clean %>% mutate(Buyer_Type = if_else(.data$Origin == "SWC", "SWC Buyer", 
                                               if_else(.data$Buyer %in% SE_Buyers, "Sourcing Executive", 
                                                       if_else(.data$Buyer%in%Non_Inv, "Non-Inventory", 
                                                               if_else(.data$Buyer %in% Inv, "Inventory", "NA")))))

Buyer_Type <- Clean %>% group_by(Buyer_Type) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))

#Trying to get number of Vendors per platform 
Vendor_List <- Clean %>% distinct(.data$Vendor_Name, .keep_all = TRUE)
Vendors_by_Platform <- Vendor_List %>% group_by(Platform) %>% summarise("QTY of Vendors" = n())
Vendors_and_POs_by_Platform <- left_join(Vendors_by_Platform, POs_by_Platform) %>% select(Platform, `QTY of Vendors`, `QTY of POs`)

#April 
Apr_Platform <- Clean %>% filter(lubridate::month(.data$`PO Date`, label = TRUE) == "Apr") %>% group_by(Platform) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))
Apr_Buyers <- Clean %>% filter(lubridate::month(.data$`PO Date`, label = TRUE) == "Apr") %>% group_by(Buyer_Type) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))

#May 
May_Platform <- Clean %>% filter(lubridate::month(.data$`PO Date`, label = TRUE) == "May") %>% group_by(Platform) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))
May_Buyers <- Clean %>% filter(lubridate::month(.data$`PO Date`, label = TRUE) == "May") %>% group_by(Buyer_Type) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))

#Data Frame w/ Platform, POID, Amount, Buyer, Origin, Biz Unit, Vendor 
Simple <- Clean %>% select(Platform, POID, Amt, Buyer, Origin, Bussiness_Unit, Vendor_Name)

#Export
write.csv(POs_by_Platform, "POs_by_Platform.csv")
write.csv(Buyer_Type, "POs_by_BuyerType.csv")
write.csv(May_Buyers, "May18_POs_By_BuyerType.csv")
write.csv(May_Platform, "May18_POs_By_Platform.csv")
write.csv(Apr_Buyers, "Apr18_POs_By_BuyerType.csv")
write.csv(Apr_Platform, "Apr18_POs_By_Platform.csv")
write.csv(Vendors_and_POs_by_Platform, "Vendors_and_POs_by_Platform.csv")




#5 - SpendByPlatformSub1000 #Note: This part of the script might not be necessary to our overall goal, but still includes the incorrect query which is why I'm including it - it may be relevant.


# Author: Jenna Goldberg
# Last Updated 6/25/18 by Jansen Manahan to add the <= $1000 filter
# Pulls the total spend and count of POs below $1000 divided by the platform on which they were solicited

library(tidyverse)
library(readxl)
setwd("O:/Codebase/0_FMIS_Source_Data")
raw <- read_excel("July1-2017_June15-2018_Pull/REQ_TO_PO_SPILT_PO_SIDE_PUBLIC_July1_June15.xlsx", skip = 1)
Bids2018 <- read_excel("Bids2018_BizCenter.xlsx", col_names = FALSE, na = "N/A")
Bids2017 <- read_excel("Bids2017_BizCenter.xlsx", col_names = FALSE, na = "N/A")
setwd("O:/Codebase/2_PO Count and Spend by Platform/Under 1K")

#Aggregating PO Data 
colnames(raw)[colnames(raw) == "PO No."] <- "POID" # Equivalent to rename()
colnames(raw)[colnames(raw) == "Req ID"] <- "REQID"
Clean <-
  raw %>%
  group_by(.data$`POID`) %>%
  summarize(Amt = sum(.data$`Sum Amount`)) # Amt is the sum amount across the PO number; NOTE THAT THIS DOES NOT ACCOUNT FOR BUSINESS UNIT
Clean <-
  left_join(raw, Clean, key = POID) %>% # mutate might be cleaner here
  select(POID, Amt, Buyer, Bussiness_Unit, Origin, REQID, Vendor_Name, QuoteLink, `PO Date`) # great place to rename columns
Clean <-
  Clean %>%
  distinct() # why?  Also, probably shouldn't overwrite "Clean"

#Dealing with Biz Center Data - match via Req IDs from Bid Log Sheets. Will need a sheet of all req.ids from these. 
#Note that until 2018, these are split between Capital and Operating.  
All_Bids <- 
  dplyr::bind_rows(Bids2017, Bids2018) %>% 
  distinct() # removes about 2/3s of records (but why is it undertaken?)
Bids_List <- c(All_Bids$X__1) # X__1 is req IDs. Has Nulls.

# Identify FairMarkIt by existence of quotelink, Origin = SWC: Platform = Commbuys, Business Center by matching Req IDs
# Not using info directly from FairMarkIt since the existence of an FM bid does not determine that the best bid was made there
# It is possible to use gather the winning bid data from the FM JSON file, but has not yet been done well enough to integrate it
Clean <- Clean %>% mutate(Platform = if_else(.data$Origin == "SWC", "Commbuys",  
                                             if_else(.data$REQID %in% Bids_List, "Business Center",
                                                     if_else(str_detect(.data$QuoteLink, "^h") == TRUE, "FairMarkIt"
                                                             , "Uncategorized"))))
#NA = Uncategorized
Clean$Platform[is.na(Clean$Platform)] <- "Uncategorized"
#remove duplicate PO#s caused by multiple matched Req IDs
Clean <- Clean %>% distinct(POID, .keep_all = TRUE)

Blacklist <- c(
  4000087408, # The requisition and the PO don't match because the req is from MBTAF ($890000) and the PO is from MBTAN ($150)
  4000088447  # This is excluded because it is now 2 cents, changed from $10,300 probably a a workaround for cancelling the PO
)

# Filter out high value POs, which affects all following CSVs
Clean <-
  Clean %>%
  filter(.data$Amt <= 1000, !.data$POID %in% Blacklist) # <= 1000 meshes with the definition of under $1000 from folder 3_Under1000

#By Platform
POs_by_Platform <-
  Clean %>%
  filter(!.data$POID %in% Blacklist) %>% 
  group_by(.data$Platform) %>%
  summarize("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))

#Export
write.csv(POs_by_Platform, "POs_by_Platform_Under_1K.csv")

# Use highlvl.r from the other folder under 2_By_Platform to get other information for POs under $1000 split by platorm
