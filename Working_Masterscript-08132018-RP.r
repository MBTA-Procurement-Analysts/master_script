#READING/DPLYR:
#install.packages("tidyverse") + install.packages("readxl") + install.packages("janitor") + install.packages("lubridate")
#install.packages("kableExtra") + install.packages("knitr") +  install.packages("ggthemes") + install.packages("stringr")
library(tidyverse,readxl)
library(janitor,kableExtra)
library(knitr,ggthemes)
library(stringr,lubridate)

setwd("O:/Raw Data")
Raw <- readxl::read_excel("NO_MASTER.xlsx") 
Raw1 <- Raw %>% filter(!.data$Buyer %in% c("CGUIOD", "SMCDONNELL", "CCHEEK"),.data$Unit == "MBTAF" ) %>% rename(Amt = .data$`Sum Amount`) #1-MBTAF_Vendors_YTD
Raw2 <- Raw %>% rename("Business_Unit"="Unit","PO_No" = "PO No.", "PO_Date" = "PO Date", "Sum_Amount" = "Sum Amount", "Req_ID" = "Req ID", "Bid_Category" = "Bid Category") #2-Summary_by_Business_Unit_v3
Raw3 <- Raw #3-50K_Bins-Thresholds
Raw4 <- Raw #4-highlvl
Bids2017 <- read_excel("Bids2017_BizCenter_RP_Copy.xlsx", col_names=FALSE, na = "N/A")
Bids2018 <- read_excel("Bids2018_BizCenter_RP_Copy.xlsx", col_names = FALSE, na = "N/A")

#ANALYSIS:
Vendor_Total_Table <- Raw1 %>% group_by(.data$Vendor) %>% summarize(Vendor_Total = sum(.data$Amt)) %>% 
arrange(.data$Vendor_Total) %>% mutate("Vendor Total (MM$)" = Vendor_Total / 1000000)
Vendor_Num_POs <- Raw1 %>% distinct(.data$`PO No.`, .data$Vendor) %>% group_by(.data$Vendor) %>% 
summarize(Num_POs = n())
Vendor_Total_Table <- left_join(Vendor_Num_POs, Vendor_Total_Table, by = "Vendor") %>% arrange(desc(.data$Vendor_Total))
Vendor_Total_Table$Vendor <- factor(Vendor_Total_Table$Vendor, levels = as.character(Vendor_Total_Table$Vendor))
Millionaire_Purchasers <- Vendor_Total_Table %>% filter(.data$`Vendor Total (MM$)` >= 1)

clean_raw <- Raw2 %>% filter(.$Status != c("O","PA", "PX"))
twomonth_clean <- clean_raw %>% filter(.data$PO_Date >= ymd("2018-04-01") & .data$PO_Date < ymd("2018-05-01"))
twomonth_amt <- twomonth_clean %>% group_by(.data$Business_Unit) %>% summarize(twomonth_spend = sum(.data$Sum_Amount))
twomonth_cnt <- twomonth_clean %>% distinct(.data$PO_No, .keep_all = T) %>% group_by(.data$Business_Unit) %>% summarize(twomonth_cnt = n())
twomonth <- twomonth_amt %>% left_join(twomonth_cnt, by="Business_Unit")
onemonth_clean <- clean_raw %>% filter(.data$PO_Date >= ymd("2018-05-01") & .data$PO_Date < ymd("2018-06-01"))
onemonth_amt <- onemonth_clean %>% group_by(.data$Business_Unit) %>% summarize(onemonth_spend = sum(.data$Sum_Amount))
onemonth_cnt <- onemonth_clean %>% distinct(.data$PO_No, .keep_all = T) %>% group_by(.data$Business_Unit) %>% summarize(onemonth_cnt = n())
onemonth <- onemonth_amt %>% left_join(onemonth_cnt, by="Business_Unit")
YTD_amt <- clean_raw %>% group_by(.data$Business_Unit) %>% summarize(ytd_spend = sum(.data$Sum_Amount))
YTD_cnt <- clean_raw %>% distinct(.data$PO_No, .keep_all = T) %>% group_by(.data$Business_Unit) %>% summarize(ytd_cnt = n())
YTD <- YTD_amt %>% left_join(YTD_cnt, by="Business_Unit")
summary_biz_unit <- YTD %>% left_join(twomonth, by="Business_Unit")
summary_biz_unit <- summary_biz_unit %>% left_join(onemonth,by="Business_Unit") %>% replace(is.na(.),0) %>%  arrange(desc(.data$ytd_spend)) %>% adorn_totals()
summary_biz_unit <- summary_biz_unit %>% rename("Total Spend" = "ytd_spend", "Total No. of POs" = "ytd_cnt","Spend in April" = "twomonth_spend", "No. of POs in April" = "twomonth_cnt", "Spend in May" = "onemonth_spend","No. of POs in May" = "onemonth_cnt")
summary_biz_unit$`Total Spend` <- paste("$",format(round(summary_biz_unit$`Total Spend`,0), big.mark=","),sep="")
summary_biz_unit$`Spend in April` <- paste("$",format(round(summary_biz_unit$`Spend in April`,0), big.mark=","),sep="")
summary_biz_unit$`Spend in May` <- paste("$",format(round(summary_biz_unit$`Spend in May`,0), big.mark=","),sep="")
summary_biz_unit$`Total No. of POs` <- format(round(summary_biz_unit$`Total No. of POs`,0), big.mark=",")
summary_biz_unit$`No. of POs in April` <- format(round(summary_biz_unit$`No. of POs in April`,0), big.mark=",")
summary_biz_unit$`No. of POs in May` <- format(round(summary_biz_unit$`No. of POs in May`,0), big.mark=",")

cleaned_po <- Raw3 %>% filter(.data$`Sum Amount` >= 0 & !is.na(.data$`PO No.`) & .data$Status != c("O","PA", "PX"))
PO_and_Amt <- cleaned_po %>% group_by(.data$`PO No.`) %>% summarise(PO_Total_Amt = sum(.data$`Sum Amount`))
Bin_50k <- PO_and_Amt %>% mutate(Under_Equal_50K = .data$PO_Total_Amt >= 50000)
final_50K <- Bin_50k %>% group_by(.data$Under_Equal_50K) %>% summarise(Cnt = n(), Spend = sum(.data$PO_Total_Amt))
Bin_3_levels <- PO_and_Amt %>% mutate(Bin_level = ifelse(.data$PO_Total_Amt <= 3500, "Micro", ifelse(.data$PO_Total_Amt > 3500 & .data$PO_Total_Amt < 50000, "Small", ifelse(.data$PO_Total_Amt >= 50000,"Large","Other"))))
Bin_3_Levels_Output <- Bin_3_levels %>% group_by(.data$Bin_level) %>% summarise(Cnt = n(), Spend = sum(.data$PO_Total_Amt))
Bin_3_Levels_Output <-  Bin_3_Levels_Output %>% mutate(Pct_Cnt = Cnt/sum(Cnt)*100, Pct_Spend = Spend/sum(Spend)*100)
Spend_threshhold_levels <- c("Micro","Small","Large")
Bin_3_Levels_Output$Bin_level <- factor(Bin_3_Levels_Output$Bin_level, levels = Spend_threshhold_levels)
Bin_3_Levels_Output <- Bin_3_Levels_Output %>% arrange(Bin_level)
Bin_3_Levels_Output <- Bin_3_Levels_Output %>% select(Bin_level,Cnt,Pct_Cnt,Spend, Pct_Spend)

colnames(Raw4)[colnames(Raw4) == "PO No."] <- "POID"
colnames(Raw4)[colnames(Raw4) == "Req ID"] <- "REQID"
Clean <- Raw4 %>%  group_by(.data$`POID`) %>% summarise(Amt = sum(.data$`Sum Amount`))
Clean <- left_join(Raw4, Clean, key = POID) %>% select(POID, Amt, Buyer, Unit, Origin, REQID, QuoteLink, `PO Date`,Vendor) 
Clean <- Clean %>% distinct()
All_Bids <- dplyr::bind_rows(Bids2017, Bids2018)
All_Bids <- All_Bids %>% distinct()
Bids_List <- c(All_Bids$X__1)
Clean <- Clean %>% mutate(Platform = if_else(.data$Origin == "SWC", "Commbuys",  
                                     if_else(.data$REQID %in% Bids_List, "Business Center",
                                     if_else(str_detect(.data$QuoteLink, "^h") == TRUE, "FairMarkIt","Uncategorized"))))
Clean$Platform [is.na(Clean$Platform)] <- "Uncategorized"
Clean <- Clean %>% distinct(POID, .keep_all = TRUE)
POs_by_Platform <- Clean %>% group_by(Platform) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))
Uncategorized <- Clean %>% filter(.data$Platform == "Uncategorized")
Uncategorized <- Uncategorized %>% mutate(Size = if_else(.data$Amt <50000, "<50,000", "50,000+"))
Size_of_Uncategorized <- Uncategorized %>% group_by(Size) %>% summarise("QTY of POs" = n())
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
Vendor_List <- Clean %>% distinct(.data$Vendor, .keep_all = TRUE)
Vendors_by_Platform <- Vendor_List %>% group_by(Platform) %>% summarise("QTY of Vendors" = n())
Vendors_and_POs_by_Platform <- left_join(Vendors_by_Platform, POs_by_Platform) %>% select(Platform, `QTY of Vendors`, `QTY of POs`)
Apr_Platform <- Clean %>% filter(lubridate::month(.data$`PO Date`, label = TRUE) == "Apr") %>% group_by(Platform) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))
Apr_Buyers <- Clean %>% filter(lubridate::month(.data$`PO Date`, label = TRUE) == "Apr") %>% group_by(Buyer_Type) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))
May_Platform <- Clean %>% filter(lubridate::month(.data$`PO Date`, label = TRUE) == "May") %>% group_by(Platform) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))
May_Buyers <- Clean %>% filter(lubridate::month(.data$`PO Date`, label = TRUE) == "May") %>% group_by(Buyer_Type) %>% summarise("QTY of POs" = n(), "Total Spend" = sum(.data$Amt))
Simple <- Clean %>% select(Platform, POID, Amt, Buyer, Origin, Unit, Vendor)

#OUTPUT:
write_csv(Vendor_Total_Table, "O:/Raw Data/Vendors with MBTA% spend YTD.csv")
write_csv(Millionaire_Purchasers, "O:/Raw Data/Millionaire_Purchasers.csv")

write_csv(summary_biz_unit, "O:/Raw Data/Summary_by_Business_Unit.csv")

write.csv(Bin_3_Levels_Output, file = "50K_Bin_Output.csv")

write.csv(POs_by_Platform, "POs_by_Platform.csv")
write.csv(Buyer_Type, "POs_by_BuyerType.csv")
write.csv(May_Buyers, "May18_POs_By_BuyerType.csv")
write.csv(May_Platform, "May18_POs_By_Platform.csv")
write.csv(Apr_Buyers, "Apr18_POs_By_BuyerType.csv")
write.csv(Apr_Platform, "Apr18_POs_By_Platform.csv")
write.csv(Vendors_and_POs_by_Platform, "Vendors_and_POs_by_Platform.csv")