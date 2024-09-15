##MBA Reporting Project 

install.packages(c('plyr','stringdist', 'tidyr', 'rvest', 'stringr', 'reshape2'))
install.packages(c('stringdist'))

##load relevant libraries
library(plyr)
library(stringdist)
library(tidyr)
library(rvest)
library(stringr)
library(reshape2)

setwd('C:/Users/Yoonhee Kim/OneDrive - Drexel University/TA work/Winter 2024/')

##data cleansing
#load data
## load survey data from qualtrics, which has the second and third header rows removed
MBA_Full <- read.csv("C:/Users/Yoonhee Kim/OneDrive - Drexel University/TA work/Winter 2024/Leadership_Survey_651.csv", stringsAsFactors=F)
## load the data dictionary
MBA_Dict <- read.csv("C:/Users/Yoonhee Kim/OneDrive - Drexel University/TA work/Winter 2024/DD_Fall 2018_2019.csv", stringsAsFactors=F)
## load the sequence information, which includes national averages and report measures
MBA_Seq <- read.csv("C:/Users/Yoonhee Kim/OneDrive - Drexel University/TA work/Winter 2024/MBA_Cat_Seq1.csv", stringsAsFactors=F)

## Create duplicate dataframe, to retain original data
MBA <- MBA_Full

## Remove rows without unique identifiers
MBA <- MBA[!(MBA$Q1=="" & MBA$Q2==""),]

## Convert names (temporarily) to lower case
MBA$Q1 <- tolower(MBA$Q1)
MBA$Q2 <- tolower(MBA$Q2)


##tranform names to formal formatting

#function to reformat names 
simpleCap <- function(x) {
  s <- strsplit(x, " ")[[1]]
  paste(toupper(substring(s, 1, 1)), substring(s, 2),
        sep = "", collapse = " ")
}

#Apply the function to the first and last name
MBA$Q1 <-sapply(MBA$Q1, simpleCap)
MBA$Q1
MBA$Q2 <-sapply(MBA$Q2, simpleCap)
MBA$Q2

## Create a Full name column
MBA$Full <- paste(MBA$Q1, MBA$Q2)
MBA$Full

## Recode Programs to display names
MBA$Q3[MBA$Q3==1] <- "Malvern MBA"
MBA$Q3[MBA$Q3==2] <- "University City MBA"
MBA$Q3[MBA$Q3==3] <- "Full-Time MBA"
MBA$Q3[MBA$Q3==4] <- "Online MBA"

## Screen for duplicates
n_occur <- data.frame(table(MBA$Full))

#Return students with duplicate entries by last name
n_occur[n_occur$Freq > 1,]


##Remove incomplete survey responses according to threshold
#set threshold (ie .75)
threshold <- .75
#replace NULL values with NAs
MBA[MBA=="#NULL!"]<-NA

#set counting indices to count blank entries for scale answers (excluding network questions, which can be left blank)
i <- 25:146
ix <- c(i, 247:418)
MBA[ix] <- lapply(MBA[ix], as.numeric) 

## Calculate percent of survey left blank
MBA$Pct_Blank <- apply(is.na(MBA[ix]), 1, sum)/ncol(MBA[ix])
MBA$Pct_Blank

#save data before reduction
MBA_red <- MBA
#reduce data according to threshold
MBA <- MBA[!(MBA$Pct_Blank>=threshold),]


##Identify Duplicates
n_occur_2 <- data.frame(table(MBA$Full))
#Return students with duplicate entries by full name
n_occur_2[n_occur_2$Freq > 1,]


##If duplicates still remain, view duplicates and decide how to handle
##if deleting one of the rows, replace 5 with correct row number
#r <- 5
#MBA <- MBA[-5,]



### Recode Scales (Recoding indicator is located and read from the Data Dictionary)
Rec <- MBA_Dict[MBA_Dict$Recode_Ind==1 & MBA_Dict$Scale==7,]
Rec <- Rec[!is.na(Rec$New_Q),]
Rec_7 <- Rec$New_Q

##Recode Measures with 7 Point Likert Scale
#create list of variables that need to be recoded
#function to recode
rec7 <- function(x, column){
  reco_7 <- 8-x[,column]
  return(reco_7)
}
#loop to recode
for (i in 1:(length(Rec_7)-2)){
  MBA[[paste0(Rec_7[i])]]  <- rec7(MBA, Rec_7[i])
}


##calculate individuals mean values for variables
##only calculate for the student, not supervisor
#define variable list and variable groups

dims1 <- unique(MBA_Dict$Category)
dims1
dims1 <- dims1[!dims1 %in% c("N1","N2", "N3", "N4", "N5")]
dims1
# Remove ""
dims1 <- dims1[-1]
dims1

## Define variable list
for (o in 1:length(dims1)){
assign(dims1[o],MBA_Dict$New_Q[MBA_Dict$Category==dims1[o]& is.na(MBA_Dict$Category==dims1[o])==F& is.na(MBA_Dict$Supervisor_Ind==dims1[o])==T])
}


##calculate average variable values for each fellow
Variable_G_1 <- data.frame(matrix(0,nrow=nrow(MBA),ncol=length(dims1)))
names(Variable_G_1) <- dims1
for (i in 1:length(dims1)){
  Variable_G_1[,i] <- rowMeans(MBA[eval(as.name(dims1[i]))], na.rm = T)
}
Variable_G_1[,c("NA1")]
PS <- c("NA1", "II", "SA", "AS")
Variable_G_1$PS <- rowMeans(Variable_G_1[c("NA1", "II", "SA", "AS")], na.rm = T) 
EMP_L <- c("LBE", "PDM", "COACH", "INF", "ShowCon")
Variable_G_1$EMP_L <- rowMeans(Variable_G_1[c("LBE", "PDM", "COACH", "INF", "ShowCon")], na.rm = T)
PSY_EMP <- c("MEAN", "COMP", "SD", "IMP")
Variable_G_1$PSY_EMP <- rowMeans(Variable_G_1[c("MEAN", "COMP", "SD", "IMP")], na.rm = T)

## Calculate supervisor information
Super_Vars <- MBA_Dict[MBA_Dict$Supervisor_Ind==1,]
Super_Vars <- Super_Vars[!is.na(Super_Vars$New_Q),]
Supers_Cat <- unique(Super_Vars$Category)
Supers_Cat


for (p in 1:length(Supers_Cat)){
  assign(paste(Supers_Cat[p]),MBA_Dict$New_Q[MBA_Dict$Category==Supers_Cat[p]& is.na(MBA_Dict$Category==Supers_Cat[p])==F& is.na(MBA_Dict$Supervisor_Ind==Supers_Cat[p])==F])
}

Variable_super <- data.frame(matrix(0,nrow=nrow(MBA),ncol=length(Supers_Cat)))
names(Variable_super) <- Supers_Cat
for (i in 1:length(Supers_Cat)){
  Variable_super[,i] <- rowMeans(MBA[eval(as.name(Supers_Cat[i]))], na.rm = T)
}
Super_Avg <- colMeans(Variable_super, na.rm=T)
Super_Avg

##Combine names, mean variables and added varables
MBA_Avg <- data.frame(MBA$Q1, MBA$Q2, MBA$Full, MBA$Q3, MBA$DistributionChannel, Variable_G_1, stringsAsFactors = F)
names(MBA_Avg)[1] <- "First_Name"
names(MBA_Avg)[2] <- "Last_Name"
names(MBA_Avg)[3] <- "Full_Name"
names(MBA_Avg)[4] <- "Program"
names(MBA_Avg)[5] <- "Section"
MBA_Avg[,6:50] <- lapply(MBA_Avg[,6:50], as.numeric)
MBA_Avg[is.na(MBA_Avg)]<-NA


##add MBA average information to next blank rows

next_row <- nrow(MBA_Avg)+1
MBA_Avg[next_row,6:length(MBA_Avg)] <- colMeans(MBA_Avg[,6:ncol(MBA_Avg)], na.rm=T)
MBA_Avg[next_row,1] <- "Average"
MBA_Avg[next_row,2] <- "MBA"
MBA_Avg[next_row,3] <- "MBA Average"

next_row <- nrow(MBA_Avg)+1
MBA_Avg[next_row,6:length(MBA_Avg)] <- colMeans(MBA_Avg[MBA_Avg$Program=="Malvern MBA",6:ncol(MBA_Avg)], na.rm=T)
MBA_Avg[next_row,1] <- "Average"
MBA_Avg[next_row,2] <- "Malvern"
MBA_Avg[next_row,3] <- "Malvern MBA Average"

next_row <- nrow(MBA_Avg)+1
MBA_Avg[next_row,6:length(MBA_Avg)] <- colMeans(MBA_Avg[MBA_Avg$Program=="University City MBA",6:ncol(MBA_Avg)], na.rm=T)
MBA_Avg[next_row,1] <- "Average"
MBA_Avg[next_row,2] <- "University City"
MBA_Avg[next_row,3] <- "University City MBA Average"

next_row <- nrow(MBA_Avg)+1
MBA_Avg[next_row,6:length(MBA_Avg)] <- colMeans(MBA_Avg[MBA_Avg$Program=="Full-Time MBA",6:ncol(MBA_Avg)], na.rm=T)
MBA_Avg[next_row,1] <- "Average"
MBA_Avg[next_row,2] <- "Full-Time"
MBA_Avg[next_row,3] <- "Full-Time MBA Average"

next_row <- nrow(MBA_Avg)+1
MBA_Avg[next_row,6:length(MBA_Avg)] <- colMeans(MBA_Avg[MBA_Avg$Program=="Online MBA",6:ncol(MBA_Avg)], na.rm=T)
MBA_Avg[next_row,1] <- "Average"
MBA_Avg[next_row,2] <- "Online"
MBA_Avg[next_row,3] <- "Online MBA Average"


## Add National Average information
Nat_avg <- data.frame(MBA_Seq$Category,MBA_Seq$National_Avg, stringsAsFactors = F)
test <- match(colnames(MBA_Avg),Nat_avg[,1])
next_row <- nrow(MBA_Avg)+1

#add national average information to next blank row
MBA_Avg[next_row,] <- Nat_avg[test,2]
MBA_Avg[next_row,1] <- "Average"
MBA_Avg[next_row,2] <- "National"
MBA_Avg[next_row,3] <- "National Average"

##Reformat data for excel dropdown use
MBA_long <- melt(MBA_Avg, id.vars = c("First_Name", "Last_Name", "Full_Name", "Program", "Section"))
MBA_long$Type <- "Individual"
MBA_long$variable <- as.character(MBA_long$variable)
MBA_long$Type[MBA_long$First_Name=="Average"] <- "Average"


#export data as needed
#write.csv(MBA_Avg, "C:/Users/Chels/Desktop/mba_avg8222018.csv", row.names = F)



##Export
# write.csv(MBA_long, paste0("C:/Users/chels/Desktop/mba_",format(Sys.Date(),"%m%d%Y"),".csv"))


###Generate Supervisor Average Data
##Combine names, mean variables and added varables
MBA_Spr_Avg <- data.frame(MBA$Q1, MBA$Q2, MBA$Full, MBA$Q3, MBA$DistributionChannel, Variable_super, stringsAsFactors = F)
names(MBA_Spr_Avg)[1] <- "First_Name"
names(MBA_Spr_Avg)[2] <- "Last_Name"
names(MBA_Spr_Avg)[3] <- "Full_Name"
names(MBA_Spr_Avg)[4] <- "Program"
names(MBA_Spr_Avg)[5] <- "Section"
MBA_Spr_Avg[is.na(MBA_Spr_Avg)] <- NA
MBA_Avg[,6:50] <- lapply(MBA_Avg[,6:50], as.numeric)
MBA_Avg[is.na(MBA_Avg)]<-NA


#add fellows average information to next blank row

next_row <- nrow(MBA_Spr_Avg)+1
MBA_Spr_Avg[next_row,6:length(MBA_Spr_Avg)] <- colMeans(MBA_Spr_Avg[,6:ncol(MBA_Spr_Avg)], na.rm=T)
MBA_Spr_Avg[next_row,1] <- "Average"
MBA_Spr_Avg[next_row,2] <- "MBA"
MBA_Spr_Avg[next_row,3] <- "MBA Supervisor Average"

next_row <- nrow(MBA_Spr_Avg)+1
MBA_Spr_Avg[next_row,6:length(MBA_Spr_Avg)] <- colMeans(MBA_Spr_Avg[MBA_Spr_Avg$Program=="Malvern MBA",6:ncol(MBA_Spr_Avg)], na.rm=T)
MBA_Spr_Avg[next_row,1] <- "Average"
MBA_Spr_Avg[next_row,2] <- "Malvern"
MBA_Spr_Avg[next_row,3] <- "Malvern MBA Supervisor Average"

next_row <- nrow(MBA_Spr_Avg)+1
MBA_Spr_Avg[next_row,6:length(MBA_Spr_Avg)] <- colMeans(MBA_Spr_Avg[MBA_Spr_Avg$Program=="University City MBA",6:ncol(MBA_Spr_Avg)], na.rm=T)
MBA_Spr_Avg[next_row,1] <- "Average"
MBA_Spr_Avg[next_row,2] <- "University City"
MBA_Spr_Avg[next_row,3] <- "University City MBA Supervisor Average"

next_row <- nrow(MBA_Spr_Avg)+1
MBA_Spr_Avg[next_row,6:length(MBA_Spr_Avg)] <- colMeans(MBA_Spr_Avg[MBA_Spr_Avg$Program=="Full-Time MBA",6:ncol(MBA_Spr_Avg)], na.rm=T)
MBA_Spr_Avg[next_row,1] <- "Average"
MBA_Spr_Avg[next_row,2] <- "Full-Time"
MBA_Spr_Avg[next_row,3] <- "Full-Time MBA Supervisor Average"

next_row <- nrow(MBA_Spr_Avg)+1
MBA_Spr_Avg[next_row,6:length(MBA_Spr_Avg)] <- colMeans(MBA_Spr_Avg[MBA_Spr_Avg$Program=="Online MBA",6:ncol(MBA_Spr_Avg)], na.rm=T)
MBA_Spr_Avg[next_row,1] <- "Average"
MBA_Spr_Avg[next_row,2] <- "Online"
MBA_Spr_Avg[next_row,3] <- "Online MBA Supervisor Average"



MBA_Spr_long <- melt(MBA_Spr_Avg, id.vars = c("First_Name", "Last_Name", "Full_Name", "Program", "Section"))
MBA_Spr_long$Type <- "Supervisor"
MBA_Spr_long$Type[MBA_Spr_long$First_Name=="Average"] <- "Supervisor"
MBA_Spr_long$variable <- as.character(MBA_Spr_long$variable)
MBA_long$variable <- as.character(MBA_long$variable)

### Network Data

dims2 <- c("N1", "N2", "N3", "N4", "N5")

MBA_Ntwk <- data.frame(MBA$Q1, MBA$Q2,MBA$Full, MBA$Q3, MBA$DistributionChannel, MBA[,grep("Q23.",colnames(MBA))], stringsAsFactors = F)
names(MBA_Ntwk)[1] <- "First_Name"
names(MBA_Ntwk)[2] <- "Last_Name"
names(MBA_Ntwk)[3] <- "Full_Name"
names(MBA_Ntwk)[4] <- "Program"
names(MBA_Ntwk)[5] <- "Section"

nonNAs <- function(x) {
  as.vector(apply(x, 2, function(x) length(which(!is.na(x)))))
}

for (o in 1:length(dims2)){
  assign(dims2[o],MBA_Dict$New_Q[MBA_Dict$Category==dims2[o]& is.na(MBA_Dict$Category==dims2[o])==F])
}

##Obtain Fellows Network Average Data
##cross functional

MBA_NtwkQ1 <- data.frame(MBA$Full, MBA[,grep("Q23_._1",colnames(MBA))], stringsAsFactors = F)
MBA_NtwkQ1.2 <- data.frame(MBA$Full, MBA[,grep("Q23_.._1",colnames(MBA))], stringsAsFactors = F)
MBA_NtwkQ1 <- merge(MBA_NtwkQ1, MBA_NtwkQ1.2)
MBA_Netw_Avg <- data.frame(matrix(0,nrow=nrow(MBA_Ntwk),ncol=7))
MBA_Netw_Avg[,1:4] <- MBA_Ntwk[,1:4]
names(MBA_Netw_Avg) <- c(names(MBA_Ntwk[1:4]), "Cross_Func", "Ext", "High_Lvl")
nt_vars <- c("N1", "N2", "N3")


MBA_NtwkQ2 <- data.frame(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23_._2",colnames(MBA_Ntwk))], stringsAsFactors = F)
MBA_NtwkQ2.2 <- data.frame(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23_.._2",colnames(MBA_Ntwk))], stringsAsFactors = F)
MBA_NtwkQ2 <- merge(MBA_NtwkQ2, MBA_NtwkQ2.2)
MBA_Net_size <- apply(MBA_NtwkQ2[,6:length(MBA_NtwkQ2)], 1, function(x) length(which(!is.na(x))))
MBA_Q2_CF <- apply(MBA_NtwkQ2[,6:length(MBA_NtwkQ2)], 1, function(x) length(which(x==2)))
CF <- MBA_Q2_CF/MBA_Net_size
mean(CF, na.rm=TRUE)



##external contacts
MBA_NtwkQ3 <- data.frame(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23_._3",colnames(MBA_Ntwk))], stringsAsFactors = F)
MBA_NtwkQ3.2 <- data.frame(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23_.._3",colnames(MBA_Ntwk))], stringsAsFactors = F)
MBA_NtwkQ3 <- merge(MBA_NtwkQ3, MBA_NtwkQ3.2)
MBA_Net_sizeQ3 <- apply(MBA_NtwkQ3[,6:length(MBA_NtwkQ3)], 1, function(x) length(which(!is.na(x))))
MBA_Q3_EC <- apply(MBA_NtwkQ3[,6:length(MBA_NtwkQ3)], 1, function(x) length(which(x==2)))
EC <- MBA_Q3_EC/MBA_Net_sizeQ3
mean(EC, na.rm=TRUE)


##higher levels
MBA_NtwkQ4 <- data.frame(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23_._4",colnames(MBA_Ntwk))], stringsAsFactors = F)
MBA_NtwkQ4.2 <- data.frame(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23_.._4",colnames(MBA_Ntwk))], stringsAsFactors = F)
MBA_NtwkQ4 <- merge(MBA_NtwkQ4, MBA_NtwkQ4.2)
MBA_Net_sizeQ4 <- apply(MBA_NtwkQ4[,6:length(MBA_NtwkQ4)], 1, function(x) length(which(!is.na(x))))
MBA_Q4_HLC <- apply(MBA_NtwkQ4[,6:length(MBA_NtwkQ4)], 1, function(x) length(which(x==3)))
HLC <- MBA_Q4_HLC/MBA_Net_sizeQ4
mean(HLC, na.rm=TRUE)



net_breadth <- data.frame(MBA_Ntwk[,1:5], CF, EC, HLC, stringsAsFactors = F)
next_row <- nrow(net_breadth)+1
net_breadth[next_row,1] <- "Average"
net_breadth[next_row,2] <- "MBA"
net_breadth[next_row,3] <- "MBA Average"
net_breadth[next_row,6] <- mean(CF, na.rm=TRUE)
net_breadth[next_row,7] <- mean(EC, na.rm=TRUE)
net_breadth[next_row,8] <- mean(HLC, na.rm=TRUE)

#Reformat for excel
Network_breadth_long <- melt(net_breadth, id.vars = c("First_Name", "Last_Name", "Full_Name", "Program", "Section"))
Network_breadth_long$Type <- "Individual"
Network_breadth_long$Type[Network_breadth_long$First_Name=="Average"] <- "Average"


##Non-Participants
#################### Complete all the Non-participant codes AFTER running all other code
# Import the filled in 'Nonparts_Template.csv' from your computer  
nons <- read.csv("C:/Users/Yoonhee Kim/OneDrive - Drexel University/TA work/Winter 2024/Nonparts_Template.csv", stringsAsFactors = FALSE)


nons$First_Name <- tolower(nons$First_Name)
nons$Last_Name <- tolower(nons$Last_Name)

nons$First_Name <-sapply(nons$First_Name, simpleCap)
nons$Last_Name <-sapply(nons$Last_Name, simpleCap)
nons$Full_Name <- paste(nons$First_Name, nons$Last_Name)
nonp <- data.frame(matrix(0,nrow=nrow(nons),ncol=length(MBA_Avg)))
names(nonp) <- names(MBA_Avg)

nonp[,1] <- nons$First_Name
nonp[,2] <- nons$Last_Name
nonp[,3] <- nons$Full_Name
nonp[,4] <- nons$Program
nonp[,5] <- nons$Section

nonpslong <- melt(nonp, id.vars = c("First_Name", "Last_Name", "Full_Name", "Program", "Section"))
nonpslong$Type <- "Individual"
nonpslong$variable <- as.character(nonpslong$variable)

nonp_super <- data.frame(matrix(0,nrow=nrow(nons),ncol=length(MBA_Spr_Avg)))
names(nonp_super) <- names(MBA_Spr_Avg)
nonp_super[,1] <- nons$First_Name
nonp_super[,2] <- nons$Last_Name
nonp_super[,3] <- nons$Full_Name
nonp_super[,4] <- nons$Program
nonp_super[,5] <- nons$Section

nonpsuplong <- melt(nonp_super, id.vars = c("First_Name", "Last_Name", "Full_Name", "Program", "Section"))
nonpsuplong$Type <- "Supervisor"
nonpsuplong$variable <- as.character(nonpsuplong$variable)


##nonparts
nonp_NB <- data.frame(matrix(0,nrow=nrow(nons),ncol=length(net_breadth)))
names(nonp_NB) <- names(net_breadth)
nonp_NB[,1] <- nons$First_Name
nonp_NB[,2] <- nons$Last_Name
nonp_NB[,3] <- nons$Full_Name
nonp_NB[,4] <- nons$Program
nonp_NB[,5] <- nons$Section

nonpNBlong <- melt(nonp_NB, id.vars = c("First_Name", "Last_Name", "Full_Name", "Program", "Section"))




##Network Size and Strength

head(MBA_Ntwk[,6:25])
MBA_Ntwk[,6:25][MBA_Ntwk[,6:25]==2] <- 1
MBA_NtwkSS_Avg <- cbind(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23.5_",colnames(MBA_Ntwk))],MBA_Ntwk[,grep("Q23.1_",colnames(MBA_Ntwk))])
MBA_NtwkSS_St_Avg <- cbind(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23_._1",colnames(MBA_Ntwk))])
MBA_NtwkSS_St_Avg.2 <- cbind(MBA_Ntwk[,1:5],MBA_Ntwk[,grep("Q23_.._1",colnames(MBA_Ntwk))])
MBA_NtwkSS_St_Avg <- merge(MBA_NtwkSS_St_Avg, MBA_NtwkSS_St_Avg.2)
MBA_NtwkSS_Avg$Size <- apply(MBA_NtwkSS_St_Avg[,6:length(MBA_NtwkSS_St_Avg)], 1, function(x) length(which(!is.na(x))))


Weak <- apply(MBA_NtwkSS_St_Avg[,6:(length(MBA_NtwkSS_St_Avg))], 1, function(x) length(which(x==1)))
Strong <- apply(MBA_NtwkSS_St_Avg[,6:(length(MBA_NtwkSS_St_Avg))], 1, function(x) length(which(x==3)))
Total_Size <- MBA_NtwkSS_Avg$Size
Weak+Strong==Total_Size

MBA_NtwkSS_St_Avg$Weak <- apply(MBA_NtwkSS_St_Avg[,6:(length(MBA_NtwkSS_St_Avg))], 1, function(x) length(which(x==1)))
MBA_NtwkSS_St_Avg$Strong <- apply(MBA_NtwkSS_St_Avg[,6:(length(MBA_NtwkSS_St_Avg))], 1, function(x) length(which(x==3)))
MBA_NtwkSS_St_Avg$Total_Size <- MBA_NtwkSS_Avg$Size
mean(MBA_NtwkSS_St_Avg$Total_Size)

MBA_Str <- data.frame(MBA_NtwkSS_St_Avg[,1:5], Weak, Strong, Total_Size, stringsAsFactors = F)

next_row <- nrow(MBA_Str)+1
MBA_Str[next_row,1] <- "Average"
MBA_Str[next_row,2] <- "MBA"
MBA_Str[next_row,3] <- "MBA Average"
MBA_Str[next_row,6] <- round(mean(MBA_NtwkSS_St_Avg$Weak),2)
MBA_Str[next_row,7] <- round(mean(MBA_NtwkSS_St_Avg$Strong),2)
MBA_Str[next_row,8] <- round(mean(MBA_NtwkSS_St_Avg$Total_Size),2)

next_row <- nrow(MBA_Str)+1
MBA_Str[next_row,1] <- "Average"
MBA_Str[next_row,2] <- "Malvern"
MBA_Str[next_row,3] <- "Malvern MBA Average"
MBA_Str[next_row,6] <- round(mean(MBA_NtwkSS_St_Avg$Weak[MBA_NtwkSS_St_Avg$Program=="Malvern MBA"]),2)
MBA_Str[next_row,7] <- round(mean(MBA_NtwkSS_St_Avg$Strong[MBA_NtwkSS_St_Avg$Program=="Malvern MBA"]),2)
MBA_Str[next_row,8] <- round(mean(MBA_NtwkSS_St_Avg$Total_Size[MBA_NtwkSS_St_Avg$Program=="Malvern MBA"]),2)

next_row <- nrow(MBA_Str)+1
MBA_Str[next_row,1] <- "Average"
MBA_Str[next_row,2] <- "Full-Time"
MBA_Str[next_row,3] <- "Full-Time MBA Average"
MBA_Str[next_row,6] <- round(mean(MBA_NtwkSS_St_Avg$Weak[MBA_NtwkSS_St_Avg$Program=="Full-Time MBA"]),2)
MBA_Str[next_row,7] <- round(mean(MBA_NtwkSS_St_Avg$Strong[MBA_NtwkSS_St_Avg$Program=="Full-Time MBA"]),2)
MBA_Str[next_row,8] <- round(mean(MBA_NtwkSS_St_Avg$Total_Size[MBA_NtwkSS_St_Avg$Program=="Full-Time MBA"]),2)

next_row <- nrow(MBA_Str)+1
MBA_Str[next_row,1] <- "Average"
MBA_Str[next_row,2] <- "University City"
MBA_Str[next_row,3] <- "University City MBA Average"
MBA_Str[next_row,6] <- round(mean(MBA_NtwkSS_St_Avg$Weak[MBA_NtwkSS_St_Avg$Program=="University City MBA"]),2)
MBA_Str[next_row,7] <- round(mean(MBA_NtwkSS_St_Avg$Strong[MBA_NtwkSS_St_Avg$Program=="University City MBA"]),2)
MBA_Str[next_row,8] <- round(mean(MBA_NtwkSS_St_Avg$Total_Size[MBA_NtwkSS_St_Avg$Program=="University City MBA"]),2)

next_row <- nrow(MBA_Str)+1
MBA_Str[next_row,1] <- "Average"
MBA_Str[next_row,2] <- "Online"
MBA_Str[next_row,3] <- "Online MBA Average"
MBA_Str[next_row,6] <- round(mean(MBA_NtwkSS_St_Avg$Weak[MBA_NtwkSS_St_Avg$Program=="Online MBA"]),2)
MBA_Str[next_row,7] <- round(mean(MBA_NtwkSS_St_Avg$Strong[MBA_NtwkSS_St_Avg$Program=="Online MBA"]),2)
MBA_Str[next_row,8] <- round(mean(MBA_NtwkSS_St_Avg$Total_Size[MBA_NtwkSS_St_Avg$Program=="Online MBA"]),2)



SS_long <- melt(MBA_Str, id.vars = c("First_Name", "Last_Name", "Full_Name", "Program", "Section"))
SS_long$Type <- "Individual"
SS_long$Type[SS_long$First_Name=="Average"] <- "Average"


nonp_SS <- data.frame(matrix(0,nrow=nrow(nons),ncol=length(MBA_Str)))
names(nonp_SS) <- names(MBA_Str)
nonp_SS[,1] <- nons$First_Name
nonp_SS[,2] <- nons$Last_Name
nonp_SS[,3] <- nons$Full_Name
nonp_SS[,4] <- nons$Program
nonp_SS[,5] <- nons$Section

nonpSSlong <- melt(nonp_SS, id.vars = c("First_Name", "Last_Name", "Full_Name", "Program", "Section"))
nonpSSlong$variable <- as.character(nonpSSlong$variable)
nonpSSlong$Type <- "Individual"

nonpNBlong$variable <- as.character(nonpSSlong$variable)
nonpNBlong$Type <- "Individual"

Network_Data <- rbind(SS_long, Network_breadth_long, nonpNBlong, nonpSSlong)


#combine all long reports
MBA_Reporting_Measures <- rbind(MBA_long, MBA_Spr_long,nonpslong, nonpsuplong, Network_Data)
MBA_Reporting_Measures <- MBA_Reporting_Measures[order(MBA_Reporting_Measures$Last_Name, MBA_Reporting_Measures$First_Name),]


##Export
write.csv(MBA_Reporting_Measures, "C:/Users/Yoonhee Kim/OneDrive - Drexel University/TA work/Winter 2024/240221.csv", row.names=F)










