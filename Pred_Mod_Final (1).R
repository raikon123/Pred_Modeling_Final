#clear memory
rm(list=ls())

install.packages("rpart")
install.packages("randomForest")
install.packages("xl")
install.packages("caTools")
install.packages("rpart.plot")
install.packages("ggraph")
install.packages("forestFloor")
install.packages("rlang", repos=c("http://rstudio.org/_packages", "http://cran.rstudio.com"))
install.packages("caret")
install.packages("tree")
install.packages("maptree")
install.packages("CORElearn")
install.packages("rtf")

library(rpart.plot)
library(rpart)	
library(readxl)
library(randomForest)
library(caTools)
library(party)
library(forestFloor)
library(caret)
library(tree)
library(CORElearn)
library(rtf)

final <- read_excel("C:\\Users\\njohnson1\\Desktop\\Pred_Mod\\Pred_Mod\\Final_Project_PredMod.xlsx", sheet = 3 )
finaldata <- data.frame(final)


#turn all blanks to NA
is.na(finaldata) <- finaldata == '' 
is.na(finaldata) <- finaldata == 'Null'

#remove columns not needed for calculation
#I added new columns to the excel file to take care of missing values
#The new rows have the same name with a "1" behind it
#I added a null to records where a 0 would be counted as a score

finaldata$Education <- NULL
finaldata$Children_cnt <- NULL
finaldata$NewCarOwnership <- NULL
finaldata$CashUsage <- NULL
finaldata$Credit_Limit1 <- NULL
finaldata$ENR <- NULL
finaldata$delinquency_Cards <- NULL
finaldata$LateFeeAmt_3M <- NULL
finaldata$NO_ATM <- NULL
finaldata$usage_Travel <- NULL
finaldata$usage_Medical <- NULL
finaldata$usage_Insurance <- NULL
finaldata$behav_score <- NULL
finaldata$Seg3m <- NULL
finaldata$CreditGrade_PL_Cards <- NULL
finaldata$Ref <-NULL
finaldata$OccupationCode <- NULL
finaldata$Seg3m1 <- NULL
finaldata$CreditGrade_PL_Cards1 <- NULL
finaldata$Occupation <- NULL
finaldata$Nationality <- NULL
finaldata$SMOKER_STAT <- NULL

#after GINI I removed the folowing
finaldata$salaryCreditFlag <- NULL
finaldata$usage_Travel1 <- NULL
finaldata$usage_Medical1 <- NULL
finaldata$usage_Insurance1 <- NULL
finaldata$deliquency_Cards1 <- NULL
finaldata$F_girp <- NULL
finaldata$F_inv <- NULL
finaldata$b_girp <- NULL
finaldata$b_inv <- NULL
finaldata$deliquency_Cards1 <- NULL
finaldata$NewCarOwnership1 <- NULL
finaldata$MaritalStat <- NULL
finaldata$NO_ATM1 <- NULL
finaldata$CashUsage1 <- NULL
finaldata$LateFeeAmt_3M1 <- NULL
finaldata$behav_Score1 <- NULL


#Maybe branch should be removed

finaldata$Branch <- NULL



#set categorical values
#finaldata$Nationality = as.factor(finaldata$Nationality)
#finaldata$Occupation = as.factor(finaldata$Occupation)
#finaldata$NewCarOwnership1 = as.factor(finaldata$NewCarOwnership1)
finaldata$CustomerSegment = as.factor(finaldata$CustomerSegment)
finaldata$MaritalStat = as.factor(finaldata$MaritalStat)
#finaldata$SMOKER_STAT = as.factor(finaldata$SMOKER_STAT)
#finaldata$salaryCreditFlag = as.factor(finaldata$salaryCreditFlag)
#finaldata$Seg3m1 = as.factor(finaldata$Seg3m1)
#finaldata$Branch = as.factor(finaldata$Branch)
finaldata$UW = as.factor(finaldata$UW)
finaldata$SEX = as.factor(finaldata$SEX)
finaldata$Education1 = as.factor(finaldata$Education1)
finaldata$PREM_FREQ = as.factor(finaldata$PREM_FREQ)
#finaldata$behav_Score1 = as.factor(finaldata$behav_Score1)
finaldata$Nationality1 = as.factor(finaldata$Nationality1)

finaldata$MaritalStatus1 =as.factor(finaldata$MaritalStatus1)

#summary of fields used
summary(finaldata)
#finaldatasummary <-RTF("Summarydata.doc",)  # this can be an .rtf or a .doc
write.csv(summary(finaldata),"finaldatasummary.csv")

finaltree <- rpart(UW ~ ., data=finaldata, cp=.02, method = "anova")
finaltree


set.seed(123)
#finalsample <- sample(nrow(finaldata),0.7 * nrow(finaldata))3
#train <- finaldata[finalsample, ]
#test <- finaldata[-finalsample, ]

split <- sample.split(finaldata, SplitRatio=0.7)
training_set = subset(finaldata, split == TRUE)
test_set = subset(finaldata, split == FALSE)

str(training_set)

finalmodel <- rpart(UW ~ ., data = training_set, na.action=na.exclude)

finalmodel1 <- randomForest(UW ~ ., data = finaldata, na.action = na.exclude)

#finalmodel2 <- randomForest(UW ~ Salary + Education1 + Children_cnt1 + NewCarOwnership1 + CustomerSegment + MaritalStat +
 #                             AGE_AT_ENTRY + SEX + PREM_FREQ + aum + F_dep + deposits + F_girp + b_girp + F_inv +
  #                            b_inv + F_cards + b_cardspfm + F_hl + b_hlpfm + f_personal_loan + salaryCreditFlag +
   #                           vintage_mth + Branch + CashUsage1 + Credit_Limit1 + ENR1 + deliquency_Cards1 + f_cards_open +
    #                          trb + LateFeeAmt_3M1 + NO_ATM1 + usage_Travel1 + usage_Medical1 + usage_Insurance1 + behav_Score1, 
     #                       data = training_set, ntree = 200, na.action = na.exclude)

#with this model and graph it shows that there isnt enough relationships between the data to use random forest.
finalmodel2 <- randomForest(UW ~ ., data = training_set, ntree = 100, na.action = na.exclude)
finalmodel2 
plot(finalmodel2)

#change to 20 trees it doesnt need 100
finalmodel20 <- randomForest(UW ~ ., data = training_set, ntree = 20, na.action = na.exclude)
finalmodel20 
plot(finalmodel20)

is.list(finalmodel2)
x<-unlist(finalmodel2)


#rpart
printcp(finalmodel)
summary(finalmodel)
plotcp(finalmodel)

plot(finalmodel, uniform = TRUE )
text(finalmodel, use.n = TRUE, all = TRUE, cex = .8)

#random forest
print(finalmodel2)
importance(finalmodel2) #returns GINI
table(predict(finalmodel2), training_set$UW)
plot(finalmodel2)

#plot(importance(finalmodel2))
varImpPlot(finalmodel20)
getTree(finalmodel2, k=1, labelVar=FALSE)

#Mengchu's cleaned data
dataW <- read_excel("c:\\users\\njohnson1\\downloads\\UnderwritingModel_W.xlsx", sheet = "UnderwritingModelW")
names(dataW)

#dataW$UW <- ifelse(dataW$UW=="nonSTD",1,0)
dataW$UW <- factor(dataW$UW)
dataW$Education <- factor(dataW$Education)
dataW$CustomerSegment <- factor(dataW$CustomerSegment)
dataW$Nationality <- ifelse(dataW$Nationality == "SG",1,0)
dataW$Nationality <- factor(dataW$Nationality)
dataW$Occupation <- factor(dataW$Occupation)
dataW$Children_cnt <- factor(dataW$Children_cnt)
dataW$MaritalStat <- factor(dataW$MaritalStat)
dataW$vintage_mth <- as.numeric(dataW$vintage_mth)
dataW$F_dep <- factor(dataW$F_dep)
dataW$F_girp <- factor(dataW$F_girp)
dataW$F_inv <- factor(dataW$F_inv)
dataW$trb[is.na(dataW$trb)] <- 0
dataW$trb <- as.numeric(dataW$trb)
dataW$SEX <- factor(dataW$SEX)
dataW$F_cards <- factor(dataW$F_cards)
dataW$F_hl <- factor(dataW$F_hl)
dataW$f_personal_loan<- factor(dataW$f_personal_loan)
dataW$f_cards_open <- factor(dataW$f_cards_open)
dataW$PREM_FREQ <- factor(dataW$PREM_FREQ)
dataW$salaryCreditFlag <- factor(dataW$salaryCreditFlag)
dataW$Branch_L <- factor(dataW$Branch_L)
finaldataM <- dataW[c(1,4,7,10,11,12,13,14,16,19,21,22,30,32,33)]
finaldataM


set.seed(123)

split1 <- sample.split(finaldataM, SplitRatio=0.7)
training_set1 = subset(finaldataM, split1 == TRUE)
test_set1 = subset(finaldataM, split1 == FALSE)

str(training_set1)

#100 trees
finalmodelM <- randomForest(UW ~ ., data = training_set1, ntree = 100, na.action = na.exclude)
finalmodelM 
plot(finalmodelM)

#change to 20 trees
finalmodelM20 <- randomForest(UW ~ ., data = training_set1, ntree = 20, na.action = na.exclude)
finalmodelM20 
plot(finalmodelM20)

#random forest
print(finalmodelM)
importance(finalmodelM) #returns GINI
table(predict(finalmodelM), training_set1$UW)

#plot(importance(finalmodel2))
varImpPlot(finalmodelM)
getTree(finalmodelM, k=1, labelVar=FALSE)



#tree package
#too many variable...32 ir less
tr <- tree(UW ~ ., data = finalmodelM)
summary(tr)
plot(tr); text(tr)

#party package
(ct = ctree(UW ~ ., data = training_set))
plot(ct, main = "nonSTD and STD")
#error predictions


#est class probabilities
tr.pred <- predict(ct, newdata = finaldata, type = "prob")
tr.pred

#maptree with CART
library(maptree) #part of maptree
library(cluster) #part of CART
draw.tree(clip.rpart(rpart(UW ~ ., data = finaldataM), best = 7), nodeinfo = TRUE)

#corelearn package
fit.random.forest <- CoreModel(UW ~ ., data = finaldataM)
plot(fit.random.forest)


#switch target variable for different trees


#fit <- randomForest(UW ~ ., data = finaldata)

#look to see if any values are blank

sapply(finaldata, function(x)all(is.na(x)))


list(finaldata)
