#Remove all objects stored
rm(list=ls())

#Setting working directory
setwd("D:\\edwisor\\Employee_Absentism")

#Install multiple packages at a time
install.packages(c("corrgram","DMwR","caret","ggplot2","randomForest",
                   "gbm","ROSE","xlsx","DataCombine"))
install.packages("gridExtra")

#loading libraries
library(ggplot2)
library(corrgram)
library(DMwR)
library(caret)
library(randomForest)
library(dummies)
library(gbm)
library(xlsx)
library(DataCombine)

#loading the data
data= read.xlsx('Absenteeism_at_work_Project.xls', sheetIndex = 1)

#first 5 columns of data
head(data,5)

#Shape of data
dim(data)

#structure of data
str(data)

# Variable in the data
colnames(data)

#'con' and 'cat' are continuous and categorical variables respectively
con = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
                    'Work.load.Average.day.', 'Transportation.expense',
                    'Hit.target', 'Weight', 'Height', 
                    'Body.mass.index', 'Absenteeism.time.in.hours')

cat = c('ID','Reason.for.absence','Month.of.absence','Day.of.the.week',
                     'Seasons','Disciplinary.failure', 'Education', 'Social.drinker',
                     'Social.smoker', 'Son', 'Pet')

#Box plot of continuous data
#Distance from Residence to Work
ggplot(aes_string(y = "Distance.from.Residence.to.Work" , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

#Service time
ggplot(aes_string(y = "Service.time" , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

#Age
ggplot(aes_string(y = "Age" , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

#'Work.load.Average/day '
ggplot(aes_string(y = "Work.load.Average.day. " , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

#Transportation Expense
ggplot(aes_string(y = "Transportation.expense" , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

#Hit.target
ggplot(aes_string(y = "Hit.target" , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

#Height
ggplot(aes_string(y = "Height" , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

#Weight
ggplot(aes_string(y = "Weight" , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

#Body mass Index
ggplot(aes_string(y = "Body.mass.index" , x = "Absenteeism.time.in.hours"), data = subset(data))+ 
stat_boxplot(geom = "errorbar", width = 0.5)+
geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,outlier.size=1, notch=FALSE)                                                  

##MISSING VALUE ANALYSIS-------------------------------------------------------------

missing_val = data.frame(apply(data,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(data)) * 100
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL
missing_val = missing_val[,c(2,1)]

#data[61,20]
#data[61,20]= NA
#original value = 31
#mean = 26.67797     
#median = 25      
#KNN = 31         

#Mean
#data$Body.mass.index[is.na(data$Body.mass.index)] = mean(data$Body.mass.index, na.rm = T)

#Median
#data$Body.mass.index[is.na(data$Body.mass.index)] = median(data$Body.mass.index, na.rm = T)

#kNN Imputation
data = knnImputation(data, k = 3)

sum(is.na(data))

##OUTLIER ANALYSIS----------------------------------------


#Replace all outliers with NA and impute
for(i in con)
{
  val = data[,i][data[,i] %in% boxplot.stats(data[,i])$out]
  #print(length(val))
  data[,i][data[,i] %in% val] = NA
}

#knn imputation
data = knnImputation(data, k = 3)

#Feature Selection-----------------------------------------------------------------

#Correlation Plot 
corrgram(data[,con], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

#ANOVA test for Categprical variable
summary(aov(formula = Absenteeism.time.in.hours~ID,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Reason.for.absence,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Month.of.absence,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Day.of.the.week,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Seasons,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Disciplinary.failure,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Education,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Social.drinker,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Social.smoker,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Son,data = data))
summary(aov(formula = Absenteeism.time.in.hours~Pet,data = data))

## Dimension Reduction
data = subset(data, select = -c(Weight))

#Feature Scaling-------------------------------------------------------------------

#checking distribution of some variables
hist(data$Distance.from.Residence.to.Work)
hist(data$Service.time)
hist(data$Work.load.Average.day.)
hist(data$Age)
hist(data$Body.mass.index)
hist(data$Absenteeism.time.in.hours)

#since the distributions are not normal therefore we will normalize the data


con = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
        'Work.load.Average.day.', 'Transportation.expense',
        'Hit.target', 'Height','Body.mass.index')

for(i in con)
{
  data[,i] = (data[,i] - min(data[,i]))/(max(data[,i])-min(data[,i]))
}

#creating dummy variables for categorical data
library(mlr)
data = dummy.data.frame(data, cat)

  
#Modeling---------------------------------------------------------------------------

#Divide data into train and test using stratified sampling method
set.seed(123)
train.index = sample(1:nrow(data), 0.8 * nrow(data))
train = data[ train.index,]
test  = data[-train.index,]

##-------------------------------Random Forest-------------------------------------
RF = randomForest(Absenteeism.time.in.hours~., data = train)

#predicting for test data
y_pred = predict(RF,test[,names(test) != "Absenteeism.time.in.hours"])

print(postResample(pred = y_pred, obs = test[,107]))

##-----------------------------Linear Regression-----------------------------------
  
LR = lm(Absenteeism.time.in.hours ~ ., data = train)

#predicting for test data
y_pred = predict(LR,test[,names(test) != "Absenteeism.time.in.hours"])

# For testing data 
print(postResample(pred = y_pred, obs =test$Absenteeism.time.in.hours))


##------------------------------XGboost--------------------------------------------

XGB = gbm(Absenteeism.time.in.hours~., data = train)

#predicting for test data
y_pred = predict(XGB,test[,names(test) != "Absenteeism.time.in.hours"], n.trees = 500)

# For testing data 
print(postResample(pred = y_pred, obs = test[,107]))



#----------------------------PCA---------------------------------------------------
#principal component analysis
prin_comp = prcomp(train)

#compute standard deviation of each principal component
std_dev = prin_comp$sdev

#compute variance
pr_var = std_dev^2

#proportion of variance explained
prop_varex = pr_var/sum(pr_var)

#cumulative scree plot
plot(cumsum(prop_varex), xlab = "Principal Component",
     ylab = "Cumulative Proportion of Variance Explained",
     type = "b")

#add a training set with principal components
train.data = data.frame(Absenteeism.time.in.hours = train$Absenteeism.time.in.hours, prin_comp$x)

# From the above plot selecting 45 components since it explains almost 95+ % data variance
train.data =train.data[,1:45]

#transform test into PCA
test.data = predict(prin_comp, newdata = test)
test.data = as.data.frame(test.data)

#select the first 45 components
test.data=test.data[,1:45]

#----------------------------------Random forest---------------------------------

#Develop Model on training data
FR = randomForest(Absenteeism.time.in.hours~., data = train.data)

#Lets predict for testing data
y_pred = predict(FR,test.data)

# For testing data 
print(postResample(pred = y_pred, obs = test$Absenteeism.time.in.hours))

#---------------------------------Linear Regression--------------------------------------

#Develop Model on training data
RL = lm(Absenteeism.time.in.hours ~ ., data = train.data)

#Lets predict for testing data
y_pred = predict(RL,test.data)

# For testing data 
print(postResample(pred = y_pred, obs =test$Absenteeism.time.in.hours))

##------------------------------XGboost--------------------------------------------

BGX = gbm(Absenteeism.time.in.hours~., data = train.data)

#predicting for test data
y_pred = predict(BGX,test.data,n.trees = 500)

# For testing data 
print(postResample(pred = y_pred, obs = test$Absenteeism.time.in.hours))









