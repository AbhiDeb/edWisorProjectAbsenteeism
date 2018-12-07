setwd("D:\\edWisor\\Project-II\\R Code")

rm(list = ls())

#Importing the libraries
library(xlsx)
library(DMwR)
library(ggplot2)
library(plyr)
library(forecast)
library(tseries)

# Importing the dataset
dataset = read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1, header = T)

#Adding Year column to group the data
dataset[1:113,'Year'] = 2007
dataset[114:358,'Year'] = 2008
dataset[359:570,'Year'] = 2009
dataset[571:740,'Year'] = 2010

#Converting the categorical variables to factors
factor_col = c('ID',"Reason.for.absence", "Month.of.absence","Day.of.the.week", "Seasons", "Disciplinary.failure",
                     "Education", "Son", "Social.drinker", "Social.smoker", "Pet")
for(i in factor_col)
{
  dataset[,i] = as.factor(dataset[,i])
}

#Handling missing values
missing_val = data.frame(apply(dataset,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(dataset)) * 100
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL
missing_val = missing_val[,c(2,1)]
write.csv(missing_val, "Miising_perc.csv", row.names = F)

#kNN Imputation for handling missing values
dataset = knnImputation(dataset, k = 3)

#Removing data with month value 0
dataset = dataset[1:737,]

#Outlier Analysis
numeric_index = sapply(dataset,is.numeric) #selecting only numeric

numeric_data = dataset[,numeric_index]

cnames = colnames(numeric_data)

for (i in 1:length(cnames))
   {
     assign(paste0("gn",i), ggplot(aes_string(y = (cnames[i]), x = "Absenteeism.time.in.hours"), data = subset(dataset))+ 
              stat_boxplot(geom = "errorbar", width = 0.5) +
              geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                           outlier.size=1, notch=FALSE) +
              theme(legend.position="bottom")+
              labs(y=cnames[i],x="Absenteeism.time.in.hours")+
              ggtitle(paste("Box plot of Absenteeism for",cnames[i])))
   }
 
## Plotting plots together
gridExtra::grid.arrange(gn1,gn5,gn2,ncol=3)
gridExtra::grid.arrange(gn6,gn7,ncol=2)
gridExtra::grid.arrange(gn8,gn9,ncol=2)
 
 
#Replace all outliers with NA and impute
#create NA on "custAge
for(i in cnames){
 val = dataset[,i][dataset[,i] %in% boxplot.stats(dataset[,i])$out]
 #print(length(val))
 dataset[,i][dataset[,i] %in% val] = NA
}
 
dataset = knnImputation(dataset, k = 3)

#Data preparation
linear_data = ddply(dataset, c("Year", "Month.of.absence"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
linear_data = ts(linear_data$Absenteeism.time.in.hours, frequency = 12,start = c(2007, 7))

#Splitting the data
train = window(linear_data, start = c(2007,7), end = c(2009,12))
test = window(linear_data, start = c(2010))

#View decomposition plot
plot(decompose(linear_data))

#Applying ARIMA model
arima_fit = auto.arima(train, trace = T)

#Forecasting
arimafore = forecast(arima_fit, h =18)

#Accuracy
accuracy(arimafore, test)
