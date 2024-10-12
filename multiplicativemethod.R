##### PACKAGES #####
library(readxl)
library(writexl)

#### READ FILE ####
setwd('C:/Users/Rodrigo/Desktop/5A1S/PGO/Rmodels')
df = read.csv('cd_dev.csv', sep = ';')

t = df$t
zt = df$Zt
n = nrow(df) #total number of observations
#m = frequency(df)
m = 12 #cycle frequency
alfa = 0.05


#### CALCULATE THE TREND ####
Mt = rep (NA, n)
i = 1+m/2 #initial start point for the moving average

for (i in (1+m/2):(n-m/2)) {
  M_delay = mean (zt[(i-m/2):(i+m/2-1)])
  M_adv = mean (zt[(i-m/2+1):(i+m/2)])
  Mt[i] = mean(c(M_delay, M_adv))
}

reg_mod = lm(Mt~t)
intercept = reg_mod$coefficients['(Intercept)']
slope = reg_mod$coefficients['t']

Trend = intercept + t*slope


#### CALCULATE THE CYCLE ####
Cycle = Mt/Trend


#### CALCULATE THE RATIO TIME SERIES ####
Ratio = zt/Mt 


#### CALCULATE THE SEASONALITY' ####
Sline = rep(NA,n)
Sratio = rep(NA,n)
LowBound = rep(NA,n)
UpBound = rep(NA,n)
OK = rep(NA,n)

get_elements <- function(lst, i, n) {
  # Select every n-th element starting from index i
  selected <- lst[seq(i, length(lst), by = n)]
  
  # Coerce the selected elements to numeric (non-numeric will become NA)
  numeric_selected <- as.numeric(selected)
  
  # Remove NA values (if any non-numeric values were present)
  numeric_selected <- numeric_selected[!is.na(numeric_selected)]
  
  return(numeric_selected)
}

#Calculate the Sline and the standard deviation
for (i in 1:m){
  Sline[i] = mean(as.numeric(get_elements(Ratio, i, m)))
  Sratio[i] = sd(as.numeric(get_elements(Ratio, i, m)))
}

for (i in (m+1):n){
  Sline[i] = Sline[i-m]
  Sratio[i] = Sratio[i-m]
}

#Calculate the upper and lower bound
LowBound = Sline - 2 * Sratio
UpBound = Sline + 2 * Sratio


#Check if value is between range
for (i in (m/2+1):(n-m/2)){
  if (Ratio[i] >= LowBound[i] & Ratio[i] <= UpBound[i]){
    OK[i] = "OK"
  } else {
    OK[i] = "NOT OK"
  }
}

ok_count <- sum(OK == "OK", na.rm = TRUE)          # Number of "OK"
print(paste("Percentage of OK's: ", round(100*ok_count/(n-m), digits = 2), "%"))

#Calculate final seasonality value
sum_Sline = sum(Sline[1:m])
Seasonality = Sline * m/sum_Sline


#### CALCULATE THE ERROR - INHERENT VARIABILITY ####
Errort = rep(NA, n)
for (i in (m/2+1):(n-m/2)){
  Errort[i] = Ratio[i]/Seasonality[i]
}

#Hypothesis testing
##Expected value of the Errors
average_errors = mean(Errort, na.rm = TRUE)
sdev_errors = sd(Errort, na.rm = TRUE)
stat_test = (average_errors - 1) / (sdev_errors / sqrt(n-m))

if (stat_test > qnorm(alfa/2) & stat_test < qnorm(1-alfa/2)){
  print("Hypothesis testing - Expected value of the errors:")
  print("Null hypothesis: The expected value of the errors is equal to 1")
  print("Alternative hypothesis: The expected value of the errors is different than 1")
  print(paste("For an alfa of ", alfa, "the result of the test was:"))
  print("Accept the null hypothesis")
} else{
  print("Hypothesis testing - Expected value of the errors:")
  print("Null hypothesis: The expected value of the errors is equal to 1")
  print("Alternative hypothesis: The expected value of the errors is different than 1")
  print(paste("For an alfa of ", alfa, "the result of the test was:"))
  print("There is statistical evidence that allows us to reject the null hypothesis")
}

##Correlation of the errors
ro_1 = cor(x = Errort[c((m/2+1):(n-m/2-1))], y = Errort[c((m/2+2):(n-m/2))])
stat_test1 = ro_1/(1/sqrt(n-m))
if (stat_test1 > qnorm(alfa/2) & stat_test1 < qnorm(1-alfa/2)){
  print("Hypothesis testing - Correlation of the errors:")
  print("Null hypothesis: The expected value of the correlation is 0")
  print("Alternative hypothesis: The expected value of the correlation is different than 0")
  print(paste("For an alfa of ", alfa, " the result of the test was:"))
  print("Accept the null hypothesis")
} else{
  print("Hypothesis testing - Correlation of the errors:")
  print("Null hypothesis: The expected value of the correlation is 0")
  print("Alternative hypothesis: The expected value of the correlation is different than 0")
  print(paste("For an alfa of ", alfa, "the result of the test was:"))
  print("There is statistical evidence that allows us to reject the null hypothesis")
}


#### CALCULATE THE PREDICTIONS ####
Forecast = Trend*Seasonality

#Forecast Error
Error_forecast = zt - Forecast

#Absolute Percentual Error
APE = abs(Error_forecast)/zt


#### BIAS AND VARIABLITY METRICS ####
#Bias
ME = mean(Error_forecast)
#Variability
MAE = mean(abs(Error_forecast))
MAPE = mean(APE)
MSE = mean(Error_forecast^2)


#### BUILDING A DATASET WITH ALL DATA ####
df = data.frame(
  t = t,
  zt = zt,
  Mt = Mt,
  Trend = Trend,
  Cycle = Cycle,
  Ratio = Ratio,
  Sline = Sline,
  Sratio = Sratio,
  LowBound = LowBound,
  UpBound = UpBound,
  OK = OK,
  Seasonality = Seasonality,
  Errort = Errort,
  Forecast = Forecast,
  Error_forecast = Error_forecast,
  APE = APE
)

stats = data.frame(
  Name = c('Exp. Value of the Errors', 'Correlation of the errors'),
  ST = c(stat_test, stat_test1)
)



#### CONVERT INTO EXCEL ####
write_xlsx(df, "MultiModel_R.xlsx")
write_xlsx(stats, "Stats_MultiModel_R.xlsx")

























