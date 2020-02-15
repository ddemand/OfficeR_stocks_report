# ===== create_ppt.R script

# Library installation and loading
if (!require(openxlsx)) install.packages('openxlsx')
library("openxlsx") # Open xlsx files to read and write data to and from
if (!require(officer)) install.packages('officer')
library("officer") # Open an MS PowerPoint file for generating automated analytics reports
# **Flextable requires 'gdtools' package via install of cairo-devel(CentOS/RHEL) "sudo yum install cairo-devel" or cairo(OSX)
if (!require(flextable)) install.packages('flextable')
library("flextable") # For writing and rendering nice tables to HTML with markdown, Word or PowerPoint
if (!require(rvg)) install.packages('rvg')
library("rvg") # Uses and API to produce vector graphics for embedding images to PowerPoint
if (!require(tidyverse)) install.packages('tidyverse')
library("tidyverse") # General purpose data wrangling
if (!require(lubridate)) install.packages('lubridate')
library("lubridate") # Eases datetime manipulation    
if (!require(jsonlite)) install.packages('jasonlite')
library("jsonlite") # Library fror working with JSON data in R
if (!require(ggplot2)) install.packages('ggplot2')
library("ggplot2") # It's ggplot.
if (!require(ggpubr)) install.packages('ggpubr')
library("ggpubr") # Used to fix scales of ggplot + facet's
if (!require(prophet)) install.packages('prophet')
library("prophet") # A package from Facebook developers for forecasting
if (!require(plotly)) install.packages('plotly')
library("plotly")
if (!require(dygraphs)) install.packages('dygraphs')
library("dygraphs") # A package for interactive plots
if (!require(DT)) install.packages('DT')
library("DT")


# ===== Using the next line to get stocks from stooq.com 

# **Choose different stocks by changing only the "symbol" variable**
### ======== ###
symbol <- 'AMZN'
### ======== ###

URL <- paste0('https://stooq.com/q/d/l/?s=', symbol, '.US&i=d')
df <- read.csv(URL)

# The function 'str()' provides a glimpse of the structure for the dataframe
str(df)

# Conversion of date intervals using the lubridate package
df$Date <- ymd(df$Date)
df$Month <- month(df$Date, abbr = FALSE)
df$Quarter <- quarter(df$Date)
df$Spread <- df$High - df$Low

# Assembling data using dplyr.  Select the 'df' element from the Global Environent to see the entire dataset
df <- df %>%
  mutate(Date = ymd(Date),
         Month = month(Date, abbr=FALSE),
         Year = year(Date),
         Quarter = quarter(Date),
         Spread = (High - Low)
  )

# See ten rows of the dataframe
head <- head(df,5)
head
tail <- tail(df,5)
tail 

# Create a quarterly dataset from the 'df' datafram
Quarterly <- df %>%
  select(Close,Year, Quarter) %>%
  group_by(Year, Quarter) %>%
  summarise_all(funs(median(as.numeric(.)))) %>%
  mutate(QoQ = Close - lag(Close)) %>%
  arrange(Year)


# ===== Create various interactive plots of the data using the 'plotly' package
daily <- ggplot(data = df, aes(x=Date, y=Close)) + geom_line()
ggplotly(daily)

daily_tl <- ggplot(data = df, aes(x=Date, y=Close)) + geom_line() +
  geom_smooth(method = loess, se = FALSE)
ggplotly(daily_tl)
ggsave(
  "daily_tl.png",
  daily_tl,
  width = 7,
  height = 5,
  dpi = 1200
)

# Historical daily, faceted by the avaialable years
yearly <- ggplot(data = df, aes(x=Date, y=Close)) + geom_line() + facet_wrap(~ year(Date), scales = "free") +
  theme(axis.text.x = element_text(face="plain", color="black",
                                   size=8, angle=-90),
        axis.text.y = element_text(face="plain", color="black",
                                   size=8, angle=0)) +
  geom_smooth(method = loess, se = FALSE, size =.5)
ggplotly(yearly)
yearly <- ggpubr::ggarrange(yearly)
ggsave(
  "yearly.png",
  yearly,
  width = 7,
  height = 5,
  dpi = 1200
)

#Boxplots for each month of each year
ggplot(df, aes(x = Date, y = Close, group = interaction( month(Date)))) + 
  geom_boxplot()

bp <- ggplot(plotbycity, aes(x = date, y = rent, group = city, color = city,
    text = paste('<br>Date: ', as.Date(date),
                 '<br>Obs: ', count))) +
                 geom_boxplot(data = df, aes(x=Date, y=Close))  
  theme(axis.text.x = element_text(face="plain", color="black",
                                   size=8, angle=-90),
        axis.text.y = element_text(face="plain", color="black",
                                   size=8, angle=0))
ggplotly(bp)

# Deviations
df_SD <- sd(df$Close)
df_SpreadSD <- sd(df$Spread)

# Empirical Rule Upper and Lower bounds of the distribution
Lower1SD68 <- round(mean(df$Close) - sd(df$Close),2)
Upper1SD68 <- round(mean(df$Close) + sd(df$Close),2)
Lower2SD95 <- round(mean(df$Close) - (sd(df$Close)*2),2)
Upper2SD95 <- round(mean(df$Close) + (sd(df$Close)*2),2)
Lower3SD99p7 <- round(mean(df$Close) - (sd(df$Close)*3),2)
Upper3SD99p7 <- round(mean(df$Close) + (sd(df$Close)*3),2)

# Histogram with density plot of all closing prices
histo <- ggplot(df, aes(x=Close )) +
  geom_histogram(aes(y=..density..), colour="black", fill="white",bins =25)+
  geom_density(alpha=.2, fill="#0FA2C7")
ggplotly(histo)
ggsave(
  "histo.png",
  histo,
  width = 7,
  height = 5,
  dpi = 1200
)


# ===== The following chunck leverages the Facebook Prophet package to provide simple forecasts of closing prices
# Use the following line if you need to apply filtering beyond any specific date
df2 <- filter(df,Date >= '2016-11-08')

# Create the date series, y-axis values
ds <- as.POSIXct(df2$Date)
y <- (df2$Close)

# Assemble the forecasting data frame
fcst <- data.frame(ds,y)

# Quick plot the data frame
qplot(ds,y)

# Set the seasonality
d <- prophet(fcst, daily.seasonality=TRUE)

# Forecast for one year
future <- make_future_dataframe(d, periods = 365)
forecast_df <- predict(d, future)

# Simple plot with projection of the forecast
plot(d, forecast_df)
ggsave(
  "forecast.png",
  plot(d, forecast_df),
  width = 7,
  height = 5,
  dpi = 1200
)

# Plot the components to see the various trends
prophet_plot_components(d, forecast_df)
jpeg("components.png", units="in", width=7, height=5, res=300)
prophet_plot_components(d, forecast_df)
dev.off()

# Interactive plot using Dygraphs
dyplot.prophet(d, forecast_df)
jpeg("dy_plot.png", units="in", width=7, height=5, res=300)
dyplot.prophet(d, forecast_df)
dev.off()

# Filter the dataset for the forecast data greater than or equal to todays date
fcst_data <- filter(forecast_df, ds >= Sys.Date())

# See a summary of the forecast data
summary(fcst_data)

# Plot of the forecast data
fcst_plt <- ggplot(data = fcst_data, aes(x=ds, y=yhat)) + geom_line()
ggplotly(fcst_plt)


# ====== Prepare and complete the PowerPoint Presentation Deck using methods learned from https://davidgohel.github.io/officer

# PowerPoint Comments and Variables
cap1 <- paste0("In this chart we can see the histogram of all daily closings. The median close is: $", median(df$Close))
cap2 <- "In this chart we can see the historical closings for each day and the respective trend."
cap3 <- "In this chart we can see the historical closings for each day of each year and the respective trend."
cap4 <- "In this chart we can see the projected forecasts."
cap5 <- "In this chart we can see the forecast components."
myftr <- paste0("Contains confidential company information. Not for external use or disclosure without proper authorization. Created on: ", Sys.Date())
median_close <- paste0("The median close is:", median(df$Close))
close_summary <- print(summary(df$Close))
yesterday <- autofit(flextable(head(df,1)))

# Read in MS PowerPoint Deck with OfficeR
my_ppt <- officer::read_pptx("company.pptx")

# Sample the PPT layout and master data
sums <- knitr::kable(layout_summary(my_ppt))
sums
props <- knitr::kable(layout_properties(my_ppt))
props

# Title Slide
my_ppt2 <- on_slide( x= my_ppt, index= 1) %>%
  ph_with(value= paste0("Historical stock prices on : ",symbol), location= ph_location_type(type = "title")) %>%
  ph_with(value= myftr, location= ph_location_type(type = "ftr")) %>%
  add_slide(., layout= "Title and Content", master= "Crop") %>%
  ph_with_text(type = "title", index=1,str = paste0("Yesterday's ", symbol, " Stock Price Summary: ")) %>%
  ph_with(., yesterday, location = ph_location_type(type = "body"), use_loc_size = TRUE ) %>%

# Slide 2
  add_slide(., layout= "Content with Caption", master= "Crop") %>%
  ph_with_text(type = "title", index=1,str = paste0("Histogram Plot of the ", symbol, " Daily Closings")) %>%
  ph_with_text(type = "body", index=1,str = cap1) %>%  
  ph_with(., value= external_img("/Users/daviddemand/R/Projects/SimpleStocksReport/histo.png", width = 6.75, height = 5.33),
          location = ph_location_type(type = "body", index=2), use_loc_size = FALSE ) %>%  

# Slide 3
  add_slide(., layout= "Content with Caption", master= "Crop") %>%
  ph_with_text(type = "title", index=1,str = "Historical Daily Closings with Trendline") %>%
  ph_with_text(type = "body", index=1,str = cap2) %>%  
  ph_with(., value= external_img("/Users/daviddemand/R/Projects/SimpleStocksReport/daily_tl.png", width = 6.75, height = 5.33),
          location = ph_location_type(type = "body", index=2), use_loc_size = FALSE ) %>%

# Slide 4
  add_slide(., layout= "Content with Caption", master= "Crop") %>%
  ph_with_text(type = "title", index=1,str = "Historical Daily Closings with Trendline: Yearly Facet") %>%
  ph_with_text(type = "body", index=1,str = cap3) %>%  
  ph_with(., value= external_img("/Users/daviddemand/R/Projects/SimpleStocksReport/yearly.png", width = 6.75, height = 5.33),
          location = ph_location_type(type = "body", index=2), use_loc_size = FALSE ) %>%
 
# Slide 5
  add_slide(., layout= "Content with Caption", master= "Crop") %>%
  ph_with_text(type = "title", index=1,str = "Forecasted Closing Prices using FB Prophet") %>%
  ph_with_text(type = "body", index=1,str = cap4) %>%
  ph_with(., value= external_img("/Users/daviddemand/R/Projects/SimpleStocksReport/forecast.png", width = 6.75, height = 5.33),
          location = ph_location_type(type = "body", index=2), use_loc_size = FALSE ) %>%    

# Slide 6
  add_slide(., layout= "Content with Caption", master= "Crop") %>%
  ph_with_text(type = "title", index=1,str = "Forecasted Components using Facebook Prophet") %>%
  ph_with_text(type = "body", index=1,str = cap5) %>%
  ph_with(., value= external_img("/Users/daviddemand/R/Projects/SimpleStocksReport/components.png", width = 6.75, height = 5.33),
          location = ph_location_type(type = "body", index=2), use_loc_size = FALSE ) %>%
  print(., target="StocksReport.pptx") %>%
  invisible()
