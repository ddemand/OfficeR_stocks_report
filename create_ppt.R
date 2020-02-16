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

# ===== Set current working directory
setwd("~/R/Projects/OfficeR_stocks_report")

# ===== Get stocks using stooq.com 

# **Choose different stocks by changing only the "symbol" variable**
### ======== ###
symbol <- 'AMZN'
### ======== ###

URL <- paste0('https://stooq.com/q/d/l/?s=', symbol, '.US&i=d')
df <- read.csv(URL)

# The function 'str()' provides a glimpse of the structure for the dataframe
str(df)

# Assembling data using dplyr.  Select the 'df' element from the Global Environent to see the entire dataset
df <- df %>%
  mutate(Date = ymd(Date),
         Month = month(Date, label = TRUE, abbr = TRUE, locale = Sys.getlocale("LC_TIME")),
         Mon = month(Date, label = FALSE, abbr = TRUE, locale = Sys.getlocale("LC_TIME")),
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

# Create a monthly dataset from the 'df' datafram
Monthly <- df %>%
  select(Close,Year, Month,Mon) %>%
  group_by(Year, Month) %>%
  summarise_all(funs(median(as.numeric(.)))) %>%
  mutate(MoM = Close - lag(Close)) %>%
  arrange(Month,Year)

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
yearly <- ggplot(data = df, aes(x=Date, y=Close)) + geom_line() + 
  facet_wrap(~ year(Date), scales = "free") +
  theme(axis.text.x = element_text(face="plain", color="black",
                                   size=6, angle=-90),
        axis.text.y = element_text(face="plain", color="black",
                                   size=6, angle=0)) +
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
monthly_bp <- ggplot(Monthly, aes(x = Month, y = Close, group = interaction( Month))) + 
  geom_boxplot() +
  theme(axis.text.x = element_text(face="plain", color="black",
                                   size=8, angle=-90),
        axis.text.y = element_text(face="plain", color="black",
                                   size=8, angle=0)) + 
  scale_x_discrete(breaks=as.character(unique(unlist(Monthly$Month))),
                   labels=as.character(unique(unlist(Monthly$Month))))
ggplotly(monthly_bp)
ggsave(
  "monthly.png",
  monthly_bp,
  width = 7,
  height = 5,
  dpi = 1200
)

# Histogram with density plot of all closing prices
histo <- ggplot(df, aes(x=Close )) +
  geom_histogram(aes(y=..density..), colour="black", fill="white",bins =30)+
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
# df2 <- filter(df,Date >= '2016-11-08')
# ...or by a certain number of days back. In this case, three years back
df2 <- filter(df,Date >= (Sys.Date())-365*3 )

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
cap1 <- paste0("In this chart we can see the histogram of all ",symbol," daily closings. The median close is: $", median(df$Close))
cap2 <- paste0("The chart for this slide illustrates the historical closings for each day and the respective trend for ", symbol, ".")
cap3 <- paste0("Here we can see the historical closings for each day of each year and the respective trend of ", symbol, " stock price closings.")
cap4 <- paste0("The Monthly Boxplot's show the density for variances, the spread and medain of the closings of each respective month for", symbol)
cap5 <- paste0("This is the projected one year forecasts for ", symbol)
cap6 <- paste0("In this illustration we can see the ", symbol, " forecast components.")
myftr <- paste0("Contains confidential ", symbol," company information (Not really...). Not for external use or disclosure without proper authorization. Created on: ", Sys.Date())
median_close <- paste0("The median close is: $", median(df$Close), ". Yesterday's close of $", 
                       tail(df2$Close,1), " is a ", (tail(df2$Close,1)-(median(df$Close))/(median(df$Close)) *100), 
                       "% change from the median and a ", round((((tail(df2$Close,1)-(tail(df2$Close,2)[1]))/(tail(df2$Close,2)[1])) *100),3), 
                       "% change from the previous days close at $", tail(df2$Close,2)[1], ".")
close_summary <- print(summary(df$Close))
yesterday <- autofit(flextable(tail(df,1)))

# Read in MS PowerPoint Deck with OfficeR
my_ppt <- officer::read_pptx("company.pptx")

# Sample the PPT layout and master data
sums <- knitr::kable(layout_summary(my_ppt))
sums
props <- knitr::kable(layout_properties(my_ppt))
props

# Title Slide
my_ppt2 <- on_slide( x= my_ppt, index= 1) %>%
  ph_with(value= paste0("Historical stock prices on : ",symbol), location= ph_location_type(type = "ctrTitle")) %>%
  ph_with(value= paste0("A simple example of creating a PowerPoint presentation using R for the stock prices of ", symbol, "."), 
          location= ph_location_type(type = "subTitle")) %>%
  ph_with(value= myftr, location= ph_location_type(type = "ftr")) %>%
  
  #Slide 2  
  add_slide(., layout= "Title and Content", master= "Circuit") %>%
  ph_with_text(type = "title", index=1,str = paste0("The Latest Available ", symbol, "Stock Price Summary: ")) %>%
  ph_with(., yesterday, location = ph_location_type(type = "body"), use_loc_size = TRUE ) %>%
  
  # Slide 3
  add_slide(., layout= "Content with Caption", master= "Circuit") %>%
  ph_with_text(type = "title", index=1,str = paste0("Histogram Plot of the ", symbol, " Daily Closings")) %>%
  ph_with_text(type = "body", index=2,str = cap1) %>%  
  ph_with(., value= external_img("~/R/Projects/OfficeR_stocks_report/histo.png", width = 7, height = 5),
          location = ph_location_type(type = "body", index=1), use_loc_size = FALSE ) %>%  
  
  # Slide 4
  add_slide(., layout= "Content with Caption", master= "Circuit") %>%
  ph_with_text(type = "title", index=1,str = "Historical Daily Closings with Trendline") %>%
  ph_with_text(type = "body", index=2,str = cap2) %>%  
  ph_with(., value= external_img("~/R/Projects/OfficeR_stocks_report/daily_tl.png", width = 7, height = 5),
          location = ph_location_type(type = "body", index=1), use_loc_size = FALSE ) %>%
  
  # Slide 5
  add_slide(., layout= "Content with Caption", master= "Circuit") %>%
  ph_with_text(type = "title", index=1,str = "Historical Daily Closings with Trendline: Yearly Facet") %>%
  ph_with_text(type = "body", index=2,str = cap3) %>%  
  ph_with(., value= external_img("~/R/Projects/OfficeR_stocks_report/yearly.png", width = 7, height = 5),
          location = ph_location_type(type = "body", index=1), use_loc_size = FALSE ) %>%
  
  # Slide 6
  add_slide(., layout= "Content with Caption", master= "Circuit") %>%
  ph_with_text(type = "title", index=1,str = "Monthly Boxplot's for historical closings") %>%
  ph_with_text(type = "body", index=2,str = cap3) %>%  
  ph_with(., value= external_img("~/R/Projects/OfficeR_stocks_report/monthly.png", width = 7, height = 5),
          location = ph_location_type(type = "body", index=1), use_loc_size = FALSE ) %>%
  
  # Slide 7
  add_slide(., layout= "Content with Caption", master= "Circuit") %>%
  ph_with_text(type = "title", index=1,str = "Forecasted Closing Prices using Facebook Prophet Package") %>%
  ph_with_text(type = "body", index=2,str = cap4) %>%
  ph_with(., value= external_img("~/R/Projects/OfficeR_stocks_report/forecast.png", width = 7, height = 5),
          location = ph_location_type(type = "body", index=1), use_loc_size = FALSE ) %>%    
  
  # Slide 8
  add_slide(., layout= "Content with Caption", master= "Circuit") %>%
  ph_with_text(type = "title", index=1,str = "Forecasted Components using Facebook Prophet") %>%
  ph_with_text(type = "body", index=2,str = cap5) %>%
  ph_with(., value= external_img("~/R/Projects/OfficeR_stocks_report/components.png", width = 7, height = 5),
          location = ph_location_type(type = "body", index=1), use_loc_size = FALSE ) %>%
  print(., target="StocksReport.pptx") %>%
  invisible()
