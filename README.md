# Stock-Analysis Project

## Overview of Project
Given a data set with a wide variety of different stocks and prices it was my job to create a user friendly VBA application that Steve could use to help him understand trends in the market. In addition, I looked to output answers as efficiently and swift as possible by refactoring code. With this challenge, I had to enable macros through Excel to access the power of VBA. With this program, I looped over stock data from 2017 or 2018 to determine the magnitude at which different types of stocks were being traded, the rate of return for the various stocks, and refactored code to make it run as fast as possible. 

## Results
By using tickers to determine the type of stock that was being displayed I came to several conclusions based on the data given to me. One is that in general, the stock categories from 2017 tended to have a successful return (11/12 major stocks types yielded a positive return on the investment for the year of 2017). However, in the year 2018, only two of the twelve stock types had a positive return. From this, we can conclude that people probably made money off their stock in 2017, while in 2018 there was a much higher chance for their stock to have gone down in value.

## Run Time and Results
Using the refactored code, I was able to make the exectuion time decrease. To decrease the run time, I found several factors that helped cut the processing time down. Small things like accessing worksheets before the loop ran so that they were not reopened at the beginning of each pass, and cutting down on unneccessary intermediate steps to make the code less choppy helped to decrease my run time.

### Refactored Code run for 2017 Data
In 2017, eleven of the tweleve major stock types were positive for the year. 


