# Homework_Challenge1
Initial add for challenge 1
# Kickstarting with Excel

## Overview of Project

This project is intended to analyze kickstarter campaign data to provide Louise additional insight on how campaigns fared relative to different measures and parameters.  The excel kickstarter file was used as basis for the analysis and can be found here [Kickstarter_Challenge](https://github.com/lavec0324/Homework_Challenge1/blob/main/Kickstarter_Challenge.xlsx)

## Purpose
This project is intended to analyze data from several different kickstarter campaigns.  More specifically, it is intended to provide insight on how what relationships can be deduced from outcomes relative to their launch dates and their funding goals.


## Analysis and Challenges

### Analysis of Outcomes Based on Launch Date
Utilizing Excel, a pivot table and line graph was created to portray a summed monthly view of successful, failed, and canceled outcomes only.  The excel file used can be found here [Kickstarter_Challenge](https://github.com/lavec0324/Homework_Challenge1/blob/main/Kickstarter_Challenge.xlsx) and the graph can be found here ![Theater Outcomes vs Launch](https://github.com/lavec0324/Homework_Challenge1/blob/main/resources/Theater_Outcomes_vs_Launch.png)

### Analysis of Outcomes Based on Goals
Utilizing Excel, a pivot table and line graph was created to portray a view of successful, failed, and canceled outcomes of kickstarter campaigns by goal ranges.  The excel file used can be found here [Kickstarter_Challenge](https://github.com/lavec0324/Homework_Challenge1/blob/main/Kickstarter_Challenge.xlsx) and the graph can be found here ![Outcomes Based on Goals](https://github.com/lavec0324/Homework_Challenge1/blob/main/resources/Outcomes_vs_Goals.png)

### Challenges and Difficulties Encountered
While there were no specific challenges with the data itself, the time to create forumulas for each outcome based on ranges was time consuming. Some of this was offset by creating a row that captured the name of the outcome in the cell as a lookup instead of using text.  This will also allow for easy updates in the future if the outcomes names are to change.  Additionaly in hindsight, creating defined names for each of the rows would have been more efficient and also referencing goal amounts in a cell rather than hardcoding the ranges.  Also the range amounts could have been cutoff at 20k being the upper limit so that the number of ranges was reduced.  Finally, I could get to the same chart by utilizing pivot tables and charts in a fraction of the time of the sumif fromulas.  This is included in the workbook as a hidden sheet called pivot_outcomes.

## Results
### What are two conclusions you can draw about the Theater Outcomes by Launch Date?

Conclusion 1: May appears to be a good time for successul outcomes for theater.  This may have something to do with activity picking up as we approach summer and people wanting to go out to experience theater events.  However we would need to analyze country by country as May has different temperatures in Northern and Southern hemispheres.

Conclusion 2: Month of the year does not seem to have an impact on when theater kickstarters fail.  Failures range from 33-52 over the year which represents a pretty consistent range especially when looked at as a percentage of total outcomes (even though this did not represent the live outcome).

### What can you conclude about the Outcomes based on Goals?

Conclusion 1: The highest of goals (above 45k) have the greatest failure rates.

Conclusion 2: Goals set below $1000 has the highest success rate at 76%.

Conclusion 3: There appears to be a relationship between goal ranges and successful kickstarters, i.e. lower goal ranges are easier to achieve than higher.  We do however so an anomaly between $25-$45k range however this can be somewhat discounted since the total number of outcomes in this range is only 25 and not representative of the total population. 

Conclusion 4: A majority of kickstarters take place in the lower range under 10k goals.

### What are some limitations of this dataset?

The limitations of this dataset include having data represented from 2009 to 2017 and may not represent current conditions, especially considering current pandemic conditions from 2020 forward.  Currencies are not standardized so goal and pledged amounts can be misleading based on exchange rate information relative to the US dollar.

### What are some other possible tables and/or graphs that we could create?

A scatter plot of time it takes for successful campaigns to kickoff relative to goal.  Additional graphs can include regional differences of previously created countries as well.
