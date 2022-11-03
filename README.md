# financial_reporting_automation_script
This Python script enables you to automatically extract data from multiple Excel files, format it and combine everything in one report.

## 1. Problem statement
You are working as a Data Analyst and you need to extract data from multiple monthly financial reports that look something like this:

![Screen Shot 2022-11-03 at 11 23 14 AM](https://user-images.githubusercontent.com/87010794/199697256-593a9607-9b48-48ea-a931-3f4f130ddfb3.png)

Let's assume the financial department needs to analyze cost trends relative to 36 months; that would mean extracting information from at least 36
excel files..it sounds like a very time consuming job!

## 2. Solution
Why not create a tool that will automatically extract the relevant data from each excel file, reformat it and combine it all in one single report?
In this way, finance will be able to analyze trends for the last 3 years, through one single report without the need of extra resources.

The python script performs the following actions:
- Open every Excel Report located in a folder
- Processing and cleaning data
- Build the monthly report following a specific format
- Merge the monthly report with the global data frame
- Saving final results in an Excel file

The final result looks something like this:

![Screen Shot 2022-11-03 at 11 44 17 AM](https://user-images.githubusercontent.com/87010794/199701312-5e8ca6be-c5d3-4dee-92ef-fb912e3f8ec9.png)
