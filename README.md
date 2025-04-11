# Navigating Marketing Data Analysis GitHub

To run any of our code, copy this repository to your local computer. 
1. Open up a terminal/command prompt/windows shell on your computer.
2. Type in "cd Downloads"
3. Type in this command "git clone https://github.com/vws21/Marketing-Data-Analysis.git"
4. Type in "cd Marketing-Data-Analysis"
5. If you want to run our analysis code, type in "cd analysis". Then for example, "cd nikhita". Then type in "py analysis.py". If that doesn't work, type in "python analysis.py". If that doesn't work type in python3 analysis.py".
6. If you want to run our data cleaning code, type in "cd market_data_split.py". Then type in "py market_data_split.py". If that doesn't work, type in "python market_data_split.py". If that doesn't work type in market_data_split.py".
7. If you want to run any of the R files in the folders, you need RStudio and can open the R files on RStudio and just hit the play button or Run All Code Blocks button on the top right side of the screen.
8. Reach out to niv32@pitt.edu if you have trouble accessing anything!


### Analysis Folder
#### Nikhita Folder
  This folder contains all the graphs produced by analysis.py.
    These are the questions answered in this file:
  - What is the engagement of longer vs shorter subject lines?
  - What is the engagement of longer vs shorter emails?
  - What is the engagement in emails by topic?
  - What is the engagement of emails that have no personalization fields vs. 1, 2, or more personalization fields?

#### Brayden Folder
This folder contains find.py.
 These are the questions answered in this file:
 - What percentage of students receive all or most of the emails in the Prospective Student journey? What percentage of students receive >50% of the Prospective Student emails before progressing to the Applicant stage? 
 -  What percentage of students receive all or most of the emails that are in the Admitted/Matriculated journeys? What percentage of students receive >50% of these emails before enrolling or withdrawing (thereby being removed from the email list?)

#### Vince Folder
  This folder contains charts and multiple .py files to answer the following questions:
  -  With metrics such as open rates, click-through rates, and read times, which emails are opened and/or engaged with the most, and in what ways? (topEmails.py)
  -  What are student’s preferences for timing, frequency, topic, length, etc., of emails? Is this preference different based on stage? (timingPref.py, topicPrev.py, lengthPref.py)

### Data Cleaning Folder
  This folder contains market_data_split.py which is how we used Python to clean our dataset. 

### Datasets Folder
#### Updated_Marketing_Analysis.xlsx
  This is the excel workbook produced by market_data_split.py


#### rawdata.xlsx
  This is the excel workbook provided by the client before any editing or any cleaning.


#### Marketing Analysis Data.xlsx
  This is the excel workbook produced after all group members manually cleaned the sheets.


#### Cleaned_Updated_Marketing_Data.xlsx
  This is the excel workbook produced after manually cleaning  Updated_Marketing_Analysis.xlsx.
