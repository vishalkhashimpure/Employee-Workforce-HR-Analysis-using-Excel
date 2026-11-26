# Employee-Workforce-HR-Analysis-using-Excel

1. Business Problem Statement
An organization manages a workforce across multiple departments, job roles, and geographical locations. The HR department faces challenges in understanding employee demographics, workforce distribution, salary structures, employee satisfaction, performance, overtime, absenteeism, and employee turnover.

The organization wants to use historical employee data to identify workforce patterns, evaluate employee performance, understand factors associated with employee exits, and improve employee retention.

As a Data Analyst, your responsibility is to clean, organize, analyze, and visualize the HR dataset using Microsoft Excel. Use Excel formulas, functions, PivotTables, PivotCharts, conditional formatting, and interactive dashboard features to generate meaningful insights and support HR decision-making.
2. Project Objectives
1. Understand the organization's workforce distribution.
2. Analyze employee demographics.
3. Evaluate department-wise employee distribution.
4. Analyze salary structures across departments.
5. Understand employee performance.
6. Evaluate employee satisfaction.
7. Identify employee turnover patterns.
8. Analyze overtime and absenteeism.
9. Identify workforce-related trends.
10. Develop an interactive HR Analytics Dashboard.
Steps in the Project
Step 1: Understand the Data and Perform Data Cleaning
1. Load and Understand the Dataset
Load the dataset and understand its structure, columns, and data.
2. Check and Correct Data Types
Check all the data types and correct them. Ensure that date columns are stored as dates, numerical columns as numbers, and salary columns in the appropriate currency format.
3. Check for Duplicates in the Primary Column
Check for duplicates in the primary column, which is Employee ID. If there are any duplicates, remove them. Other columns can contain duplicate values.

To do this, select all the data and go to the Data tab. Click Remove Duplicates, select My data has headers, and mark only the Employee ID column. Then click OK. The duplicate rows will be removed.
4. Check for Null Values
Check for null values column by column.
5. Check Employee ID
First, check the Employee ID column. If there are any blank cells, remove the entire row using the Remove Null Rows method, as we did for duplicates. Use the COUNTBLANK function to identify blank cells.
6. Check the Age Column
Check the Age column for missing values. If there are any blank cells, replace them with the average age.

Calculate the average age using the AVERAGE function. Then select the Age column, excluding the header, press Ctrl + G, select Special, choose Blanks, and click OK. Replace the blank cells with the average age.
7. Check the Gender Column
Check the Gender column. If there are any missing values, replace them with "Unknown".
8. Check the Hire Date Column
Check the Hire Date column. If there are any blank cells, remove the entire row because, without a joining date, we cannot verify when the employee joined the company.

Select the Hire Date column, press Ctrl + G, select Special, choose Blanks, and click OK. Then press Ctrl + - and select Entire Row.
9. Check the Exit Date Column
Do not change anything in the Exit Date column because a missing Exit Date means that the employee is currently active in the company.
10. Check the Salary Column
Check the Salary column. If there are any blank cells, replace them with the median salary.

Calculate the median using the MEDIAN function. Then select the blank cells using Ctrl + G → Special → Blanks and replace them with the median salary.
11. Check Numerical Columns
In every numerical column, replace blank values with the average of that column. Press Ctrl + G, select Special, choose Blanks, and enter the average value. Then press Ctrl + Enter.
12. Check the City Column
Check the City column. If there are any blank cells, replace them with "Unknown City".

Press Ctrl + G, select Special, choose Blanks, and type "Unknown City". Then press Ctrl + Enter.
13. Check the Gender Column for Spelling Mistakes
Check the Gender column for spelling mistakes or inconsistent entries. Correct them to Male, Female, Non-binary, and Unknown using the Find and Replace method.

Select the Gender column, press Ctrl + H, enter the incorrect spelling in Find What, and enter the correct spelling in Replace With. Click Replace All and repeat until you have only the four standard options: Male, Female, Non-binary, and Unknown.
14. Check the Employee Type Column
Use the Find and Replace method to correct invalid spellings in the Employee Type column.

Repeat the process until you have only the four standard options: Full-Time, Part-Time, Contract, and Intern.
15. Check the City Names
Check the City column for spelling mistakes and correct them using the Find and Replace method until all city names are spelled correctly.
16. Check Numerical Columns Stored as Text
Check the numerical columns for numbers stored as text. If any numbers are stored as text, convert them into numbers using the Number section on the Home tab.
17. Check the Age Column for Invalid Ages
Check the Age column for invalid ages, such as ages below 18 or above 65.

If an employee's age is below 18 or above 65, remove that employee's record.

Apply a filter to the Age column, select only the invalid ages, and delete the visible rows using Ctrl + -.
18. Check the Salary Column for Invalid Salaries
Check the Salary column for negative salaries, zero salaries, and unusually high salaries, such as salaries in crores.

Apply a filter, select the invalid salaries, and delete the visible rows using Ctrl + -.
19. Check Department Names
Check the Department column for incorrect names, unnecessary spaces, and inconsistent naming.

For example, replace HR with Human Resource and IT with Information Technology using Ctrl + H (Find and Replace).
20. Check Hire Date and Exit Date Formats
Check the Hire Date column to ensure that all dates are in the proper DD-MM-YYYY format.

Apply a filter to select invalid dates. Go to the Data tab and select Text to Columns. Choose Delimited, click Next, unselect all delimiters, and click Next. Select Date and choose MDY format, then click Finish.

After that, select the dates and press Ctrl + 1. Choose Custom, change the format to DD-MM-YYYY, and click OK.

Repeat the same process for the Exit Date column.
21. Clean and Reduce Job Title Names
Clean and reduce job title names as much as possible using the Find and Replace option.
22. Complete Data Cleaning
All the data cleaning is now complete, and our dataset is ready for the next step.
Step 2: Feature Engineering / Data Processing
1. Create an Employee Tenure Column
Create a column for employee tenure in years, which indicates how many years an employee has served in the company.

Use the following formulas and copy them throughout the entire column.

Formula 1:

=DATEDIF(H2,TODAY(),"Y")

This calculates tenure in completed years.

Formula 2:

=ROUND((TODAY()-H2)/365,2)

This calculates tenure in years, including decimal values.

ROUND rounds the result to two decimal places. The number 2 means two digits after the decimal point. H2 is the cell reference.
2. Create Age Groups
Create age groups to understand which age groups have the highest number of employees.

For this, use the nested IF function.

=IF(B2<30,"Junior",IF(B2<40,"Senior",IF(B2<52,"Super_Senior","top_level_manager")))
3. Create Salary Categories
Create a Salary Category column using the IF function.

For this, first calculate the monthly salary by dividing the annual salary in H2 by 12. Then create the salary category.

=IF(J2<600000,"Low",IF(J2<1200000,"Medium","High"))
4. Create an Employee Status Column
Create an Employee_Status column to check whether an employee is active or exited. Use the IF function.

=IF(I2="Active","Active","Excited")
5. Create an Employee Experience Category
Create an Employee Experience Category column using the IF function.

=IF(P2<3,"Fresher",IF(P2<6,"Senior employee","Experienced"))
6. Create a Performance Category
Create a Performance Category column using the employee rating and the IF function.

=IF(K2<=2,"Low",IF(K2<=4,"Medium","High"))
7. Create a Satisfaction Category
Create a Satisfaction Category column using the IF function.

=IF(L2<=2,"Low",IF(L2=3,"Medium","High"))
8. Calculate Annual Overtime Hours
Create a column for annual overtime hours using the following formula:

=M2*12
9. Create an Overtime Category
Create an Overtime Category column using the IF function.

=IF(X2<200,"Low",IF(X2<400,"Medium","High"))
10. Create an Absenteeism Category
Create an Absenteeism Category column using the IF function.

=IF(N2<=5,"Low",IF(N2<=10,"Medium","High"))
11. Calculate Employee Service Duration
Create a column to calculate the service duration of every employee using the DATEDIF function.

=IF(I2="",DATEDIF(H2,TODAY(),"M"),DATEDIF(H2,I2,"M"))
12. Create Hire Year and Hire Month Columns
Use the Hire Date column to extract the hire year and hire month using the following formulas.

Hire Year:

=YEAR(H2)

Hire Month:

=TEXT(H2,"mmmm")
13. Convert Formulas into Values
One important thing to remember is that after creating every new column, select the entire created column and copy it.

Then go to the Paste option in the Home tab, click the dropdown, select Paste Special, choose Values, and click OK.

Finally, paste the values into the same column from which you copied them.
Step 3: HR Data Analysis
1. Calculate the Total Number of Employees
Calculate the total number of employees using the following formula to check the total workforce in the company.

=COUNTA(A2:A9894)

COUNTA counts the cells that are not empty.
2. Calculate the Total Number of Active Employees
Calculate the total number of active employees to check how many employees are currently working in the company.

=COUNTIF(R2:R9894,"Active")

COUNTIF counts cells based on a condition.
3. Calculate the Total Number of Exited Employees
Calculate the total number of employees who have exited the company.

=COUNTIF(R2:R9894,"Exited")
4. Calculate the Average Employee Age
Calculate the average age of employees using the following formula.

=AVERAGE(B2:B9894)
5. Calculate the Average Annual Salary
Calculate the average annual salary of employees using the following formula.

=AVERAGE(J2:J9894)
6. Find the Highest and Lowest Annual Salaries
Use the following formulas to find the highest and lowest annual salaries.

Highest Salary:

=MAX(J2:J9894)

Lowest Salary:

=MIN(J2:J9894)
7. Calculate Total Annual Salary Expenditure
Calculate the total annual salary expenditure of the company using the following formula.

=SUM(J2:J9894)
8. Calculate Department-Wise Employee Counts
Calculate the number of employees in each department using the following formula.

=COUNTIF(D2:D9893,"IT")

Replace the department name each time to calculate the count for other departments.
9. Calculate Department-Wise Average Salaries
Calculate the average salary for each department using the following formula.

=AVERAGEIF(D2:D9894,"IT",J2:J9894)

Replace IT with the other department names.
10. Calculate Average Satisfaction Score by Department
Calculate the average satisfaction score for each department using the following formula.

=AVERAGEIF(D2:D9894,"IT",L2:L9894)
11. Identify Employees Receiving Salaries Above the Company Average
Identify employees whose salaries are above the company average using the following formula.

=IF(J2>AVERAGE($J$2:$J$10001),"Above Average","Below Average")
12. Calculate Employee Exit Rate
Calculate the employee exit rate using the following formula.

=COUNTIF(R2:R10001,"Exited")/COUNTA(A2:A10001)*100
13. Calculate Average Performance Rating by Department
Calculate the average performance rating for each department using the following formula.

=AVERAGEIF(D2:D9894,"IT",K2:K9894)
14. Calculate Male and Female Employee Counts
Calculate the number of male and female employees using the following formulas.

Male Employee Count:

=COUNTIF(C2:C9894,"male")

Female Employee Count:

=COUNTIF(C2:C9894,"female")
Step 4: PivotTables & PivotCharts
When creating PivotTables, make sure you select the same data source for all tables.
1. Create a PivotTable Showing Employee Count by Department
Select all the data, go to Insert, and select PivotTable.

Drop Department into Rows and Employee ID into Values.
2. Analyze Employee Distribution by Gender and Department
Select all the data, go to Insert, and select PivotTable.

Drop Department into Rows, Employee ID into Values, and Gender into Columns.
3. Analyze Average Salary by Department and Job Title
Select all the data, go to Insert, and select PivotTable.

Drop Department into Rows, Employee Salary into Values, and Job Title into Rows.
4. Analyze Employee Count by Age Group
Select all the data, go to Insert, and select PivotTable.

Drop Age Group into Rows and Employee ID into Values.
5. Analyze Active and Exited Employees by Department
Select all the data, go to Insert, and select PivotTable.

Drop Employee Status into Rows and Employee ID into Values.
6. Compare Average Performance Ratings Across Departments
Select all the data, go to Insert, and select PivotTable.

Drop Average Performance Rating into Values and Department into Rows.
7. Analyze Employee Satisfaction by Department
Select all the data, go to Insert, and select PivotTable.

Drop Satisfaction Score into Values and Department into Rows.
8. Analyze Overtime Hours by Department
Select all the data, go to Insert, and select PivotTable.

Drop Overtime Hours into Values and Department into Rows.
9. Analyze Absenteeism by Department
Select all the data, go to Insert, and select PivotTable.

Drop Absenteeism into Values and Department into Rows.
10. Analyze Employee Hiring Trends by Year and Month
Select all the data, go to Insert, and select PivotTable.

Drop Hiring Month and Year into Rows and Employee ID into Values.
11. Remove Unnecessary Elements from Charts
One important thing to remember is that after creating every chart, remove unnecessary elements such as legends, buttons, and titles if you don't need them.

To remove buttons, right-click on a button and select Hide All Field Buttons on Chart.

To remove legends and titles, click the + icon next to the chart and unselect the unwanted elements.
Step 5: HR Analysis KPIs
1. Create a KPI card showing Total Employees.
2. Create KPI cards showing Active Employees and Exited Employees.
3. Create a KPI card showing Average Annual Salary.
4. Create a KPI card showing Employee Exit Rate.
5. Create a chart showing Department-Wise Employee Distribution.
6. Create a chart showing Salary Distribution Across Departments.
7. Create a chart showing Employee Performance and Satisfaction.
8. Create an interactive dashboard using slicers for Department, Gender, Employment Type, and Work Location.
Step 6: Design the Dashboard
1. Create Shapes
Go to Insert → Shapes and select the shapes you want to use for the dashboard title and KPI cards.
2. Arrange the Shapes
Arrange the shapes according to your preferred dashboard layout.
3. Add KPI Cards
Add KPI values to the shapes and include the dashboard title and other required details.
4. Arrange the Visuals
Arrange all the visuals according to your preferred dashboard layout.
5. Add Visuals to the Dashboard
Bring the required visuals from other sheets to the dashboard one by one. Add only the visuals you need.
6. Adjust the Visuals
Adjust the visuals in the dashboard so that everything is visible in one place.
7. Create and Connect Slicers
Click on any visual, go to PivotChart Analyze, select Insert Slicer, and connect the slicers to the other PivotTables.
8. Connect Slicers
To connect slicers, right-click on the slicer and go to Report Connections.
9. Change Dashboard Colors
If you want to change the dashboard colors, go to Page Layout → Colors and select the color theme you prefer.
10. Save the File
After completing the dashboard, make sure to save the file. Don't forget to save your work.
Step 7: Insights and Recommendations
Key Insights
After analyzing the employee dataset and creating the dashboard, I found some interesting insights about the organization's workforce.
- The organization has a total of 9,893 employees, out of which 8,145 are currently active and 1,748 have exited.
- The employee exit rate is 18%, which indicates that employee retention needs attention.
- The average annual salary is approximately ₹5,47,187.
- Customer Support has the highest number of employees, with approximately 1,467 employees.
- Finance has relatively high salary expenditure compared with several other departments.
- Average employee performance ratings range from approximately 3.11 to 3.29 across departments.
- Employee satisfaction scores range from approximately 3.27 to 3.46 across departments.
- Senior employees represent the largest workforce group at 48%, followed by Junior employees at 28% and Super Senior employees at 23%. Top-level managers account for approximately 1%.
- Hiring and exit trends fluctuate over the years, showing changes in workforce growth and employee turnover.
- The dashboard also provides an overview of employee distribution based on gender, employment type, department, and work location.
Conclusion
Through this project, I was able to understand the organization's workforce in more detail. The analysis provided a clear picture of employee distribution, salary structures, performance, satisfaction, hiring trends, and employee exits.

One of the important findings was the 18% employee exit rate, which shows that employee retention is an area that needs attention. I also observed differences in employee performance and satisfaction across departments.

Overall, this project helped me understand how Excel can be used to analyze HR data, identify important patterns, and present information through an interactive dashboard to support better decision-making.
Recommendations
Based on my analysis, I would suggest a few areas where the organization can focus on improvement.
- Employee Retention: Understand why employees are leaving the organization and take steps to reduce unnecessary employee turnover.
- Employee Satisfaction: Collect regular employee feedback and improve workplace conditions, communication, and employee engagement.
- Employee Performance: Provide training and development opportunities to help employees improve their skills and performance.
- Salary Structure: Review salaries across departments and job roles to maintain fair and competitive compensation.
- Workforce Planning: Use hiring and exit trends to understand future employee requirements and plan recruitment accordingly.
- Career Growth: Provide employees with opportunities for promotions, skill development, and career advancement.
- Regular HR Analysis: Monitor employee count, exit rate, salary, performance, and satisfaction regularly to understand changes in the workforce.
What I Worked On
I worked on an Employee Workforce & HR Analytics project using Microsoft Excel. The main purpose was to analyze employee data and understand different HR-related factors such as workforce distribution, salary structures, employee performance, satisfaction, overtime, absenteeism, and employee turnover.

I divided the project into different stages.
1. Data Cleaning
I started by understanding the dataset and checking for duplicate records, missing values, incorrect data types, spelling mistakes, invalid ages, and salary values. I also standardized department names, gender, employment types, and city names to make the data consistent.
2. Data Processing & Feature Engineering
After cleaning the data, I created additional columns to make the analysis easier. These included employee tenure, age groups, salary categories, employee status, experience categories, performance categories, satisfaction categories, overtime categories, absenteeism categories, and hiring year and month.
3. Data Analysis
I used Excel formulas and functions to calculate total employees, active and exited employees, average age, average salary, highest and lowest salaries, total salary expenditure, department-wise employee count, average salary, performance, satisfaction, and employee exit rate.
4. PivotTables & PivotCharts
I created PivotTables and PivotCharts to analyze employee distribution across departments, gender, job roles, age groups, employment types, salaries, performance, satisfaction, overtime, absenteeism, and hiring trends.
5. Interactive Dashboard
Finally, I designed an interactive HR Analytics Dashboard with KPI cards, charts, and slicers for Department, Gender, Employment Type, and Work Location. This makes it easier to explore the data and understand important HR metrics.
What I Learned
This project gave me a good opportunity to improve my practical Excel and data analytics skills.

During this project, I learned:
- How to clean raw datasets and handle missing values and duplicates.
- How to use Excel formulas and functions such as IF, COUNTIF, AVERAGEIF, SUM, MAX, MIN, DATEDIF, YEAR, and TEXT.
- How to create new columns using feature engineering.
- How to analyze large datasets using PivotTables and PivotCharts.
- How to create KPI cards and interactive dashboards using slicers.
- How to identify useful patterns and trends from data.
- How to convert business requirements into meaningful analysis and visualizations.
Most importantly, I learned how to complete the entire data analysis process, starting from raw data cleaning and ending with an interactive dashboard that presents meaningful insights.
Tools & Technologies Used
Microsoft Excel
- Data Cleaning
- Excel Formulas & Functions
- Feature Engineering
- PivotTables
- PivotCharts
- Conditional Formatting
- KPI Cards
- Slicers
- Interactive Dashboard
