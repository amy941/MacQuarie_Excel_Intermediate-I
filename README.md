# WEEK 1: 
# 🔗Link: [Week 1_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250218_Week%201)
### - Multiple Worksheets
### - 3D Formulas
### - Linking Workbooks
### - Consolidating by Positions
### - Consolidating by Reference (category)
  
💥 **- Week 1_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250218_Week%201/W1_PracticeChallenge_HeadOffice.xlsx)

💥 **- Week 1_Advanced Practice Challenge:** [adv challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250218_Week%201/W1_AdvPracticeChallenge.xlsx)

💥💥 **- Week 1_Assessment:** [assessment_Week 1](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250218_Week%201/assessment)

---

# WEEK 2
# 🔗Link: [Week 2_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250225_Week%202)
### - Combining text (CONCAT, &):
  
  *FirstName.LastName@pushpin.com* = CONCAT(C4,".",B4,"@pushpin.com") /or = C4&"."B4&"@pushpin.com"

### - Text case (UPPER, LOWER, PROPER):

  =PROPER(CONCAT(C4," ",B4))

  =LOWER(C4&"."&B4&"@pushpin.com")
  
### - Extracting text (LEFT, MID, RIGHT):

  =LEFT(**text**,[num_chars]) = LEFT(K4,2)
  
  =RIGHT(**text**,[num_chars]) = RIGHT(K4,4)

  =MID(**text**,start_num,[num_chars]) = MID(K4,4,4)

  
### - Finding text (FIND)

  = FIND(**find_text**, within_text, [start_num]) = FIND(" ", K4)-4
  
  =CONCAT(RIGHT(Inventory!F4,3),MID(Inventory!F4,FIND(",",Inventory!F4)+2,4)) --- nesting function
  
### - Date calculation (DATE, NOW, TODAY, YEARFRAC)

  = YEARFRAC(start_date, end_date) = YEARFRAC(F4,TODAY())
  
💥 **- Week 2_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250225_Week%202/C2-W2-Practice-Challenge.xlsx)

💥💥 **- Week 2_Assessment:** [assessment_Week 2](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250225_Week%202/C2-W2-Assessment-Workbook.xlsx)

---

# WEEK 3
# 🔗Link: [Week 3_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250228_Week%203)
### - Names Ranges:
  
  =N4*Pension_Rate
  
### - Create and manage ranges
  
  =AVERAGE(Annual_Salary)
  
  =MIN(Next_Review)
  
  =MAX(Date_of_Hire)
  
### - Apply ranges to formulas
  
💥 **- Week 3_Practice Challenge:** None🚫

💥💥 **- Week 3_Assessment:** [assessment_Week 3](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250228_Week%203/C2-W3-Assessment-Workbook.xlsx)

---

# WEEK 4
# 🔗Link: [Week 4_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250301_Week%204)
### - COUNT funct.

- **COUNT():** only counts the number of occurrences of cells that contain **NUMERIC** values. However, COUNT() still works on **dates**.

- **COUNTA()** only counts the number of cells that contain **alpha/numeric/alphanumeric** values.

- **COUNTBLANK():** only counts the number of occurrences of **blank** cells.

### - Counting w Criteria (COUNTIFs)

- COUNTIFS(**criteria_range1,** criteria1,...) = COUNTIFS(**State,** A5)
  
- Others:

  =COUNTIFS(**State,** A5)
  
  =COUNTIFS(**Order_Year,** 2013)
  
  =COUNTIFS(**Order_Priority,** "High") ⚠️ place a string in " "
  
  =COUNTIFS(**Order_Quantity** ">40") ⚠️ place a condition in " "
 
### - Adding w Criteria (SUMIFs)

- SUMIFS(sum_range, criteria_range1, **criteria1**, [criteria_range2, criteria2], [criteria_range3, criteria3]...)
  
  = SUMIFS(Total, Account_Manager, A21)
  
  = SUMIFS(Total, Account_Manager, **$A21**, Order_Year, **C$20**)
  
### - Sparklines

- Both **Row data** and **Column Data** can be used to create a single Sparkline
  
-  Features can be highlighted on a sparkline: high point, low point, first point, last point, markers, negative points


### - Advanced Charting

- Switching row & column
- Selecting data,
- Changing chart type
- Adding a secondary axis

  
### - Trendlines
  
💥 **- Week 4_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250301_Week%204/C2-W4-Practice-Challenge.xlsx)

💥💥 **- Week 4_Assessment:** [assessment_Week 4](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250301_Week%204/C2-W4-Assessment-Workbook.xlsx)

---

# WEEK 5
# 🔗Link: [Week 5_folder]()
### - Create and format tables

- 
### - Sort and filter in tables
### - Automation
### - Subtotalling

  
💥 **- Week 5_Practice Challenge:** [challenge]()

💥💥 **- Week 5_Assessment:** [assessment_Week 5]()

---

# WEEK 6
# 🔗Link: [Week 6_folder]()
### - Combining text
### - Text case
### - Extracting text
### - Finding text
### - Date calculation
  
💥 **- Week 6_Practice Challenge:** [challenge]()

💥💥 **- Week 6_Assessment:** [assessment_Week 6]()







