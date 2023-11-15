# WOC_Automation
Letting engineers know about their **WOC and WO details** on every Monday


Prerequisite’s 
===============

1) Make sure Google Sheet contains the data as per below format [ Date, Days, Email ]

![Screenshot 2023-11-15 at 10 34 57 PM](https://github.com/Surenu1248/WOC_Automation/assets/31179719/483eb6b0-b55a-49e0-bd03-2d9ea620ad6d)


2) Open the AppScript and create a new file with '.gs' exetension. Copy paste the 'Code.gs' file code and save it [ attached to the main branch ]
   
   a) Provide your sheet name on line no: 3
      ```
     "var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WOC");"
      ```
   
4) Similary create a new file with '.html' extension. Copy paste the 'EmailOutput.html' file code and save it [ attached to the main branch ]


#Scenario 1

Trigger this on every **Monday**

1) Make sure the line 95 should contain the following code

```
   var WOCDateMinus5 = new Date(WOCDate.getTime()-5*(24*3600*1000));
```

2) Similarly line 100 should be

```
   if(todayDate == WOCDateMinus5 && dayInNumber == 1)
```

3) Deploy the changes.
 
4) Now Select 'Triggers' [ left side 4th option ]

5) Click on 'Add Trigger' and fill the details as per below and save it.

   ![Screenshot 2023-11-16 at 12 54 53 AM](https://github.com/Surenu1248/WOC_Automation/assets/31179719/a443c36a-e11d-49f4-8264-a46a3f8fdd4c)



#Test Output

<img width="783" alt="Screenshot 2023-11-16 at 3 55 37 AM" src="https://github.com/Surenu1248/WOC_Automation/assets/31179719/73542cc0-f5e1-4c16-9dd2-a6443f0f1a99">





