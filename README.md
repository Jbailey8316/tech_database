# tech_database

## Purpose
* Software: Micorsoft Excel, Micorsoft Access, VBA, SQL

The purpose of this project was to capture production concerns across 12 production lines and multiple departments within the business to be utilized by production support.  It allowed for data collection from mulitple workbooks to be collected and stored within a database.  That data could then be sorted and accesssed by relevant employees (techs) to be addressed as needed.  

## Results
The first initiative was to collect the needed data.  For this task I used Microsoft Excel and Access.  Excel was readily availble within the company and most employees were comfortable with using it. Access provided a platform to build a storage database for the collected data and could be linked with the Excel workbooks.

I accomplished the data collection by creating multiple standardized workbooks for the techs and production supervisors for each production line.  These workbooks were all given unique identifiers that were saved within a hidden and locked worksheet.  Using VBA sub-routines with integrated SQL these workbooks were able to communicate with the Micorsoft Access database to both store and retrieve data.  Custom forms and sheet buttons were created within Micorsoft Excel for all workbooks to simplify the entry and retrieval processes.

With this format, all production supervisors could open their respective workbooks, enter a request and update the database.

### Transaction Entry Form
![](https://github.com/Jbailey8316/tech_database/blob/main/images/TransForm.PNG)

![](https://github.com/Jbailey8316/tech_database/blob/main/images/LoadTransForm.PNG)

![](https://github.com/Jbailey8316/tech_database/blob/main/images/submitreq.PNG)

![](https://github.com/Jbailey8316/tech_database/blob/main/images/SubmitTrans.PNG)

![](https://github.com/Jbailey8316/tech_database/blob/main/images/Loadtrans.PNG)

After storing the needed data in MS Access from the production lines, the tech workbooks would retrieve and organize the data based on the unique identifiers.  The unique identifiers allowed the data to be assigned to the relevant tech for each production line.  The tech worbooks were used for tracking of all open, in-process and completed requests.  They allowed techs to assign categories to each issue which could then be utilized for cost tracking and other initiatives.  Reports could be generated and would be displayed on the dashboard of the tech workbook giving interested parties high level visibility of the collected data and its current status.

![dashboard1](https://github.com/Jbailey8316/tech_database/blob/main/images/dashboard1.PNG)

![dashboard2](https://github.com/Jbailey8316/tech_database/blob/main/images/dashboard2.PNG)

As techs updated the stored data, Access would also be modified to reflect the changes.  Upon opening or refreshing the other workbooks all data would then be pulled and updated.  This would give all connected workbooks access to live data.


## Summary
With this project implemented, the data was able to be analyzed and reports could be created and presented to management. This helped to drive accountability, implement cost reductions, and improve communication across all involved departments.  The main focus of this project was to collect and analyze bill of material (BOM) concerns related to incorrect parts, incorrect quantities, and kit structures.  It also allowed for the collection of continuos improvement projects, new model integration tracking, engineering print changes, production support requests, and employee suggestions. This setup also allowed all production lines to track open and closed issues and to follow up as needed.
