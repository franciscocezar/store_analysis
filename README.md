# Indicators Automation

### Objective: Train and create a complete project involving the automation of a process done in the computer

### Description:

Imagine you work in a large chain of clothing stores with 25 stores spread all over Brazil.

Every day, in the morning, the data analysis team calculates the so-called OnePages and sends them to the manager of each store, along with all the information used to calculate the indicators.

A OnePage is a very simple and straight to the point summary, used by the store management team to know the main indicators of each store and allow in 1 page (hence the name OnePage) both the comparison between different stores, and which indicators that store managed to meet that day or not.

Your role, as a Data Analyst, is to be able to create a process as automatically as possible to calculate each store's OnePage and send an email to each store's manager with their OnePage in the body of the email and also the complete file with their respective store's data attached.

### Important files and information

- Emails.xlsx file with the name, store, and email of each manager.   
Note: I suggest replacing the email column for each manager with your own email, so you can test the result.

- Sales.xlsx file with the sales of all stores.  
Note: Each manager should only receive the OnePage file and an attached Excel file with his or her store's sales. Information from another store should not be sent to the manager that is not from that store.

- Stores.csv file with the name of each store.

- At the end, you should also send an email to the board of directors (the information is also in the Emails.xlsx file) with 2 store ratings attached, 1 rating of the day and another rating of the year. Also, in the body of the email, you should highlight which was the best and worst store of the day and also the best and worst store of the year. The ranking of a store is given by the revenue of the store.

- The spreadsheets for each store must be saved inside the store's folder with the date of the spreadsheet, in order to create a backup history.

### OnePage Indicators

- Billings -> Target Year: 1,650,000 / Target Day: 1000
- Product diversity (how many different products were sold in this period) -> Target Year: 120 / Target Day: 4
- Average ticket per sale -> Target Year: 500 / Target Day: 500

Note: Each indicator should be calculated on the day and year. The day indicator must be the last available day in the sales spreadsheet (the most recent date).

Note: Tip for the green and red characters: take the character from this site (https://fsymbols.com/keyboard/windows/alt-codes/list/) and format it with html.

# Credits

This is an exercise from the `Hashtag Python Course`

website: [Portal | Hashtag](https://portalhashtag.com/login)  
YouTube channel: [Hashtag Programação](https://www.youtube.com/channel/UCafFexaRoRylOKdzGBU6Pgg)

Check it out if you want to learn tecnologies like **Python**, **SQL**, **Power BI**, **Data Science**, **Excel**, and other **amazing courses**.
