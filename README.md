# pythonMiniproject

## Introduction :
This miniproject is based on automated excelsheets where we can retrive and extract a particular data from a large group of files into a master sheet using python automation.
our main aim is to make data entry job easy and We are going to write a Python program that can process thousands of spreadsheets and manage all the calculations under a second for the user.

## Folder Structure

Folder             | Description
-------------------| -----------------------------------------
`1_Requirements`   | Documents detailing requirements and research
`2_Design`         | Documents specifying design details
`3_Implementation` | All code and documentation
`4_Test_plan`      | Documents with test plans and procedures

## About the project
The aim of the project is to extract the data present in different spreadsheets in one excel file as required by the user by different paths given by him. The excel sheet scrolls through all the spreadsheets with the following data common in all the sheets:

  * `Name :`
  * `Ps Number : `
  * `Email id : `
  
The user defines the data that needs to be searched on the basis of the common data. The python program then reads the data corresponding to the particular data from different spreadsheets of excel. It then creates a mastersheet and adds the data from all the sheets to it. In the end, the data to be provided to the user is printed to the console.
## Structure to Run the program present in 3_Implementation/src/source.py
*	STEP 1 : Run the source.py file
*	STEP 2 : Enter your multiple Workbook Path Present in 3_Implementation/data.xlsx/ file 
*	STEP 3 : If entered data is invalid :: Do you want enter data again Y/N: (# If yes enter multiple path if no go to step 4)
*	STEP 4 : Do you want to: A) Search By PS Number : B) Search By User Name : [A/B]? : (Enter Details)
*	STEP 5 : Open the Masterfile present at src folder to get output data.



**SLNo** | **Library name** | **Operation** | **Install Library Code**
------ | ------------------ |---------------| ------
1 |  `Openpyxl`   | Reading and writing excel sheet | pip install openpyxl
2 | `Pandas`     | To automate excel sheet | pip install pandas
3 | `OS` | searching the File path | pip install os_sys

