# NuvoRetail Assignment

## Project Overview

This project automates data loading, cleaning, dashboard updates, report generation, and emailing for e-commerce sales data analysis, leveraging PowerQuery, VBA, and Python.

### Requirements
- Automate data loading and dashboard refresh using PowerQuery and VBA.
- Use Python for production-level automation and emailing.
- Implement an alternative to Excel's substring function.

### Data Source

Dataset: Amazon e-commerce sales data, including SKU, Category, Size, Status, Fulfillment, and Sales Amount, for product optimization insights.

*Data Link:*  
[Amazon Sales Data on Kaggle](https://www.kaggle.com/datasets/thedevastator/unlock-profits-with-e-commerce-sales-data/data?select=Amazon+Sale+Report.csv)

### Data Automation with VBA

*VBA Script Overview:*
- Loads data from "AmazonSaleData" to "TransformedSaleData" worksheet, performing transformations like:
  - Filling empty cells with "Null."
  - Extracting meaningful "Status" info by removing prefixes.
  - Marking "Cancelled" rows with "Null" in currency and amount fields.

*Automatic Trigger on New Data:*  
A VBA Worksheet_Change event triggers data transformation whenever a new row is added.

### Dashboard Automation

*Metrics Visualized:*
1. *Total Sales Amount by Date* – Line chart.
2. *Order Status Distribution* – Pie chart.
3. *Sales Amount by Fulfillment Type* – Bar chart.

*Dashboard Refresh Macros:*
Each chart has a macro to refresh on demand, updating data and visuals.

### Automated Report Generation

A VBA module generates comprehensive reports by:
1. Refreshing all charts.
2. Compiling charts from dashboards into a "Report" sheet.
3. Saving the report with a timestamped filename in a "reports" folder.

### Email Automation with Python

A Python script automates report generation and emailing every hour:
1. Prompts for recipient email list, sender email, and app password.
2. Generates reports and emails them at regular intervals.

### Folder Structure and Run Instructions

Place the AmazonSaleData.csv file, Excel workbook, and Python script in the same directory. Run the Python script to automate the entire workflow.

## Automating Generating and Emailing Reports

Finally a python script has been written to link together the workflow. Below is the workflow and this run keeps generating and emailing the reports every hour. The reports are stored in the same directory as the python script and the workbook.

### Steps

1. The script first asks for the list of email ids that you want to send the report to.  
2. Then it asks for your email and password(this can be your mail password or app password shown in the next section I have used app password and gmail).  
3. It next asks for the subject and the body of the email.  
4. The script starts first checks if the dir named reports is present if not it will create the directory.  
5. Opens excel and calls the macro to generate the report mentioned in the last section.  
6. The report is generated and save under the reports directory.  
7. The script takes the latest report that was created and sends email to all the email address mentioned in step 1\.   
8. Step 5 to 7 keeps repeating every hour till we stop the script.

We can configure the frequency of sending these reports in the code itself.

#### Generating app password in gmail to make this work.

To generate an *App Password* for Gmail, follow these steps. This password is a unique 16-character code that allows less secure applications to access your Google account without using your main Google password. You’ll need this if you're using Gmail in a third-party app or service that doesn't support two-step verification directly.

##### Prerequisites

1. *Enable Two-Step Verification: You need to have **two-step verification* enabled on your Google account to use app passwords.

##### Steps to Generate a Gmail App Password

1. *Go to Your Google Account*:  
   * Open [https://myaccount.google.com/](https://myaccount.google.com/) and sign in to your Google account.  
2. *Navigate to Security*:  
   * On the left sidebar, click on *Security*.  
3. *Enable Two-Step Verification* (if not already enabled):  
   * Scroll to *"Signing in to Google"* and click on *Two-Step Verification*.  
   * Follow the instructions to set it up if you haven’t already.  
4. *Create an App Password*:  
   * Once two-step verification is enabled, go back to the *Security* page.  
   * Under *"Signing in to Google", click on **App Passwords*.  
5. *Select the App and Device*:  
   * In the App Passwords section, you’ll be prompted to choose the app and device for which you need the password.  
   * From the *Select App* dropdown, choose *Mail*.  
   * From the *Select Device* dropdown, choose the device you're generating this for, or select *Other* if you want to label it manually.  
6. *Generate the Password*:  
   * Click *Generate*. Google will provide a 16-character password.  
   * Copy this password (you’ll need it to set up your app).  
7. *Use the App Password*:  
   * Enter this app password instead of your Google password in the application you’re setting up (e.g., Outlook, Thunderbird, etc.)

## How To Run 

#### Prerequisites 

1. Latest version of python is required I am using Python 3.13  
2. Excel with macros enabled  
3. Also make sure the following packages are installed or install them using and run the script fill in all the the user inputs and you should see
```  
pip install email
pip install pywin32
python automation.py
```
