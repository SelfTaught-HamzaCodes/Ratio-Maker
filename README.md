# Ratio-Maker

## Introduction
Ratio-Maker is a versatile software tool designed to simplify working with data containing ratios. Whether you're dealing with packing lists, inventory management, or any scenario involving ratio-based calculations, Ratio-Maker can streamline your tasks. This user-friendly software provides a step-by-step guide to extract, manipulate, calculate, and export data efficiently. It supports both XLSX and XLS file formats. We appreciate your feedback as we continue to improve and expand the capabilities of this software.

![image](https://github.com/SelfTaught-HamzaCodes/Ratio-Maker/assets/123310424/72549e34-0d5d-40d3-bf29-28ba806bb726)

⚠ This software has been developed using **Python** and **Tkinter**. We have a plan in place to enhance and modernize it by leveraging **Flet** (Modern Front-End Framework in Python) and **AI** technologies. Stay tuned for exciting updates and improvements.
<br></br>
## Table of Contents
- [Getting Started](#getting-started).
  - [System Requirements](#system-requirements).
  - [Installation](#installation).
- [Features](#features).
- [Using Ratio-Maker](#using-ratio-maker).
  - [Step 1: Opening an Excel File](#step-1-opening-an-excel-file).
  - [Step 2: Identifying Columns](#step-2-identifying-columns).
  - [Step 3: Column Manipulation](#step-3-column-manipulation).
- [Calculating Ratios](#calculating-ratios).
  - [Step 4: Selecting Columns](#step-4-selecting-columns).
  - [Step 5: Displaying Results](#step-5-displaying-results).
  - [Step 6: Handling Multiple Items](#step-6-handling-multiple-items).
- [Exporting Data](#exporting-data).
  - [Step 7: Managing Selections](#step-7-managing-selections).
- [Feedback and Support](#feedback-and-support).
<br></br>
## Getting Started

### System Requirements

Before starting Ratio-Maker, ensure your system meets the following requirements:

- **Operating System**: Windows 8 or later.
- **Excel Format**: XLS and XLSX.
- **Memory**: 2 GB RAM or higher.
- **Processor**: 1 GHz processor or faster.
- **Hard Disk**: 90 MB of free disk space.

### Installation

Follow these steps to install Ratio-Maker on your Windows system:

#### Download

- **Download the Ratio-Maker application from [**Google Drive**](https://drive.google.com/file/d/1Ph6b1wCiY_QjUhW5ZZPeBnHos6BP1jDN/view?usp=sharing)**.

    - ⚠ Extract the compressed file in the folder of your choice, and use the exe to run the application.
    ***
- ⚠ Please note that the 'Ratio-Maker (Console Version)' application includes a console interface for reporting errors. In case of any issues or unexpected behavior, you can use the provided application executable (exe) to capture error messages and report them for assistance.
  - Download the Ratio-Maker (Console Version) from [**Google Drive**](https://drive.google.com/file/d/1UrGv5xqjaceq_-IpBvScN8W8lZJ3s53Y/view?usp=sharing).
<br></br>
## Features

- **Ratio Extraction**: Easily extract single or ranges of values from your data.
- **Column Manipulation**: Streamline data preparation by segregating combined values.
- **Efficient Calculations**: Calculate ratios for different columns with ease.
- **Multiple Item Handling**: Manage multiple items within your dataset.
- **Export Options**: Copy and paste results into Excel or other applications.
- **User-Friendly Interface**: Designed to be intuitive for users of all levels.
<br></br>
## Using Ratio-Maker
⚠ A **detailed overview** with screenshots can be seen using this PDF: [**More on Ratio Maker**](https://github.com/SelfTaught-HamzaCodes/Ratio-Maker/blob/main/More%20on%20Ratio-Maker.pdf).
### Step 1: Opening an Excel File

- Click "Open Microsoft Excel File" to open the last visited directory.
- Supported file types:
  - A. XLSX 
  - B. XLS

After opening a file, the following details will be displayed:

- 1st Paragraph: Note on how to open an Excel file while opening it.

- 2nd Paragraph: Displays File Name (Absolute Path).

- 3rd Paragraph: Displays the number of sheets followed by sheet names (in order).

- 4th Paragraph: Displays the first 5 rows x 3 columns for the first sheet.

### Step 2: Identifying Columns

- From the 3rd Paragraph, enter the sheet name where your packing list exists.
- From the 4th Paragraph, identify the row number where your columns exist and mention it in the field below row number. For the above case, the row number is 1.

Once done, press "Extract Column Names."

### Step 3: Column Manipulation

- Above Separator, display a list of all column names extracted from your packing list so you can see if the names are correct or not.

- Below Separator, an optional tool: In case your SIZE or any value is combined together, you can cross the "Segregate Box," specify the column and the delimiter (thing that separates those values), and finally press "Segregate Columns" to make additional columns.
- *Currently, it only works for SIZE (Thickness * Width) in V 1.0*.

New Columns Added - Status shows results of Segregation.

## Calculating Ratios

### Step 4: Selecting Columns

These drop-down menus will be automatically filled with the extracted columns.

- Multiple Items: Specify here if your packing list has more than one item to take out the ratio for each column.
- Calculate: Mostly this column is for Net Weight, as its sum will be calculated for all unique values in the "From."
- From: This is the value for which you would like to calculate the ratio. For example, for each Thickness (From:), calculate the Sum of Net Weight.

Press "Calculate Ratio" once done.

### Step 5: Displaying Results

This tree-view (tabular data) will display the sum of each unique value (If Calculate: Weight). "Copy Text" will copy this database, and you can paste it in an Excel file with separate columns and rows. Or you can "Import Selection":
- Single Selection, by clicking on any of the rows.
- Specific Selection, by pressing CTRL and clicking rows you want.
- Continuous Selection, by pressing a value, then pressing SHIFT and clicking on a value. This will select everything in between.

### Step 6: Handling Multiple Items

If you have Multiple Items in Step 4, then this drop-down menu will have all the items shown. "Add Range" creates a row that starts with a label for Serial No, followed by 2 Entry Fields to enter a range. The first one is "From," and the second one is "Till." NOTE: If you only need one value and not a range, you can just fill the first entry field and keep the second one empty. Limit: 20 rows. Once done, press "Get Range."

**NOTE:** Step 5 and Step 6 can be done separately or both of them can be used for precise ratio calculation.

## Exporting Data

### Step 7: Managing Selections

All selections from Step 5 and Step 6 will show in this text box. "Clear Text" can be used to clear this text widget. "Copy Text," like Step 5, can copy the entire text widget into the clipboard and can be pasted anywhere.

**IMPORTANT:** This software is under development, and your feedback is valuable in improving it.

## Feedback and Support

That’s all for now. If you have any other questions, feedback, or need support, please feel free to reach out at GitHub.

Thanks again for using Ratio-Maker.
