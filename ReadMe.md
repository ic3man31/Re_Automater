# Re_Automater

Version 1.0

Re_Automater v1.0 is a Python script that automates the creation of a report document from an Excel file containing vulnerability data. This script performs several tasks, including:

1. Escaping special characters in the Excel file to prevent rendering issues.
2. Creating a bar chart representing vulnerabilities by risk level and saving it as `chart.png`.
3. Filling in a Word document template with the data from the Excel file and the generated chart.
4. Modifying the layout of the Word document, including changing it to landscape orientation.
5. Removing double spaces in the document.
6. Deleting the first column (Index) from a specific table in the document.
7. Coloring cells in the "Risk" column based on their contents (Critical: dark red, High: red, Medium: orange, Low: yellow).
8. Generating a pivot table in a new Excel file to summarize vulnerabilities by risk level and creating a 3D bar chart from the pivot table

It doesn't automate everything, you still need to get your hands dirty.
*Feedbacks are always welcome.*

## Requirements

For now the script can completely run on **Windows OS**, with other OS may you encounter some issue. In the next versions I'm going to resolve this thing.
Besides the installation of [requirements.txt](requirements.txt) these are the other requirements:

- Windows OS
- Microsoft Word and Excel installed


## Installation

My suggestion is to run this script in a python virtual envinroment:
`python -m venv <namevirtualenviroment>`

Once done, to install all the packages/library you need run the following command:
`pip3 install -r requirements.txt`

## HowTo

Before running the script is important that you have the following things:

-  **.xlsx** file extension, not csv, which contains the table you want to insert into the table in [Template_Sample.docx](template/Template_Sample.docx). Specifically, the latter table has the following headers, as you can see on [Sample_Vulnerability_Excel.xlsx](template/Sample_Vulnerability_Excel.xlsx)

| Item | ID  | Risk | Name | CVE | System | Description | CVSS_Base_Score | CVSS_Temporal_Score | Solution |
| ---- | --- | ---- | ---- | --- | ------ | ----------- | --------------- | ------------------- | -------- |


If you have different headers, you can use [auto_table.py](auto_table), which will create a table in a new .docx file by entering the data from the .xlsx file you gave. Or you can edit the script to suit your needs.

- A template with **.docx** extension. This template requires **{{** example **}}** to be present so that you can replace it with the input you will give. The [Template_Sample.docx](template/Template_Sample.docx) file is just an example to show how the script works. Once you understand how it works you can modify the template or script to suit your needs.

Once you have this 2 things, you can run the script `python3 re_automater.py` .
It will ask you 10 input:

1) `Enter the path of the Excel file: ` - Here it is important that you have file with **.xlsx** extension not csv.
2) `Please enter the name of the sheet: ` - You must know the name of the sheet where your table is.
3) `Please enter the template .docx you want to use: ` - Enter the template you want to use. Please take a look on Template_Sample.docx
4) `Enter the name of the CLIENT: `
5) `Enter the name of the INFRASTRUCTURE: `
6) `Enter Start Date: `
7) `Enter End Date:  `
8) `Select the name of the final .docx output: ` Enter the name of the final output
9) `Do you want to proceed with LANDSCAPE LAYOUT for the Word Document(docx) ? (y/n): ` - This will put all the .docx file with LANDSCAPE LAYOUT
10) `Choose the path and the name of the new .xlsx: ` - This will create a new .xlsx file which contain a pivot table and a chart between Item column and Risk column

## Authors

 [J.Rosales](https://it.linkedin.com/in/johnchri-rosales31)

## License

[License](LICENSE)





