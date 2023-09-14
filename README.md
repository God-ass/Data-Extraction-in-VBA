# Data-Extraction-in-VBA

This project involves creating a VBA subroutine that automates the process of importing data from multiple CSV files into a central Excel spreadsheet. The data represents hourly output from a hypothetical bioreactor, with each CSV file containing measurements of temperature, pH, and dissolved oxygen concentration. For example, here is what one might look like, at 3:00 a.m. (in the “RalphieReactor_10-22-17_300.csv” file):

![image](https://github.com/God-ass/Data-Extraction-in-VBA/assets/92200827/6897f43b-3084-46fb-b213-18d4a5008602)


The goal of the project is to facilitate daily quality assurance tasks by providing an automated solution for data import and analysis. Instead of manually copying and pasting data from 24 hourly summary files each day, the VBA subroutine allows for efficient data extraction and consolidation into a main file. The data in cells B3:B6 (column) in each hourly summary file should be placed row-wise into the “data_extraction.xlsm” file, as shown below:

![image](https://github.com/God-ass/Data-Extraction-in-VBA/assets/92200827/f6379fb9-354e-4277-b7fb-98caf8f662c2)


The extracted data is then ready for further analysis and visualization, making it easier to monitor and maintain optimal conditions in the bioreactor.

Please note: This project assumes familiarity with VBA (Visual Basic for Applications) and its application in Excel for task automation.

## Project Structure
- Ralphie Reactor: This folder contains 24 .csv files, each representing an hourly output from the bioreactor.
- data_extraction.xlsm: This is the main Excel file where the imported data from the .csv files is consolidated.
- RalphieReactor subroutine: A VBA subroutine is created to automate the data extraction process.

## Usage
To use this project, open the data_extraction.xlsm file and run the RalphieReactor subroutine. The subroutine will prompt you to select the .csv files to import. Once the files are selected, it will automatically import the data from each file into the main Excel file.


