# CSV to Excel Converter
Python based desktop application for Windows which converts CSV files into an excel sheet

This is a simple software application which converts files in the CSV format into an user-friendly, readable excel workbook. It utilizes the Pandas, xlsxwriter and Tkinter libraries to help users select the CSV file they need to convert, modify field (column) names and store the final excel workbook in the required folder/directory.

Prerequisites:
-Python 3.10 or higher
-Microsoft Excel
-A python IDE like Jetbrains Pycharm or Microsoft Visual Studio (optional)
-python package installer - pip will be available by default, other options like Anaconda are also available

Installation steps- using Windows command prompt and Python IDLE:
- Choose the path/directory where you want this application to be installed. Create a virtual environment:
- Open Windows command prompt: Use the cd command to move to the required directory- for example: 'cd C:\Program Files'
- Now create a virtual environment using this command: python -m -venv project_name (Replace project_name with any name of your choice) 
- Once the virtual environment is created, check if a new folder 'project_name' exists using the dir command, or simply navigate to the directory using the file manager as usual
- On the command prompt terminal- use the cd command to move into the Scripts folder: Example - cd project_name/Scripts
- Activate the virtual environment: Run the activate.bat file for windows - project_name\Scripts\activate
- Now, install Pandas, XlsxWriter and Tkinter
- pip install pandas
- pip install xlsxwriter
- pip install tkinter
- Once all three libraries are installed successfully, copy the python scripts from this repository into your project_name folder, and create three folders within this path, named file_source,file_merge and file_dest
- To specify the list of column names to be used, change the field names in CSVconverter2.py as shown below:
- In class userfilechoice() within the ctafheader(self) method, add a line 'reportlist.append('FIELD NAME,'), where FIELD_NAME represents a column heading for the excel sheet
- Add or remove as many fields as required

Steps to run the application on Python IDLE:
- Open python IDLE on your desktop and open the guitest.py script from your project folder. Select the run->Run module option or press F5. A window will pop up with the usual prompts and listboxes, asking for the input CSV file. Multiple files can be selected. Please ensure that only txt files are selected as the input.

Input CSV file: Sample format:

Line 1: Header - YYYYMMDDHHSS FILE_NAME SHORT ONE LINE DESCRIPTION OF REQUIREMENTS in CSV format
Lines 2 - n: CSV records- field1,field2,.....,fieldn
Line n + 1: Trailer- YYYYMMDDHHSS, total no of records, file name, in CSV format

Output file- excel sheet:

Line 1: Header - YYYYMMDDHHSS FILE_NAME SHORT ONE LINE DESCRIPTION OF REQUIREMENTS in CSV format - removed
Line 2: FIELD 1, FIELD 2, SUBFIELD 1, FIELD 3,SUBFIELD 4,......,FIELD N - From CSVconverter2.py
Line 3 onwards: Data from input CSV file. (Trailer will be removed)

Caution: It is better to create the virtual environment using IDEs like Pycharm

Sources:
GeeksforGeeks: 

https://www.geeksforgeeks.org/how-to-remove-blank-lines-from-a-txt-file-in-python/

https://www.geeksforgeeks.org/remove-last-character-from-the-file-in-python/

Pandas: https://pandas.pydata.org/docs/index.html

XlsxWriter: https://xlsxwriter.readthedocs.io/

Tkinter: https://docs.python.org/3/library/tk.html




