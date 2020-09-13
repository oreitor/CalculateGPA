# Calculate GPA

This project is related to GPA calculation with Python Pandas library using Student Information System data at Yildiz Technical University.
- Upload excel document about the courses from Student Enrollment Info on USIS.
- Rename this document to "export.xls".
- Run to [CalculateGPA](https://github.com/oreitor/CalculateGPA/blob/master/CalculateGPA.py) code.

Moreover, [EditableVersion](https://github.com/oreitor/CalculateGPA/blob/master/EditableVersion.py) can be used if do you want some changes about the courses. 
With this code, new courses can be added, remove any courses or change the grade of the selected course.
After the changes, you can calculate the new GPA and easily estimate your future GPA.

### Additional Information

##### BeautifulSoup
Beautiful Soup is a Python library for getting data out of HTML, XML, and other markup languages. It helps you pull particular content from a webpage, remove the HTML markup, and save the information. The question is why we use this module in this project, although its data is taken from excel. The reason is that, the excel table taken from the USIS system is created in HTML format. You can see that this is not a regular text file when you want to read the file or open with notepad. Therefore, we have to import Beautiful Soup module for this project. 

##### Unicode: encoding="UTF-16"
Additionally, UTF-8 is used by about 95% of all web pages. In contrast, USIS system is in the very rare UTF-16 format. That's why we prefer UTF-16 for Unicode when reading with the open() function.

# Possible Errors

### Suppressing Warning Message Using Registry Key

If there is an *"The file format and extension of "export.xls" don’t match. The file could be corrupted or unsafe."* error when opening the excel file, below are the steps to be taken.

- Press Windows + R to open the Run dialog box.
- Type regedit and press Enter to open Registry Editor.
- Go to the following location: HKEY_CURRENT_USER\Software\Microsoft\Office\X\Excel\Security
- Right-click on an empty space and select NEW > Dword (32-bit) Value.
- Name it ExtensionHardening, double-click on it and set the Base to Hexadecimal and the value to 0.
- Close the Registry Editor and restart the computer.

### Disable Protected View

To turn off the *"Protected View"* alert.

- Open Microsoft Excel > File > Options.
- Select the Trust Center tab > Trust Center Settings.
- Select Protected Views, uncheck all Protected View condition option then click Ok to save the changes.
- Restart Microsoft Excel.
