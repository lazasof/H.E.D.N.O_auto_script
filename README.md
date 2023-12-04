#HEDNO Excel Files Processing Script  
This script is designed specifically to handle HEDNO (Hellenic Electricity Distribution Network Operator) Excel files. It performs various operations to prepare the data for further analysis.  
  
Functionality:  
Removes the first sheet of the Excel file.  
Utilizes user input to select specific rows for data processing.  
checks for user input mistakes.  
check for special cases(eg. input -1 brings back all fields)  
saves the variables as column names along with their units of measurements.    
Concatenates all remaining sheets into a single sheet, creating a comprehensive dataset.    
  
Requirements:  
Python 3.x
Pandas library
OpenPyXL library
pyinstaller(is used to bake the script and make it usable without Python)
  
Clone the Repository:
  
```bash
git clone https://github.com/your/repository.git
```
```Install Dependencies:
!pip install pandas 
!pip install openpyxl
!pip install pyinstaller 
```
Run the Script:  
```python
python hednoscript.py  
```
Instructions
Input Excel File:
Provide the path to the HEDNO Excel file when prompted.  
  
Row Selection:  
Enter the row numbers you want to include in the concatenated dataset with space as delimiter. For instance: 2 5 8 to include rows 2, 5, 8.  
    
Processing:  
The script will remove the first sheet,check for mistakes and special cases(eg. inputing -1 brings back all fields), process the specified rows, and concatenate all remaining sheets into a single dataset.
  
Output:
The resulting processed data will be stored in a new Excel file for further analysis.
  
Example:
```python
$ python hedno_excel_processing.py

Enter the path to the HEDNO Excel file: /path/to/your/file.xlsx
Enter row numbers to include (e.g., 2, 5, 8-10): 2, 4-7

Processing...

Resultant dataset saved as 'combined_sheets.xlsx'
```  
Notes:  
Ensure the specified row numbers are within the range of each sheet in the Excel file.  
Make sure the Excel file follows the HEDNO format for proper processing.  
