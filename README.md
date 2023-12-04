#**HEDNO Excel Files Processing Script**  <br>
This script is designed specifically to handle HEDNO (Hellenic Electricity Distribution Network Operator) Excel files. It performs various operations to prepare the data for further analysis.  <br>
  <br>
**Functionality:**  <br>
Removes the first sheet of the Excel file. <br> 
Utilizes user input to select specific rows for data processing.  <br>
checks for user input mistakes.  <br>
check for special cases(eg. input -1 brings back all fields)  <br>
saves the variables as column names along with their units of measurements. <br>   
Concatenates all remaining sheets into a single sheet, creating a comprehensive dataset.<br>    
  <br>
**Requirements:**  <br>
Python 3.x<br>
Pandas library<br>
OpenPyXL library<br>
pyinstaller(is used to bake the script and make it usable without Python)<br>
  <br>
**Clone the Repository:**  <br>
  <br>
```bash
git clone https://github.com/lazasof/H.E.D.N.O_auto_script.git
```
<br>

**Install Dependencies:**    
```bash
!pip install pandas 
!pip install openpyxl
!pip install pyinstaller 
```
<br>

**Run the Script:**  
```bash
python HEDNOscript.py  
```
<br>

**Instructions**
**Input Excel File:**
Provide the path to the HEDNO Excel file when prompted.  
  <br>
  
**Row Selection:**  <br>
Enter the row numbers you want to include in the concatenated dataset with space as delimiter. For instance: 2 5 8 to include rows 2, 5, 8.<br>  
 <br>   
 
**Processing:** <br>  
The script will remove the first sheet,check for mistakes and special cases(eg. inputing -1 brings back all fields), process the specified rows, and concatenate all remaining sheets into a single dataset.<br>
  <br>
  
**Output:** <br>
The resulting processed data will be stored in a new Excel file for further analysis.<br>
  <br>
  
**Example:**
```bash
$ python HEDNOscript.py

Enter the path to the HEDNO Excel file: /path/to/your/file.xlsx
Enter row numbers to include (e.g., 2, 5, 8): 2 5 8

Processing...

Result dataset saved as 'combined_sheets.xlsx'
```
<br>

**Notes:**  <br>
Ensure the specified row numbers are within the range of each sheet in the Excel file.  <br>
Make sure the Excel file follows the HEDNO format for proper processing.  <br>
