# NLPLineWrapper

A helper tool for [Trb2xlsx](https://github.com/LovePlusProject/Trb2xlsx), used for translations of the 3DS game [「NEWラブプラス+」/「NEWLOVEPLUS+」](https://youtu.be/Sz6p45GsLJQ?si=p1IOx1_ORt1iHWpl).  
The tool automatically processes .xlsx files and inserts a specific line-break symbol (◙) into translations at a user-defined character limit.  
This tool is specifically configured to read and format text located in Column C of your Excel spreadsheet.

# Features

- Custom Character Limits: Specify the maximum number of characters allowed per line.
- Row Range Selection: Choose exactly which rows to process, or leave the end row blank to process the entire file.
- Symbol Preservation: Choose whether to keep existing ◙ symbols (treating them as intentional hard breaks) or remove them and re-wrap the text from scratch.
- Force Wrap Long Words: Optionally force words that exceed the character limit to be broken across multiple lines.
- Automatic Backups: The original file is safely preserved and renamed with a .bac extension before saving the new changes.

# Prerequisites
Binary is available [here](https://github.com/Aikin-Ai/NLPLineWrapper/releases) for windows.  
To run this script from its source code, you need:
- Python 3.6 or higher installed on your computer.
- The openpyxl library (used for safely reading and writing Excel files).

# Installation & Setup
1. Download or save the `excel_wrapper.py` file to your computer.
2. Open your computer's terminal or command prompt.
3. Navigate to the folder where you saved the excel_wrapper.py script. For example:
```bash
cd path/to/your/folder
```
4. Install the required libraries by running the following command:
```bash
pip install -r requirements.txt
```
5. Run the script using Python:
```bash
python excel_wrapper.py
```
(Note: On some systems, especially macOS/Linux, you might need to use `python3 excel_wrapper.py`)

# How to Use the GUI
Once the application launches:
1. Select File: Click the "Browse..." button and select your target `.xlsx` translation file.
2. Set Range: Enter the Start Row and End Row (leave End Row blank to process to the bottom of the sheet).
3. Set Limit: Enter your desired Max Chars per Line.
4. Configure Settings:
  - Existing '◙': Select "Preserve" to respect manual line breaks you've already added, or "Remove" to recalculate everything.
  - Long Words: Check the box to split single words that are longer than your maximum character limit.
5. Process: Click "Process Excel File".
6. A success popup will tell you how many cells were modified. Your new file will have the exact same name as the original, and your original file will now end in `.xlsx.bac`.
