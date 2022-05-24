# Excel separate sheets
Batch separation of excel file pages into separate directories

## Install Python
https://www.python.org/downloads/

## Install requirements
1. Clone repository or download and extract zip
2. In terminal navigate to downloaded directory
3. Create virtual environment. In terminal run command `python3 -m venv venv` for mac and linux or `python -m venv venv`
4. Activate virtual environment, if not activated after its created. Run in terminal `venv\Scripts\activate.bat` for windows and `source venv/bin/activate` for mac or linux
5. Install requirements. Run in terminal `pip install -r requirements.txt`
6. Put your multisheet excel files into **source** directory and run script. Run in terminal `python3 main.py`
7. After a certain amount of time, the script will terminate with the output of code 0.
8. Optional use `--by-folder` param tou separate results by folders.
9. Go to **out** directory and you will be see:
  * if use option `--by-folder` - directories named as files in source directory and into the directories will be files with separated sheets
  * default all separated sheets as self files will be placed in out directory with source file prefix name
