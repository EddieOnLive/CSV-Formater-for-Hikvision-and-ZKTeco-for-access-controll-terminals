# CSV-Formatter-for-Hikvision-and-ZKTeco-access-control-terminals

These two Python scripts format CSV files from access control devices into a much more readable Excel format. If you like to organize your data with one user per sheet and one file per branch just like I do, this tool is for you! :)

## How to use it:
* Use **"HikAutoReportes"** for HikVision devices.
* Use **"ZKAutoReportes"** for ZKTeco devices.

## Setup & Instructions:
1. Just place your .csv files into a `ZKentrada` folder (for ZKTeco devices) or a `Hikinput` folder (for HikVision) inside the **same directory** where you are storing the `.py` files. (you can do more than one file at a time, as long as they are from the same device brand)
2. Run the code, and it should automatically create a folder for the output files.

---

*Note: This code was tinkered with Spanish in mind. Feel free to play around with the column names and the encoding so it suits your language and needs! <3*
