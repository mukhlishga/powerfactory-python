# powerfactory-python

DIgSILENT PowerFactory is a power system analysis simulation software that is commonly used by electrical engineer to conduct various electrical power system analysis. One common use case is for power flow/load flow calculation of a power grid. One advantage of using PowerFactory is that this program can be embedded with python scripts, so we can implement various automated operations in our power system simulation. We can also generate thousands of power flow calculation to make a power flow dataset that can be utilized for machine learning task.

Please first set up the PowerFactory configuration so that it can be embedded with python scripts. You can find the instruction in the PowerFactory user manual 2020 page 359. The python scripts in this repo can be inserted to the python scripts file.

Please first install openpyxl, a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files, so we can store the data generated by the python scripts to an excel file. This excel file can be used for other uses.

To install openpyxl you can type the following line to your terminal:
`py -m pip install openpyxl`

Unless you state where you want to put your output file in the code, the default generated excel file can be found in this location:
C:\Users\{your pc name}\AppData\Local\Temp\PowerFactory.MrfN6dVP
