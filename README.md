# Installation

## Windows
Get Python from [Microsoft Store](https://www.microsoft.com/pl-pl/p/python-37/9nj46sx7x90p?rtc=1&activetab=pivot:overviewtab)<br>
PIP install [here](https://www.activestate.com/resources/quick-reads/how-to-install-pip-on-windows/)

## Linux
Python install ```sudo apt install python3```<br>
PIP install ```sudo apt install python3-pip```

## Packages
```shell
cd signal-average
pip3 install -r requirements.txt
```

# Running
To process example Excel Sheets run
```shell
cd signal-average
python3 signal_average.py data/L1A1.xlsx
python3 signal_average.py data/L1A2.xlsx
python3 signal_average.py data/L2A3.xlsx
python3 signal_average.py data/L3A4.xlsx
python3 signal_average.py data/L4A5.xlsx
```