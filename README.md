# IngramMicroDatasheet

## Description

This program takes in raw inventory information from businesses and formats the information into a usable, pre-styled spreadsheet. It automatically performs calulations on the data and displays it in an excel friendly manner,

## If you are running the executable:
Make sure it is an executable (```read_file```)
Double click it

Notes:
Make sure there is a .env with the ```RAW_FILE``` name, ```NAME``` and ```FISCAL_PERIODS```
You can go into the .env to change the .txt file it reads from
You can also modify the ```FISCAL_PERIODS``` to update it for a new calendar year
It will take a second to run
If you want to remove a logo, delete it from the logos folder



## If you are trying to run the source code:

To use this program go to the terminal and type the following (first time running it)

```pip install pandas numpy python-dotenv```

Then in the terminal navigate to the read_file.py and type (a trick to doing this is right clicking the folder read_file.py is in and click the open terminal at folder:

```python3 read_file.py```


If you want to remake the .exe:

```source myapp_env/bin/activate``` 

```pyinstaller --exclude-module torch \
           --exclude-module torchvision \
           --exclude-module tensorflow \
           --exclude-module sklearn \
           --exclude-module scipy \
           --exclude-module PIL \
           --exclude-module matplotlib \
           --exclude-module transformers \
           --exclude-module IPython \
           --exclude-module jupyter \
           --exclude-module jedi \
           --exclude-module parso \
           --exclude-module pygments \
           --exclude-module fsspec \
           --exclude-module pydantic \
           --exclude-module jinja2 \
           --exclude-module regex \
           --exclude-module yt_dlp \
           --exclude-module mutagen \
           --exclude-module brotli \
           --exclude-module secretstorage \
           --exclude-module curl_cffi \
           --exclude-module certifi \
           --exclude-module urllib3 \
           --exclude-module requests \
           --exclude-module wcwidth \
           --exclude-module charset_normalizer \
           --exclude-module win32com \
           --onefile \
           read_file.py```
