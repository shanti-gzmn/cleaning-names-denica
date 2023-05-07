# Cleaning latin names database

I develop this project to help a small dental clinic to migrate from working with spreadsheets to a new ERP. The clinic worked almost two and half years using Excel Spreadsheets as databases of patients. 

The task:
The personal recorded names and sometimes there were typos, wrong grammar, acents, and so on. By the time, the archive of a treatment for a specific patient was needed, it was hard to find it because the database contained a lot of different names for one patient. Initially the database had over 1,000 patient names, after cleaning the database the total amout of real patients names is around 500. 


## Installation

Install libraries with aliases 

```bash
import pandas as pd 
import numpy as np
import openpyxl
import xlwings as xw
from fuzzywuzzy import fuzz, process
from Levenshtein import distance
from unidecode import unidecode
```
    
