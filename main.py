import pandas as pd 
import numpy as np
import openpyxl
import xlwings as xw
from fuzzywuzzy import fuzz, process
from Levenshtein import distance
from unidecode import unidecode

# files
file_ventas = "C:/Users/shant/OneDrive/Documentos/Ventas_Historicas_Dental.xlsx" 
file = "C:/Users/shant/Dropbox/Admin/Pacientes_DB.xlsx"
pacientes = pd.read_excel(file)
ventas = pd.read_excel(file_ventas)

# Convert to dataframe, extract the column and convert values as string
px_df = pd.DataFrame()
px_df['Name'] = ventas['Nombre']
px_df['Name'].astype('string')
px_df['Name'] = px_df['Name'].str.title()

# function to remove strings or phrases
def remove_string(df, column, string):
    df[column] = df[column].str.replace(string, "")

remove_string(px_df, 'Name', "Se realizó HC, valoración. Consulta")  

# dropna
px_df.dropna(subset=['Name'],inplace=True)


remove_string(px_df, 'Name', " ")
remove_string(px_df, 'Name', ".")
  
# use unidecode to remove accents
px_df['Name'] = px_df['Name'].apply(lambda x: unidecode(x))

# create a list from the columns
px_list_names = list(px_df['Name'])

# using levenshtein 
def levenshtein_distance(s1, s2):
    return distance(s1, s2)


# create a new list empty 
px_new_names = []

# for loop that iterates to exctract the most similar name from original list and append it to the empty list 
for name in px_list_names:
    name_str = name.title()
    new_name = process.extractOne(name_str, px_list_names)[0]
    if new_name not in px_new_names:
        px_new_names.append(new_name)
    else:
        pass

# for loop that reformat the names with a space
formatted_names_list = []

for name in px_new_names:
    formatted_name = ""
    for i in range(len(name)):
        if i > 0 and name[i].isupper():
            formatted_name += " "
        formatted_name += name[i]
    print(formatted_name)
    formatted_names_list.append(formatted_name)


# compare the name in sales and replace it with the corrected name 

def get_closest_match(name, name_list):
    closest_match, score = process.extractOne(name, name_list)
    return closest_match

# closest_match, score = process.extractOne('SHant Gzm', formatted_names_list)
# print(closest_match, score)

ventas['Nombre'] = ventas['Nombre'].str.title()
ventas['Nombre'] = ventas['Nombre'].astype(str)
ventas['Name_Corrected'] = ventas['Nombre'].apply(lambda x: get_closest_match(x, formatted_names_list))


create_output(ventas)


similar_pairs_100 = []
similar_pais_gt_95 = []
similar_pais_gt_90 = []
similar_pairs_lt_90 = []

for i in range(len(formatted_names_list)):
    for j in range(i+1, len(formatted_names_list)):
        name1 = formatted_names_list[i]
        name2 = formatted_names_list[j]
        ratio = fuzz.partial_ratio(name1, name2)
        # identical pairs
        if ratio == 100:
            similar_pairs_100.append((name1, name2))
        elif ratio >= 95:
            similar_pais_gt_95.append((name1, name2))
        elif ratio >= 90:
            similar_pais_gt_90.append((name1, name2))
        else:
            similar_pairs_lt_90.append((name1, name2))


print("Similar pairs:", similar_pairs)


removing_names = [
    'Carlos Perez Ramale',
    'Nohemy Adriana Rojas', 
    'Giselle Villegas',
    'Nohemi Chavez Hernandez',
    'Juan Liebre', 
]


final_list_names = set(formatted_names_list) - set(removing_names)
print(len(final_list_names))
list(final_list_names)

# create new dataframe with formatted patient names and xlwings
def create_output(list_):
    df = pd.DataFrame(list_)
    book = xw.Book()
    sheet = book.sheets[0]
    sheet.range('A1').value = df
create_output(formatted_names_list)

#create_output(formatted_names_list, 'Cleaned_Names')
# create_output(similar_pairs_lt_90, 'New_name')
create_output(similar_pairs_100)
create_output(similar_pais_gt_90)
create_output(similar_pais_gt_95)




w1 = "Luisa Espindola Rodriguez"
w2 = 'Luisa Espinola Rodriguez'
fuzz_ratio = fuzz.ratio(w1, w2)
fuzz_partial_ratio = fuzz.partial_ratio(w1, w2)

print(f"Fuzzy Ratio is {fuzz_ratio}")
print(f"Fuzzy Partial Ratio is {fuzz_partial_ratio}")