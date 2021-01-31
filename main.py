# PREVIOUS VS CURRENT AUTOMATION

# Packages:
import os
import glob
import pandas as pd
import numpy as np
from datetime import datetime
import re

print('''
      # PREVIOUS VS CURRENT SNAPSHOT AUTOMATION #
''')

# Get script's address
script_directory = os.getcwd().replace("\\", "/")
print('Directory: ' + script_directory)

# Get .xls file from 'dir'
os.chdir(script_directory)

try:
    filename = glob.glob('*.xls')[0]  # Think of how to throw an error if no xls found
    print('Found .xls file: ' + filename)
except FileNotFoundError:
    print('\n.xls file is not found\n')


# Read file and create df using pandas:
df = pd.read_excel(filename, header=None)

# Save two first cells in first column as title
reportTitle = list(df.iloc[0:2, 0])
reportTitle = ", ".join(reportTitle)

# Get index of the first row in the dataframe (look for first non-Nan in column 1)
row_start = df[1].first_valid_index()

df.columns = df.iloc[row_start]
df = df[row_start+1:].reset_index(drop=True)

# Drop Status column. Errors='ignore' means ignore if column with such name does not exist. Inplace means overwrite
df.drop('Status', axis=1, errors='ignore', inplace=True)

# Removing rows that contain certain substrings in Product column (if remove_products.csv is provided):
try:
    remove = pd.read_csv('remove_products.csv', delimiter=',', header=None)
    print('Found remove_products.csv file' )
    remove = list(remove[0])
    df = df[~df.Product.str.contains('|'.join(remove))]
except FileNotFoundError:
    print('\nremove_products.csv has not been provided\n')

# Check for NaNs and replace with 0
df.isna().sum().sum()
df = df.fillna(0)


# I'm splitting the dataframe into 4: current, previous, %diff and actual diff. First, getting column names as list:
cols = df.columns.tolist()

# Column names that don't contain digits will be shared by both dataframes:
main_colnames = [colname for colname in cols if not any(char.isdigit() for char in colname)]


# Function to select columns containing substring and returning colnames as a list
def select_cols(substring):
    result = [colname for colname in cols if substring in colname]
    return result

# Apply function to select columns related to curr, prev:
curr_colnames = select_cols('curr')
prev_colnames = select_cols('prev')

curr = df[main_colnames + curr_colnames]
prev = df[main_colnames + prev_colnames]


curr.columns = main_colnames+[int(col.replace('curr', '')) for col in curr_colnames]
curr = curr.melt(main_colnames, value_name='Current')

prev.columns = main_colnames+[int(col.replace('prev', '')) for col in prev_colnames]
prev = prev.melt(main_colnames, value_name='Previous')

df_joined = curr.merge(prev, how="left")
df_joined['Actual Difference'] = df_joined.Current - df_joined.Previous
df_joined['% Difference'] = 100 * ((df_joined.Current/df_joined.Previous)-1)

df_joined['% Difference'].fillna(0, inplace=True)
df_joined.rename(columns={'variable': 'Year'}, inplace=True)

# Asking user to enter current year and break if not found:
def get_curr_year(prompt):
    try:
        current_year = int(input(prompt))
    except ValueError:
        print("Does not seem like an integer. Try again.")
        return get_curr_year(prompt)
    
    if current_year not in df_joined['Year'].unique():
        min_y = min(df_joined['Year'].unique())
        max_y = max(df_joined['Year'].unique())
        print(f"File contains years between {min_y} and {max_y}. Try again.")
        return get_curr_year(prompt)
    else:
        return current_year

current_year = get_curr_year('Please enter current year: ')
print('\nCreating DataFrame object...')

df_joined['Period'] = np.where(df_joined.Year <= current_year, 'Historic', 'Forecast')
df_joined['Period'] = np.where(df_joined.Year == current_year, 'Current', df_joined['Period'])


grouping = ['Region', 'Country', 'Sector', 'Product', 'Data type', 'Unit']
df_joined.sort_values(grouping, inplace=True)

df_hist = df_joined[df_joined['Period']=='Historic'].reset_index(drop=True)
df_curr = df_joined[df_joined['Period']=='Current'].reset_index(drop=True)
df_forecast = df_joined[df_joined['Period']=='Forecast'].reset_index(drop=True)


hist_ch = df_hist.groupby(grouping, as_index=False)['% Difference'].apply(lambda x: x.abs().max())\
    .rename(columns={'% Difference': 'Historic years max abs %diff'})
curr_ch= df_curr.groupby(grouping, as_index=False)['% Difference'].apply(lambda x: x.abs().max())\
    .rename(columns={'% Difference': 'Current year max abs %diff'})
forecast_ch = df_forecast.groupby(grouping, as_index=False)['% Difference'].apply(lambda x: x.abs().max())\
    .rename(columns={'% Difference': 'Forecast max abs %diff'})


# Splitting and pivoting back to wide
id_vars = main_colnames + ['Year']
datatype_cols = list(df_joined.columns.difference(id_vars+['Period']))

dfs = pd.DataFrame()
for col in datatype_cols:
    table = df_joined[id_vars + [col]]\
        .melt(id_vars, var_name='dtype')\
        .pivot(index=main_colnames+['dtype'], columns='Year', values='value').reset_index()
    dfs = dfs.append(table)

output = dfs.merge(hist_ch).merge(curr_ch).merge(forecast_ch)\
    .sort_values(['Region', 'Country', 'Sector', 'Product', 'Data type', 'dtype']).reset_index(drop=True)

# Combining Data type and Unit just to make it look nicer. Dropping Unit col
output['Data type'] = output['Data type']+', '+output['Unit']
output.drop('Unit', axis=1, inplace=True)

# Sorting dtype according to a custom list:
order = ['Current', 'Previous', 'Actual Difference', '% Difference']
output['dtype'] = pd.Categorical(output['dtype'], order, ordered=True)

#output.to_excel('output.xlsx', index=False)

global_output = output[(output['Region'] == output['Country']) & (output['Sector'] == output['Product'])]\
    .sort_values(['Region', 'Country', 'Sector', 'Product', 'Data type', 'dtype']).reset_index(drop=True)

country_output = output[output['Region'] != output['Country']]\
    .sort_values(['Region', 'Country', 'Sector', 'Product', 'Data type', 'dtype']).reset_index(drop=True)

country_output['Data Entry Level'] = np.where(country_output['Sector'] == country_output['Product'],
                                              'Category level', 'Data entry level')

# Reordering columns: place before 7th index
col_order = list(country_output.columns)
col_order.insert(7, col_order[-1])
col_order.pop(-1)
country_output = country_output[col_order]


print('\nDataFrame generated. Creating Excel workbook...\n')

#######################################################################################################################
# Writing Excel workbook

proj = "-".join([x.strip() for x in list(output["Sub-project"].unique()
                                     if len(output["Sub-project"].unique()) > 1
                                     else output["Sub-project"].unique()[0])])
timestamp = datetime.now().strftime("%d %b %Y")
title = str(proj + " Key Data Revisions - " + timestamp)
file = str(title + ".xlsx")

writer = pd.ExcelWriter(file, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
global_output.to_excel(writer, sheet_name='Global and Regional', startrow=4, index=False, header=True)
country_output .to_excel(writer, sheet_name='By Country', startrow=4, index=False, header=True)

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
global_sheet = writer.sheets['Global and Regional']
country_sheet = writer.sheets['By Country']

cell_formatA1 = workbook.add_format({'bold': True, 'font_size': 14})
cell_formatA2 = workbook.add_format({'bold': True, 'font_size': 12})

for sheet in [global_sheet, country_sheet]:
    sheet.write('A1', title, cell_formatA1)
    sheet.write('A2', reportTitle, cell_formatA2)


global_sheet.autofilter(4, 0, 4 + global_output.shape[0], global_output.shape[1]-1)
country_sheet.autofilter(4, 0, 4 + country_output.shape[0], country_output.shape[1]-1)


text_format = workbook.add_format({'text_wrap': True})

header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'vcenter',
    'align': 'center',
    'fg_color': '#D7E4BC',
    'border': 1
})

# Write the column headers with the defined format.
for col_num, value in enumerate(global_output.columns.values):
    global_sheet.write(4, col_num, value, header_format)

for col_num, value in enumerate(country_output.columns.values):
    country_sheet.write(4, col_num, value, header_format)

# Get indices of rows countaining '% Difference' in dtype column:
global_dtype_list = global_output['dtype'].tolist()
country_dtype_list = country_output['dtype'].tolist()

perc_rows_global = [i for i, x in enumerate(global_dtype_list) if x == '% Difference']
perc_rows_countries = [i for i, x in enumerate(country_dtype_list) if x == '% Difference']

# Percentage formatting
format_perc = workbook.add_format({'num_format': '0%',
                                   'bg_color': '#F5DF4D'})

for i in perc_rows_global:
    i += 5
    global_sheet.conditional_format(i, 7, i, 25, {'type': 'no_errors',
                                                  'format': format_perc})

for i in perc_rows_countries:
    i += 5
    country_sheet.conditional_format(i, 8, i, 26, {'type': 'no_errors',
                                                  'format': format_perc})

#global_sheet.set_column('F:F', 12, format_perc)
#country_sheet.set_column('F:F', 12, format_perc)

# Column widths:

global_sheet.set_column('A:G', 25)
global_sheet.set_column('D:G', 13)

country_sheet.set_column('A:H', 30)
country_sheet.set_column('D:G', 13)

writer.save()

print('"' + file + '" file was created. Press any button to close.')
input()

