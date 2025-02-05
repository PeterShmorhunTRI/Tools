# Tools

## Various Python Utilities and Scripts
___

*Please Note* These should be used internally only (For the time being). Limit uploads to only the framework for our tools, rather than the data itself. *No CADE Data should be uploaded to this repo at all.* 
___
### 1. Excel_Writer
Python script to standardize our excel ouput (both for external and internal use)
- Dependencies: `pandas 2.2.3` and later
#### Use
*Intended to be used upon export from your python environment*

**rename excel sheets to something other than the dataframe name**

`dataframes = {
    "active_srdrs": assignments #('sheet name' : df')  
}` 

**adjust global format**

`# Define common formats
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9E1F2', 'border': 1})
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    number_format = workbook.add_format({'num_format': '#,##0'})`

**Loop through sheets applying excel export format**

`for sheet_name, df in dataframes.items():
        worksheet = writer.sheets[sheet_name]`
