# Test-Repository


![SQLiteDemo](https://github.com/user-attachments/assets/ac6d0056-64f2-4105-9284-2602e0a92c50)


Working with massive DataFrames in a standard IDE can be cumbersome and restrictive. [InteractiveDataViewer] bridges the gap between Python's analytical power and Excel's intuitive interface. By leveraging the [xlwings](https://www.xlwings.org/) library, this tool allows you to seamlessly push, manipulate, and explore your data in a familiar spreadsheet environment without losing your Python workflow.

### 🚀 Getting Started
#### Prerequisites
-  [Install the XLWings Library](https://www.xlwings.org/)
-  Install the Pandas and numpy Python libraries

#### Configration
- Clone the repostitory to a local drive
- After installing the XLWings library, an addtional xlwings tab should be available in Excel. Click on this tab and configure the following
  <img width="850" height="109" alt="xlwings_config" src="https://github.com/user-attachments/assets/767a0c28-0504-496d-acfa-1f1e35847ad7" />
   -  Interpreter: Full path to your Python executable (include Python.exe at the end)
   -  PYTHONPATH: Path to the src directory of the cloned repository on the local drive
   -  UDF Modules: interactive_data_vierwer_udf


### 💻 Usage
The UDF contains the following Excel dynamic array functions. Please refer to the WIKI page for more detailed information for how to use these functions.

-  ***IDVTableFromFile***
   This is the core function of the UDF that reads the contents from a file and returns the result as a table in Excel. It can currently read sqlite, parquet, Python pickle and Excel files (csv, xlsx, xls). It takes the following arguments:
   - **[FileInfo](https://github.com/gerwinkooij/Test-Repository/wiki/File-Parameters)**: Range holding information about where the file is stored and if required, what table name or attribute that holds the data to return
   - **[FilterParameters](https://github.com/gerwinkooij/Test-Repository/wiki/Filter-Parameters)**: Range containing instructions for how to filter on what rows of the data to return
   - **[TransformationParameters](https://github.com/gerwinkooij/Test-Repository/wiki/Transformation-Parameters)**: Range holding information for how to transform the data (e.g. how to aggregate, which columns to include, adding derived columns, etc.)

- ***IDVListFileColums***
   This function will read the file contents and display which columns are available. It takes the following arguments:
   - **[FileInfo](https://github.com/gerwinkooij/Test-Repository/wiki/File-Parameters)**: Range holding information about where the file is stored and if required, what table name or attribute that holds the data to return
   - **MatchString (Optional)**: Only include columns that contain the specified match str.
     
- ***IDVTableFromRange***
   This function reads the contents of the range. It takes the following arguments:
   - **Range**: Top left cell of the Excel Range that should be read as a table.
   - **[FilterParameters](https://github.com/gerwinkooij/Test-Repository/wiki/Filter-Parameters)**: Range containing instructions for how to filter on what rows of the data to return
   - **[TransformationParameters](https://github.com/gerwinkooij/Test-Repository/wiki/Transformation-Parameters)**: Range holding information for how to transform the data (e.g. how to aggregate, which columns to include, adding derived columns, etc.)
 
- ***IDVListTransformationParameters***
  Returns a list of parameters available to transform the result set, their default values and a brief explanation for how to use them. 

- ***IDVListDataFileParameters***
  Returns a list of parameters for specifying the details of the source data file.

- ***IDVListAggregationParameters***
  Returns a list of functions that can be passed as a suffix to aggregation columns separated by a pipe character (|).If no function is specified, then the default aggregation function is 'sum'.

    
     
     
