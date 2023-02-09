# ETABS-ColumnGrouping
Python code for assigning groups to columns in ETABS on the basis of their percentage reinforcement. 

This code does following:
1. Connects to a running ETABS window
2. Reads reinforcement data from data.xlsx.
3. Selects columns in ground floor.
4. Finds columns lying just above it on upper floors.
5. Assign Groups to the columns based on percentage rebar of column in ground floor.

