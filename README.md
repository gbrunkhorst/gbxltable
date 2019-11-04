# gbxltable
gbxltable is a little python module to make a formatted excel file from a pandas dataframe in Greg Brunkhorst's preferred format (hence gb) using xlsxwriter.

Note this is the first module the author has made!  I made the module to formalize a repeated workflow working with messy environmental data.    

What it does:  
* formats cells to show decimals to a number of significant (default 2)
* center-justify, bold heading, cell borders, cell width 10-30 depending on the content
* adds the date to the file name

Limitations to be aware of:
* the program drops the index.  If your index contains data, df.reset_index() before producing the excel table.
* the program can't handle multi-index

Limitations to be unaware of:
* the program doesn't do most anything you can think of!

