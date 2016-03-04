# ExcelWriter
Create excel files using only the microsoft jet engine without having to install office on the server.

I've been pretty frustrated trying to create excel files. 
I used to create a fake html file that could be opened in excel but it wasn't practical.

Finally after doing some research I found that you can create excel files without having excel installed on the server.
You will need a version of the Microsoft Access Database Engine 2010 Redistributable . 
https://www.microsoft.com/en-us/download/details.aspx?id=13255

Once you have that setup you'll want to use this in the following way.

You'll want to pass a DataView, a temp path to hold the file and the connection string.

This is the one I use. Note the fileWpath. It's where you'll be creating the excel file to.

Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileWpath + ";" + "Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"

I hope this helps people out there.
