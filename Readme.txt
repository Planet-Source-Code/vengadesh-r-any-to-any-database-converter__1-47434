Any-to-Any database converter
=============================

06 Aug 2003

by vengy, Pondichery, India.

vb4vengy@yahoo.com

Purpose
=======
Using ADO & ADOX, convert between various
database formats like Text, Access, Excel,
dBASE, Paradox, HTML. 
(Oracle and SQLServer in next version)

you are not required to have that specific application
installed to generate the destination (output) database.
for example, without having Excel installed, you can create Excel 
sheet. 

Method
======

1. Select a source database type

2. Select corresponding source database file

3. If tables are available within the selected database file (like Access),
   they will be displayed, otherwise available fileds will be
   displayed.

4. Select a table (if available), and fields within the table will
   be displayed.

5. Select the fields as per your wish, and the selected fields will
   be displayed in a list box, where you can re-order the fields.

6. Select a destination database type

7. Give the destination database file name

8. Click "Create" button. 

that is all.

Note
====

1. This program is not complete in all respects, and 
   it is not capable of converting all kinds of data
   formats. for example binary objects, etc.

2. I haven't written code for Oracle, SQLServer and html
   conversion, these will be included in the next version.

3. You need vb6, ADO 2.6 (ADO 2.0 and higher will do), 
   ADO Ext. 2.7 for DLL and Security (ADOX) to run this program. 
   No other object library is used (excel object lib for example).

4. This program will copy data only. 
   Constraints & relations are not yet handled.

5. This example program is provided "as is" with no 
   warranty of any kind. It is intended for demonstration 
   purposes only. In particular, it does no error 
   handling. 

6. Special thanks to ThomasOBascom@compuserve.de

7. Please email your comments/suggestions/questions/ideas/errors to :

   vb4vengy@yahoo.com
