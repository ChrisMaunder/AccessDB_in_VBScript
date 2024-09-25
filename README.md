# Accessing Microsoft Access databases in ASP using ADO

A simple introduction to using Access .mdb databases in your ASP pages



## Introduction

Windows DNA provides a means to provide your user interface, business logic and data
sources as separate services working together in harmony over a distributed environment.
The browser has become an extremely powerful, yet simple method of providing the user
interface, since it handles the network considerations and allows you to create rich
user interfaces through simple scripting, HTML and style sheets.

Your database considerations can be taken care of simply through the use of SQLServer
or the Microsoft Jet Engine, and your business logic - the guts of your application that
processes the data from the database and sends it to the browser - can be simple ASP
pages (enhanced with ActiveX controls if the fancy takes you).

Once you have the basics of ASP, HTML and VBScript the business logic and user 
interface are taken care of quickly and simply - but how do you use ASP to access your
database and hence complete your 3-tier application? Read on...

## Simple database Access using ADO and ASP

For this example we'll use Access .mdb databases - but we could just as easily 
use SQLServer by changing a single line (and of course, configuring the databases
correctly). We'll be assuming your application is ASP based running on Microsoft's
IIS Webserver.

We use ADO since it is portable, widespread, and very, very simple.

### The Connection

To access a database we first need to open a connection to it, which involves
creating an ADO Connection object. We then specify the connection string and call
the Connection object's `Open` method.

To open an Access database our string would look like the following:

```vbscript
Dim ConnectionString
ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};" &_
                   "DBQ=C:\MyDatabases\database.mdb;DefaultDir=;UID=;PWD=;"
```

where the database we are concerned with is located at C:\MyDatabases\database.mdb, and has
no username or password requirements. If we wanted to use a different database driver (such as SQLServer) then we simply provide a different connection string. 

To create the ADO Connection object simply `Dim` a variable and get the server to do the work.

```vbscript
Dim Connection
Set Connection = Server.CreateObject("ADODB.Connection")
```

Then to open the database we (optionally) set some of the properties of the Connection
and call `Open`

```vbscript
Connection.ConnectionTimeout = 30
Connection.CommandTimeout = 80
Connection.Open ConnectionString
```

Check for errors and if everything is OK then we are on our way.

### The Records

Next we probably want to access some records in the database. This is achieved via
the ADO RecordSet object. Using this objects `Open` method we can pass in
any SQL string that our database driver supports and receive back a set of records 
(assuming your are SELECTing records, and not DELETEing).

```vbscript
' Create a RecordSet Object
Dim rs
set rs = Server.CreateObject("ADODB.RecordSet")

' Retrieve the records
rs.Open "SELECT * FROM MyTable", Connection, adOpenForwardOnly, adLockOptimistic
```

`adOpenForwardOnly` is defined as 0 and specifies that we only wish to
traverse the records from first to last. `adLockOptimistic` is defined as 3
and allows records to be modified.

If there were no errors we now have access to all records in the table "MyTable"
in our database.

The final step is doing something with this information. We'll simply list it.

```vbscript
' This will list all Column headings in the table
Dim item
For each item in rs.Fields
Response.Write item.Name & "<br>"
next
        
' This will list each field in each record
while not rs.EOF
    
For each item in rs.Fields
    Response.Write item.Value & "<br>"
next
    
rs.MoveNext
wend

End Sub
```

If we know the field names of the records we can access them using `rs("field1")`
where `field1` is the name of a field in the table.

Always remember to close your recordsets and Connections and free any resources
associated with them

```vbscript
rs.Close
set rs = nothing

Connection.Close
Set Connection = nothing
```

## Conclusion

This has been an extremely simple demonstration without serious error checking
or even legible formatting of the output, but it's a base to start with.
