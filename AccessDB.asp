<%@ Language=VBScript %>
<% Option Explicit %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<%

' this set of scripts demonstrates how to open a Microsoft Access
' .mdb file and display the contents of a given table.


On Error Resume Next

' Constants
const adOpenForwardOnly = 0 ' Forward only movement through records
const adOpenKeyset      = 1 ' Movement any direction

const adOpenDynamic = 2
const adOpenStatic  = 3
const adLockOptimistic = 3

' CheckError: Checks for, and reports Errors. 
' Returns True if an error has occured, False otherwise
Function CheckError()
	if Err.number <> 0 then
		Response.Write "<p><FONT color=red><I>A run-time error occurred.<BR>Error Number: " &_
                           Err.number & "<BR>Error Description: " & Err.description & "</Font></I>"
		CheckError = True
	Else
		CheckError = False
	End If
end Function

' OpenDatabase: Opens a database connection
' IN:
' - DatabaseVirtualFilename: Virtual path to database file (.mdb)
' - Username: Username
' - Password: Password
' OUT:
' - Connection: The database connection
' RETURNS:
' - True if successful, False otherwise
Function OpenDatabase(DatabaseVirtualFilename, Username, Password, byref Connection)

	' Initialise
	OpenDatabase = False
	
	Dim DatabaseFilename
	DatabaseFilename = Server.MapPath(DatabaseVirtualFilename)

	Dim ConnectionString
	ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};" &_
	                   "DBQ=" & DatabaseFilename & ";DefaultDir=;" &_
	                   "UID=" & Username & ";" &_
	                   "PWD=" & Password & ";"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Connection.ConnectionTimeout = 30
	Connection.CommandTimeout = 80
	Connection.Open ConnectionString
		
	OpenDatabase = True
		
end Function

' DisplayTable: Displays the contents of a table
' IN:
' - Connection: Connection to a database
' - TableName:  The table to list
Sub DisplayTable(Connection, TableName)
		
	Response.Write "<h2>Contents of table: " & TableName & "</h2>" & vbCRLF
	
	Dim SQL
	SQL = "SELECT * FROM " & TableName

	' Create a RecordSet
	Dim rs
	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open SQL, Connection, adOpenForwardOnly, adLockOptimistic
		
	if rs.EOF or rs.BOF Then
		Response.Write "<p align=center><i> -- No records --</i>" & vbCRLF
		rs.Close
		set rs = nothing
		Exit sub
	end if
	
	' Start the table
	Response.Write "<table width='100%' border=1>" & vbCRLF

	' Write field headings
	Response.Write "<tr>"
	Dim item
	For each item in rs.Fields
		Response.Write "<td valign=top><b>" & item.Name & "</b></td>"
	next
	Response.Write "</tr>" & vbCRLF
			
	' Write out records
	while not rs.EOF
		
		Response.Write "<tr>"
		For each item in rs.Fields
			Response.Write "<td valign=top>" & item.Value & "</td>"
		next
		Response.Write "</tr>" & vbCRLF
		
		rs.MoveNext
	wend
	
	' End the table
	Response.Write "</table>" & vbCRLF

	rs.Close
	set rs = nothing

End Sub

' ////////////////////////////////////////////////////////////
' Main

' Get the database and table names. If non-null then display the
' details, otherwise just display the input form.
Dim databasefile, databasetable
databasefile = Request("database")
databasetable = Request("table")

if (databasefile <> "" and databasetable <> "") then

	Dim Connection
	if (OpenDatabase(databasefile, "", "", Connection)) Then
		DisplayTable Connection, databasetable
	End if
	
	Connection.Close
	set Connection = nothing
	
	CheckError
End If

%>
<br>
<br>

<!-- 
This form is for specifying the database and table names. The
form posts the data back to this file
-->

<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<table>
	<tr>
		<td nowrap>Database filename (virtual path):</td>
		<td><INPUT type="text" size=40 name=database value="<%=databasefile%>"></td>
	</tr>
	<tr>
		<td nowrap>Database Table:</td>
		<td><INPUT type="text" size=40 name=table value="<%=databasetable%>"></td>		
	</tr>
	<tr>
		<td nowrap></td>
		<td><INPUT type="submit" value=" Go! "></td>		
	</tr>
	</table>
</form>

</BODY>
</HTML>
