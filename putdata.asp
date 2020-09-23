<%
dim rs
dim cn
dim doc 

' Create Objects
set cn = server.CreateObject("ADODB.Connection")
set rs = server.CreateObject("ADODB.Recordset")
Set doc = server.CreateObject("MSXML2.DOMDocument")

' Open Connection - I am using a System DSN on the server here
' you need to substitute your details here or replace the parameters
' with a valid connection string
cn.Open "DSN_Name", "UserID", "Password"

' Set up Recordset 
rs.CursorLocation = 3	'adUseClient

' Load Request into DomDocument
doc.load Request
 
' Open Recordset with DomDocument as the Source
rs.Open doc

' Connect incoming Recordset to Connection
rs.ActiveConnection = cn

' Update the database behind the Connection
rs.UpdateBatch

' Tidy Up
set doc=nothing
rs.Close
set rs=nothing
cn.Close
set cn=nothing

' Return Response - will contain error message if failure occurs above
Response.Write "Data Updated"
%>