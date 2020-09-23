<%
dim rs
dim doc
dim cn
dim rqs

' Create Objects
set cn = server.CreateObject("ADODB.Connection")
set rs = server.CreateObject("ADODB.Recordset")
set doc = server.CreateObject("MSXML2.DomDocument")

' Querystring has parameter, if any specified
rqs = Request.QueryString("ID")

' Open Connection - I am using a System DSN on the server here
' you need to substitute your details here or replace the parameters
' with a valid connection string
cn.Open "DSN_Name", "UserID", "Password"

' Set up Recordset 
rs.CursorLocation = 3	'adUseClient

' Get required data
if rqs = 0 then
	rs.Open "Select * from Member", cn, 1, 4	'adOpenKeyset, adLockBatchOptimistic
else
	rs.Open "Select * from Member where Memberid = " & rqs & "", cn, 1, 4 	'adOpenKeyset, adLockBatchOptimistic
end if

' Save Recordset to DOMDoc as XML
' You need to include the ADOVBS.INC file to use adPersistXML keyword / otherwise use 1 as its value
rs.Save doc, 1		'adPersistXML '1

' Set Response property
Response.ContentType = "text/xml"

' Stylesheet for browser display
Response.Write "<?xml:stylesheet type=""text/xsl"" href=""Recordsetxml.xsl""?>" & vbcrlf

' Send XML from DOMDoc to Response
Response.Write doc.xml

' Tidy Up
set doc=nothing
rs.Close
set rs=nothing
cn.Close
set cn=nothing

%>