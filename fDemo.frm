VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fDemo 
   Caption         =   "ADO Web Data Demo"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12726
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdPut 
      Caption         =   "Write Data to Server"
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Data from Server"
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter criteria here (leave blank for all records)"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple demo - how to get data from web server, edit it and
' put it back again

' By John Pragnell - john@simsoc.com - April 2001

' *Thanks to A J Brust (of Progressive Systems Consulting, Inc.) for inspiration
' & encouragement

' REQUIRES:
'   Internet Explorer 5 on PC
'   MS XML 3.0 on PC & Server
'   PWS or IIS 4+
'   ADO 2.5+ on server on PC & Server

Option Explicit

' Recordset definition
Dim rs As New ADODB.Recordset
'if the link is invalid or mis-spelt then the error message is:
'3709 - Operation is not allowed on an object referencing a closed or invalid connection.

'One way to test the link is to to type 'http://www.simsoc.com/adoweb/getdata.asp?ID=0'
'into the address bar of your browser .. it should display the data

' Change to your URL address - I will leave my link working as long as possible
Const urlData = "http://www.simsoc.com/adoweb/"

Private Sub cmdGet_Click()
   
    Screen.MousePointer = vbHourglass
    ' Get data
    Set rs = Nothing
    
    rs.Open urlData & "getdata.asp?ID=" & Val(txtID) & ""
   
    ' Display for edits
    Set DataGrid.DataSource = rs

    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdPut_Click()
    'for XML 3.0
    Dim xml As New MSXML2.XMLHTTP
    Dim doc As New MSXML2.DOMDocument
   
    Screen.MousePointer = vbHourglass

    ' Optional line - to help reduce traffic for large data sets
    rs.Filter = adFilterPendingRecords

    ' Write changes
    xml.Open "POST", urlData & "putdata.asp", False
    rs.Save doc, adPersistXML
    xml.send doc

    Set xml = Nothing
    
    Set doc = Nothing

    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set rs = Nothing
    rs.Close
    
End Sub
