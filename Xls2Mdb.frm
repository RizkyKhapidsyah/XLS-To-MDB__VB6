VERSION 5.00
Begin VB.Form frmXls2Mdb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excel To Access"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Proceed"
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   1050
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Example of appending a worksheet of Excel workbook to  a new table in Access database (in same folder)"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   150
      Width           =   2655
   End
End
Attribute VB_Name = "frmXls2Mdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Xls2mdb.frm
'
' By Herman Liu
'
' An example of appending an Excel worksheet to an existing Access database.
'
' Required: (1) Project references to include a Microsoft DAO Object Library; (2) A
' database file and (3) A worksheet of Excel Workbook (Row 1 for field names).
'
' Note You may also wish to refer MSDN "How to Append an Excel Worksheet to a
' Database Using DAO"

Option Explicit


Private Sub Command1_Click()
    On Error GoTo errHandler
    Dim mExcelFile As String
    Dim mAccessFile As String
    Dim mWorkSheet As String
    Dim mTableName As String
    Dim mDataBase As Database
    mExcelFile = App.Path & "\Book1.xls"
    mAccessFile = App.Path & "\Db1.mdb"
    mWorkSheet = "Sheet1"
    mTableName = "Table1"
      ' Below you may use "Excel 7.0" or 8.0 depending on your installable ISAM.
    Set mDataBase = OpenDatabase(mExcelFile, True, False, "Excel 5.0")
    mDataBase.Execute "Select * into [;database=" & mAccessFile & "]." & mTableName & _
        " FROM [" & mWorkSheet & "$]"
    MsgBox "Done.  Use Access to view " & mTableName
    Exit Sub
errHandler:
    If Err.Number = 3010 Then
         MsgBox mTableName & " already exist." & vbCrLf & _
             "Delete " & mTableName & " first or use another table name."
    Else
         MsgBox Err.Number & "  " & Err.Description
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
