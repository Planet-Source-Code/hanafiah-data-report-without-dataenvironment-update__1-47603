VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "All About Data Report"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Sample 4"
      Height          =   615
      Left            =   105
      TabIndex        =   7
      Top             =   2430
      Width           =   2310
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sample 3"
      Height          =   630
      Left            =   90
      TabIndex        =   6
      Top             =   1680
      Width           =   2340
   End
   Begin VB.PictureBox Picture1 
      Height          =   2145
      Left            =   3495
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2085
      ScaleWidth      =   3810
      TabIndex        =   3
      Top             =   405
      Width           =   3870
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2145
      Left            =   3015
      OleObjectBlob   =   "Form1.frx":D1EE
      TabIndex        =   2
      Top             =   3105
      Width           =   4560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sample 2"
      Height          =   630
      Left            =   75
      TabIndex        =   1
      Top             =   960
      Width           =   2340
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample 1"
      Height          =   660
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   "Graph"
      Height          =   315
      Left            =   3570
      TabIndex        =   5
      Top             =   2820
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "Picture"
      Height          =   315
      Left            =   3585
      TabIndex        =   4
      Top             =   105
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------
' Code By Sudu_Sudu@hotmail.com
' Name : Muhamad Hanafiah Yahya
' Require
'            Microsoft Active X DataObject 2.7
'            Microsoft Chart Control ( Service Pack 4 )
'--------------------------------------------------------------------------
Dim MyCon As New ADODB.Connection
Dim MyRs As New ADODB.Recordset

Private Sub Command1_Click()
'------------------------------------------------------------------------------------------
'This sample to display Report direct from database without any additional
' operation like add data, combine and etc
'------------------------------------------------------------------------------------------
'To set Textbox Datafield
With DataReport1.Sections("Section1").Controls 'section1 mean that section you create in datareport
    .Item("txtFamily").DataField = MyRs("HouseholdName").Name
    .Item("txtAddress").DataField = MyRs("Address").Name
End With

'To Set Label caption
With DataReport1.Sections("Section2").Controls
    .Item("lblFamily").Caption = "Family Name"
    .Item("lblAddress").Caption = "Address"
End With

With DataReport1.Sections("Section4").Controls
    .Item("lblTitle").Caption = "My Family Name - Sample 1"
End With

'to set datasource for datareport
Set DataReport1.DataSource = MyRs

'show datareport
DataReport1.Show
End Sub

Private Sub Command2_Click()
'------------------------------------------------------------------------------------------
'This sample to display Report from database with any additional
' operation like add data, combine and etc
'------------------------------------------------------------------------------------------
' Create adodb record set
Dim intCount As Integer
Dim strAddress As String
Dim TempRS As ADODB.Recordset

Set TempRS = New ADODB.Recordset
TempRS.Fields.Append "tmpFamily", adVarChar, 30
TempRS.Fields.Append "tmpAddress", adVarChar, 100
TempRS.Open
MyRs.MoveFirst 'set to first record

For intCount = 1 To MyRs.RecordCount
    strAddress = MyRs("Address") & " , " & MyRs("city") & ", " & MyRs("StateOrProvince")
    TempRS.AddNew Array("tmpFamily", "tmpAddress"), Array(intCount & " " & MyRs("HouseholdName"), strAddress)
    MyRs.MoveNext
Next intCount

'To set Textbox Datafield
With DataReport1.Sections("Section1").Controls 'section1 mean that section you create in datareport
    .Item("txtFamily").DataField = TempRS("tmpFamily").Name
    .Item("txtAddress").DataField = TempRS("tmpAddress").Name
End With

'To Set Label caption
With DataReport1.Sections("Section2").Controls
    .Item("lblFamily").Caption = "Family Name"
    .Item("lblAddress").Caption = "Address"
End With

'set report title
With DataReport1.Sections("Section4").Controls
    .Item("lblTitle").Caption = "My Family Name - Sample 2"
End With

'to set datasource for datareport
Set DataReport1.DataSource = TempRS

'show datareport
DataReport1.Show
End Sub

Private Sub Command3_Click()
'------------------------------------------------------------------------------------------
'This sample to display Graph and Picture in Datareport
'------------------------------------------------------------------------------------------
'this TempRS used for dummies data only. you can manipulate step for sample1 and sample2
Dim TempRS As ADODB.Recordset
Set TempRS = New ADODB.Recordset
TempRS.Fields.Append "tmpFamily", adVarChar, 30
TempRS.Fields.Append "tmpAddress", adVarChar, 100
TempRS.Open
TempRS.AddNew Array("tmpFamily", "tmpAddress"), Array("test", "test")

'copy mschart1 image
MSChart1.EditCopy

'Must put Set for image
With DataReport2.Sections("Section1")
    Set .Controls("Image1").Picture = Form1.Picture1.Picture
    Set .Controls("Image2").Picture = Clipboard.GetData(vbCFDIB)
End With

' just for dummies data
Set DataReport2.DataSource = TempRS
DataReport2.Show

'clear clip board
Clipboard.Clear

End Sub

Private Sub Command4_Click()
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cn As New ADODB.Connection
Dim strSQL As String
cn.Open "Provider=MSDATASHAPE; Data Provider=Microsoft.JET.OLEDB.4.0;" & _
             "Data Source=" & App.Path & "\database\ADDRBOOK.mdb"
             
strSQL = "Select  * from HouseHold "

With cmd
    .ActiveConnection = cn
    .CommandType = adCmdText
    .CommandText = " SHAPE {" & strSQL & "}  AS cmdGroup Compute cmdGroup BY 'Country'"
    .Execute
End With

With rs
    .ActiveConnection = cn
    .CursorLocation = adUseClient
    .Open cmd
End With

With DataReport3
    Set .DataSource = rs
    .DataMember = ""
'this for group header
    With .Sections("section2").Controls
        .Item("Text1").DataField = "Country"
    End With
'this for group details
    With .Sections("Section1").Controls
        .Item("Text2").DataMember = "cmdGroup"
        .Item("Text2").DataField = "HouseholdName"
        .Item("Text3").DataMember = "cmdGroup"
        .Item("Text3").DataField = "Address"
    End With
    .Refresh
    .Show
End With
End Sub

Private Sub Form_Load()
Dim strPath As String

strPath = App.Path & "\database\ADDRBOOK.mdb"
 
'Set connection ke database ( strpath )
MyCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "Data Source=" & strPath
MyCon.Open

'Open The The Recordset
MyRs.ActiveConnection = MyCon

'Open sebagai Keyset, dan LockOptimistic
'MyRs.Open "HouseHold", MyCon, adOpenKeyset, adLockOptimistic, adCmdTableDirect
MyRs.Open "Select  * from HouseHold ", MyCon, adOpenKeyset, adLockOptimistic, adCmdTableDirect
End Sub

Private Sub Form_Unload(Cancel As Integer)
MyRs.Close
Set MyCon = Nothing
End Sub
