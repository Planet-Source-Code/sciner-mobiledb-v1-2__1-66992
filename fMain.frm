VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "speed test"
      Height          =   1455
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdTestAddRow 
         Caption         =   "add 1000 rows with few fields and fill data"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton cmdTestFields 
         Caption         =   "create 1000 fields"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton cmdTestTables 
         Caption         =   "create 1000 tables"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtCondition 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "9"
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cmbFields 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Search text:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "In field:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test (create table, fields and fill records)"
      Default         =   -1  'True
      Height          =   855
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvDB 
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'© 2000-2007 SCINER
'Created by SCINER: lenar2003@mail.ru
'03/11/2006 20:41
'Custom DataBase Engine v1.0

Dim DB As New cMobileDB
Dim dwtable As Long
Dim DBFileName As String

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim Tick As Long

'return correct path with slash
Function RCP(ByVal P As String) As String
  RCP = P & IIf(VBA.Right$(P, 1) = "\", vbNullString, "\")
End Function

Private Sub cmdFind_Click()
  If dwtable < 1 Then Exit Sub
  Dim lOffsetField As Long
  Dim lOffsetRow As Long
  Dim v 'value
  Dim szSearch As String
  Dim lRowIndex As Long
  lOffsetField = DB.getFieldOffset(dwtable, cmbFields.Text)
  If lOffsetField < 1 Then Exit Sub
  lOffsetRow = DB.getFirstTableRow(dwtable)
  szSearch = txtCondition.Text
  Do While lOffsetRow > 0
    lRowIndex = lRowIndex + 1
    v = DB.Value(dwtable, lOffsetField, lOffsetRow)
    'check field type for ByteArray!!!
    v = CStr(v)
    If v Like "*" & szSearch & "*" Or _
       InStr(v, szSearch) > 0 Then
      lvDB.ListItems(lRowIndex).Selected = True
      MsgBox "Found in #" & CStr(lRowIndex) & " row", 32
      Exit Sub
    End If
    lOffsetRow = DB.getNextRow(lOffsetRow)
  Loop
  MsgBox "Nothing not found", 16
End Sub

Private Sub cmdTest_Click()

  Const szTableName As String = "TestTable"
  Const szFiledNum As String = "Number"
  Const szFiledRandom As String = "Random number"
  Const szFiledComment As String = "Comment"
  
  Dim dwFiledNum As String
  Dim dwFiledRandom As String
  Dim dwFiledComment As String

  Dim lRow As Long
  Dim lField As Long
  Dim v 'buffer for value
  Dim i As Long
  Dim z As Long
  Dim capcount As Long
  Dim lItem As ListItem

  Call DB.CloseDB
  Call DeleteFile(DBFileName)

  'set database file
  DB.Filename = DBFileName
  
  cmdFind.Enabled = True

'-------------------------------------------------------
'write to database table
'-------------------------------------------------------

  'add table
  dwtable = DB.addTable(szTableName, mNormal)
  Call DB.Update
  'get table offset
  dwtable = DB.getTableOffset(szTableName)

  'if table exists
  If dwtable > 0 Then
    'add fields
    dwFiledNum = DB.addField(dwtable, szFiledNum, fldLong)
    dwFiledRandom = DB.addField(dwtable, szFiledRandom, fldDouble)
    dwFiledComment = DB.addField(dwtable, szFiledComment, fldString)
  End If
  
  'for flush new fields to file
  Call DB.Update

  'get fields offset
  dwFiledNum = DB.getFieldOffset(dwtable, szFiledNum)
  dwFiledRandom = DB.getFieldOffset(dwtable, szFiledRandom)
  dwFiledComment = DB.getFieldOffset(dwtable, szFiledComment)

  Call Randomize(Timer)
  
  'add items to table
  For i = 0 To 99
    'get new row offset
    lRow = DB.Add(dwtable)
    'set values
    DB.Value(dwtable, dwFiledNum, lRow) = i
    DB.Value(dwtable, dwFiledRandom, lRow) = Rnd
    DB.Value(dwtable, dwFiledComment, lRow) = "Created by SCINER: lenar2003@mail.ru is " & CStr(i)
    'slow speed
    'Call db.Update
  Next
  
  'for flush db records to file
  Call DB.Update
  
  Call ReadFromDb(szTableName)

End Sub

Sub ReadFromDb(ByVal szTableName As String)

'-------------------------------------------------------
'read from table from database
'-------------------------------------------------------

  Dim lRow As Long
  Dim lField As Long
  Dim v 'buffer for value
  Dim i As Long
  Dim z As Long
  Dim capcount As Long
  Dim lItem As ListItem

  Call DB.CloseDB

  'set database file
  DB.Filename = DBFileName

  Call lvDB.ColumnHeaders.Clear
  Call lvDB.ListItems.Clear

  'get table offset
  dwtable = DB.getTableOffset(szTableName)

  'get all table fields
  lField = DB.getFirstField(dwtable)
  Do While lField > 0
    Call lvDB.ColumnHeaders.Add(, , DB.getFieldName(lField))
    lField = DB.getNextField(lField)
  Loop
  
  Call cmbFields.Clear
  lField = DB.getFirstField(dwtable)
  Do While lField > 0
    Call cmbFields.AddItem(DB.getFieldName(lField))
    lField = DB.getNextField(lField)
  Loop
  If cmbFields.ListCount > 0 Then cmbFields.Text = cmbFields.List(0)
 

  With DB
    'get first table row offset
    lRow = DB.getFirstTableRow(dwtable)
    'enumerate all rows in table
    Do While lRow > 0
      'get first table field offset
      lField = .getFirstField(dwtable)
      Set lItem = lvDB.ListItems.Add(, , .Value(dwtable, lField, lRow))
      i = 1
      'get next table field
      lField = .getNextField(lField)
      Do While lField > 0
        v = Empty
        'get field type
        Select Case .getFieldType(lField)
        Case fldByteArray: v = "<BINARY>"
        Case fldJpegFileBytes: v = "<BINARY>"
        Case Else: v = .Value(dwtable, lField, lRow)
        End Select
        If IsEmpty(v) Then v = "NULL"
        lItem.SubItems(i) = v
        i = i + 1
        'get next table field
        lField = .getNextField(lField)
      Loop
      'go to next row
      lRow = DB.getNextRow(lRow)
    Loop
  End With

End Sub

'TEST - CREATE MANY ROWS WITH DATA
Private Sub cmdTestAddRow_Click()
  
  Const FIELD_COUNT As Long = 2
  Const RECORD_COUNT As Long = 1000
  
  Dim i As Long
  Dim j As Long
  Dim dwtable As Long
  Dim szTableName As String
  Dim tFields(0 To FIELD_COUNT - 1) As Long
  Dim lRow As Long
  Call DB.CloseDB
  Call DeleteFile(DBFileName)

  szTableName = "TestTable"
  DB.Filename = DBFileName
  dwtable = DB.addTable(szTableName, mNormal)
  
  Call DB.Update
  
  For i = 0 To FIELD_COUNT - 1
    tFields(i) = DB.addField(dwtable, "Field#" & CStr(i), fldLong)
  Next

  Call DB.Update

  Tick = GetTickCount
  For i = 0 To RECORD_COUNT - 1
    lRow = DB.Add(dwtable)
    For j = 0 To FIELD_COUNT - 1
      DB.Value(dwtable, tFields(j), lRow) = j * 1000000 + i
    Next
  Next
  Call DB.Update
  Tick = GetTickCount - Tick
  
  MsgBox "Elapsed " & Format$(Tick / 1000, "0.000") & " second", 32
  
  Call ReadFromDb(szTableName)

End Sub

'TEST - CREATE MANY TABLES
Private Sub cmdTestTables_Click()
  Dim i As Long
  Dim szTableName As String
  Call DB.CloseDB
  Call DeleteFile(DBFileName)
  DB.Filename = DBFileName
  Tick = GetTickCount
  For i = 0 To 999
    szTableName = "Table#" & CStr(i)
    Call DB.addTable(szTableName, mNormal)
  Next
  Call DB.Update
  Tick = GetTickCount - Tick
  MsgBox "Elapsed " & Format$(Tick / 1000, "0.000") & " second", 32
End Sub

'TEST - CREATE MANY FIELDS
Private Sub cmdTestFields_Click()
  Dim i As Long
  Dim dwtable As Long
  Dim szFieldName As String
  Call DB.CloseDB
  Call DeleteFile(DBFileName)
  DB.Filename = DBFileName
  dwtable = DB.addTable("TestTable", mNormal)
  Tick = GetTickCount
  For i = 0 To 999
    szFieldName = "Field#" & CStr(i)
    Call DB.addField(dwtable, szFieldName, fldByteArray)
  Next
  Call DB.Update
  Tick = GetTickCount - Tick
  MsgBox "Elapsed " & Format$(Tick / 1000, "0.000") & " second", 32
End Sub

Private Sub Form_Load()
  DBFileName = RCP(App.Path) & "db.sbs"
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  On Error Resume Next
  Call lvDB.Move(lvDB.Left, lvDB.Top, ScaleWidth - lvDB.Left * 2, ScaleHeight - lvDB.Top - lvDB.Left)
End Sub
