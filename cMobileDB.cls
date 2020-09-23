VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMobileDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'© 2000-2007 SCINER
'Created by SCINER: lenar2003@mail.ru
'03/11/2006 20:41
'Custom DataBase Engine v1.1

'---------------------
'- ENGLISH COMMENTS --
'REALIZED  ------------------------------------------------------------------------
' 0. unlimited number of tables
' 1. unlimited fields in tables
' 2. unlimited records in a table
' 3. the rational use of memory in a file
' 4. support 8 types of data, including JpegFileBytes
' 5. reference structure, providing high speed of treatment of base
' 6. a type of string is the buffer of variable size (memo)
'- REALIZED -----------------------------------------------------------------------
' 0. support tables with the name length to 127 byte
' 1. support fields with the name length to 127 byte and with the chosen type
' 2. addition of records
' 3. pour a record/reading
' 4. 3 events including "Change" on value is changed
'- NOT REALIZED -------------------------------------------------------------------
' 0. delete of table, fields, records not support
' 1. fragmentation of file of database not support

'---------------------
'- RUSSIAN COMMENTS --
'- ÂÎÇÌÎÆÍÎÑÒÈ --------------------------------------------------------------------
' 0. íåîãðàíè÷åííîå ÷èñëî òàáëèö
' 1. íåîãðàíè÷åííîå ÷èñëî ïîëåé â òàáëèöàõ
' 2. íåîãðàíè÷åííîå ñè÷ëî çàïèñåé â òàáëèöå
' 3. ðàöèîíàëüíîå èñïîëüçîâàíèå ïàìÿòè â ôàéëå
'    (ïîëå ôèçè÷åñêè íå ñóùåñòâóåò äî òåõ ïîð,
'    ïîêà â íåãî ÷òî íèáóäü íå çàïèñàëè)
' 4. ïîääåðæêà 8 òèïîâ äàííûõ, âêëþ÷àÿ JpegFileBytes
' 5. ññûëî÷íàÿ ñòðóêòóðà, îáåñïå÷èâàþùàÿ âûñîêóþ ñêîðîñòü îáðàáîòêè áàçû
' 6. òèï string ÿâëÿåòñÿ áóôåðîì ïåðåìåííîãî ðàçìåðà aka memo
'- ÐÅÀËÈÇÎÂÀÍÎ --------------------------------------------------------------------
' 0. ñîçäàíèå òàáëèö ñ ïðîèçâîëüíûì íàçâàíèåì äëèíîé äî 256 áàéò
' 1. ñîçäàíèå ïîëåé ñ ïðîèçâîëüíûì íàçâàíèåì äëèíîé äî 256 áàéò è ñ âûáðàííûì òèïîì
' 2. äîáàâëåíèå çàïèñåé
' 3. çàïèñü/÷òåíèå ïîëåé
'- ÏÎÊÀ ÍÅ ÐÅÀËÈÇÎÂÀÍÎ ------------------------------------------------------------
' 0. óäàëåíèå òàáëèöû, ïîëÿ, çàïèñè
' 1. ôðàãìåíòàöèþ ôàéëà áàçû äàííûõ

Private Const DB_HEAD_LENGTH = 264
Private Const SIGN_DB = 23658740
Private Const SIGN_DATA = 255
Private Const SIGN_ROW = 12548

Event ExportProgress(ByVal lMax As Long, ByVal lValue As Long)
Event ExportComplete()
Event Change(ByVal lOffsetTable As Long, ByVal lOffsetRow As Long, ByVal lOffsetField As Long, ByVal lOffsetData As Long)

Public Enum MobDbAttributes
  mSystem = vbSystem
  mReadOnly = vbReadOnly
  mNormal = vbNormal
  mHidden = vbHidden
  mAlias = vbAlias
End Enum

Public Enum DBFieldType
  fldString = 1
  fldLong = 2
  fldDouble = 4
  fldByte = 8
  fldDate = 16
  fldInteger = 32
  fldByteArray = 64
  fldJpegFileBytes = 128
  'fld = 256
  'fld = 512
  'fld = 1024
  'fld = 2048
  'fld = 4196
  'fld = 8192
  'fld = 16384
End Enum

Private Type data_type_2
  size As Long
  sizereal As Long
  next As Long
  lOffset As Long
  fieldoffset As Long
  sign As Byte
  Type As Integer
End Type

Private Type row_type_2
  sign As Integer
  reserved(0 To 1) As Byte
  Data As Long
  previous As Long
  next As Long
End Type

Private Type field_type_2
  sign As Long
  Type As Integer
  szName(255) As Byte
  next As Long
End Type

Private Type table_type_2
  sign As Long
  record As Long
  lastrecord As Long
  field As Long
  lastfield As Long
  next As Long
  RecordCount As Long
  szName(127) As Byte
  attributes As Integer
  reserved(124) As Byte
End Type

Private Type last_update_type_2
  tableoffset As Long
  rowoffset As Long
  fieldoffset As Long
  datetime As Date
End Type

Private Type db_head_type_2
  sign As Long
  table As Long
  lasttable As Long
  lastupdate As last_update_type_2
  szName(63) As Byte
  uniqueid As Long 'not used
  reserved(187) As Byte
End Type

Private Type fastfieldindex
  Key As String
  lOffset As Long
  Type As DBFieldType
End Type

Private Type fastindex
  Key As String
  lOffset As Long
  Fields() As fastfieldindex
  fieldscount As Long
End Type

Dim DB As db_head_type_2
Dim tables() As fastindex
Dim lTableCountArraySize As Long
Dim lTableCount As Long
Dim FF As Integer
Dim lpFileName As String

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'return last update information for current database
'lastupdate automaticaly update if any value is change
'
'Ïîëó÷åíèå èíôîðìàöèè î ïîñëåäíåì îáíîâëåíèè ëþáîé ÿ÷åéêè áàçû äàííûõ
Function getLastUpdate(dtDate As Date, lOffsetTable As Long, lOffsetRow As Long) As Boolean
  Dim tempdb As db_head_type_2
  If FF <> 0 Then
    Get #FF, 1, tempdb
    dtDate = tempdb.lastupdate.datetime
    lOffsetTable = tempdb.lastupdate.tableoffset
    lOffsetRow = tempdb.lastupdate.rowoffset
    getLastUpdate = True
  End If
End Function

'return offset for first table of database
'
'âîçâðàùàåò àäðåñ ïåðâîé òàáëèöû èç áàçû
Function getFirstTable() As Long
  If DB.sign = SIGN_DB Then getFirstTable = DB.table
End Function

'get offset for next table by previous table offset
'if success this function return value > 0
'if failed return value is 0
Function getNextTable(ByVal lTableOffset As Long) As Long
  Dim table As table_type_2
  If lTableOffset < 0 Then Exit Function
  Get #FF, lTableOffset, table
  If table.sign <> SIGN_DB Then Exit Function
  If table.next < 1 Then Exit Function
  lTableOffset = table.next
  Get #FF, lTableOffset, table
  If table.sign = SIGN_DB Then getNextTable = lTableOffset
End Function

'return table attributes by correctly table offset
'if success this function return value >= 0
'if failed return value is -1
Public Property Get TableAttributes(ByVal lOffsetTable As Long) As MobDbAttributes
  Dim table As table_type_2
  TableAttributes = -1
  If lOffsetTable < 1 Then Exit Property
  Get #FF, lOffsetTable, table
  If table.sign = SIGN_DB Then TableAttributes = table.attributes
End Property

'write table attributes by correctly table offset
Public Property Let TableAttributes(ByVal lOffsetTable As Long, ByVal vNewValue As MobDbAttributes)
  Dim table As table_type_2
  If lOffsetTable < 1 Then Exit Property
  Get #FF, lOffsetTable, table
  If table.sign = SIGN_DB Then
    table.attributes = vNewValue
    Put #FF, lOffsetTable, table
  End If
End Property

'forward table name by offset of table
'
'âîçâðàùàåò èìÿ òàáëèöû ïî åå àäðåñó
Function getTableName(ByVal lTableOffset As Long) As String
  Dim i As Long
  For i = 0 To lTableCount - 1
    If tables(i).lOffset = lTableOffset Then
      getTableName = tables(i).Key
      Exit Function
    End If
  Next
End Function

'return offset for first field in table
'if success this function return value > 0
'if failed return value is 0
Function getFirstField(ByVal lOffsetTable As Long) As Long
  Dim table As table_type_2
  Get #FF, lOffsetTable, table
  If table.sign = SIGN_DB Then getFirstField = table.field
End Function

'return field name by correctly field offset
Function getFieldName(ByVal lOffsetField As Long) As String
  Dim field As field_type_2
  If lOffsetField < 0 Then Exit Function
  Get #FF, lOffsetField, field
  If field.sign = SIGN_DB Then getFieldName = ByteToString(field.szName)
End Function

'return field type bytemask by offset of field
'if success this function return value >= 0
'if failed return value is -1
'
'âîçâðàùàåò áèòîâóþ ìàñêó òèïà ïîëÿ
Function getFieldType(ByVal lFieldOffset As Long) As DBFieldType
  Dim field As field_type_2
  getFieldType = -1
  If lFieldOffset < 1 Then Exit Function
  Get #FF, lFieldOffset, field
  If field.sign = SIGN_DB Then getFieldType = field.Type
End Function

'this return true if lRowOffset is correct row offset
'
'âîçâðàùàåòñÿ true, åñëè ïî äàííîìó àäðåñó ðàñïîëàãàåòñÿ
'êîððåêòíàÿ ñòðîêà òàáëèöû
Function RowExists(ByVal lRowOffset As Long) As Boolean
  Dim row As row_type_2
  Dim Data As data_type_2
  If lRowOffset < 1 Then Exit Function
  Get #FF, lRowOffset, row
  If row.sign = SIGN_ROW Then
    RowExists = True
  Else
    If row.Data > 0 Then
      Get #FF, row.Data, Data
      RowExists = Data.sign = SIGN_DATA
    End If
  End If
End Function

'comments read from function name =)
'if success this function return value > 0
'if failed return value is 0
Function getFirstTableRow(ByVal lOffsetTable As Long) As Long
  Dim table As table_type_2
  If lOffsetTable < 1 Then Exit Function
  Get #FF, lOffsetTable, table
  If table.sign = SIGN_DB Then getFirstTableRow = table.record
End Function

'comments read from function name =)
'if success this function return value > 0
'if failed return value is 0
Function getLastTableRow(ByVal lOffsetTable As Long) As Long
  Dim table As table_type_2
  If lOffsetTable < 1 Then Exit Function
  Get #FF, lOffsetTable, table
  If table.sign = SIGN_DB Then getLastTableRow = table.lastrecord
End Function

'return next row offset for any correctly row offset
'if success this function return value > 0
'if failed return value is 0
Function getNextRow(ByVal lOffsetRow As Long) As Long
  Dim row As row_type_2
  If lOffsetRow < 1 Then Exit Function
  Get #FF, lOffsetRow, row
  getNextRow = row.next
End Function

'return previous row offset for any correctly row offset
'if success this function return value > 0
'if failed return value is 0
Function getPreviousRow(ByVal lOffset As Long)
  Dim row As row_type_2
  Get #FF, lOffset, row
  If row.sign = SIGN_ROW Then getPreviousRow = row.previous
End Function

'flush database file
Public Sub Update()
  If FF <> 0 Then Put #FF, 1, DB
  Call flushDB
End Sub

'return count of records in table by table offset
'if success this function return value >= 0
'if failed return value is -1
Property Get RecordCount(ByVal lOffsetTable As Long)
  Dim table As table_type_2
  RecordCount = -1
  If lOffsetTable < 1 Then Exit Property
  Get #FF, lOffsetTable, table
  If table.sign = SIGN_DB Then RecordCount = table.RecordCount
End Property

'add new row to table
'if success this function return row offset value > 0
'if failed return value is 0
Function Add(ByVal lOffsetTable As Long) As Long
  
  'if database file isa not open
  If FF = 0 Then Exit Function
  
  Dim lOffsetLast As Long
  Dim lOffset As Long
  Dim table As table_type_2
  Dim row As row_type_2
  Dim row_last As row_type_2

  If lOffsetTable < 1 Then Exit Function
  
  Get #FF, lOffsetTable, table
  
  If table.field < 1 Then Exit Function

  lOffsetLast = table.lastrecord
  lOffset = LOF(FF) + 1
  
  table.RecordCount = table.RecordCount + 1

  If lOffsetLast < 1 Then
    table.record = lOffset
    table.record = lOffset
    table.lastrecord = lOffset
  Else
    Get #FF, lOffsetLast, row_last
    row_last.next = lOffset
    Put #FF, lOffsetLast, row_last
    row.previous = lOffsetLast
    table.lastrecord = lOffset
  End If

  row.sign = SIGN_ROW

  Put #FF, lOffset, row
  Put #FF, lOffsetTable, table
  Add = lOffset

End Function

'return field typename by offset of field
Function getFieldTypeName(ByVal lType As DBFieldType) As String
  Select Case lType
    Case fldString: getFieldTypeName = "string"
    Case fldLong: getFieldTypeName = "long"
    Case fldDouble: getFieldTypeName = "double"
    Case fldByte: getFieldTypeName = "byte"
    Case fldDate: getFieldTypeName = "date"
    Case fldInteger: getFieldTypeName = "integer"
    Case fldByteArray: getFieldTypeName = "bytearray"
    Case fldJpegFileBytes: getFieldTypeName = "jpegfilebytearray"
    Case Else: getFieldTypeName = "unknown"
  End Select
End Function

'return tablecount of current database
Property Get TableCount() As Long
  TableCount = lTableCount
End Property

'private function for maximum speed
Private Sub ReadFields()

  Dim i As Long
  Dim lOffset As Long
  Dim lTableOffset As Long
  Dim table As table_type_2
  Dim field As field_type_2

  For i = 0 To TableCount - 1
    With tables(i)
      lTableOffset = .lOffset
      If lTableOffset > 0 Then
        Get #FF, lTableOffset, table
        lOffset = table.field
        Do While lOffset > 0
          ReDim Preserve .Fields(.fieldscount)
          Get #FF, lOffset, field
          .Fields(.fieldscount).Key = ByteToString(field.szName)
          .Fields(.fieldscount).Type = field.Type
          .Fields(.fieldscount).lOffset = lOffset
          .fieldscount = .fieldscount + 1
          lOffset = field.next
        Loop
      End If
    End With
  Next

End Sub

'return next field offset for previous field offset
'if success this function return value > 0
'if failed return value is 0
Function getNextField(ByVal lOffset As Long) As Long
  Dim field As field_type_2
  Get #FF, lOffset, field
  If field.sign = SIGN_DB Then getNextField = field.next
End Function

'return field offset by table offset and field name
'if success this function return value > 0
'if failed return value is 0
Function getFieldOffset(ByVal lOffsetTable As Long, ByVal szFieldName As String) As Long
  Dim i As Long
  Dim j As Long
  For i = 0 To lTableCount - 1
    If tables(i).lOffset = lOffsetTable Then
      For j = 0 To tables(i).fieldscount - 1
        If tables(i).Fields(j).Key = szFieldName Then
          getFieldOffset = tables(i).Fields(j).lOffset
        End If
      Next
      Exit Function
    End If
  Next
End Function

'add field to table
'if success this function return correctly field offset value > 0
'if failed return value is 0
Function addField(ByVal lTableOffset As Long, ByVal szName As String, ByVal lType As DBFieldType) As Long

  Dim lOffset As Long
  Dim loffsetprevious As Long
  Dim table As table_type_2
  Dim field As field_type_2
  Dim fieldrevious As field_type_2

  If lTableOffset < 1 Then Exit Function
  If getFieldOffset(lTableOffset, szName) > 0 Then Exit Function

  Get #FF, lTableOffset, table
  lOffset = table.field
  
  If lOffset = 0 Then
    lOffset = LOF(FF) + 1
    table.field = lOffset
  Else
    lOffset = LOF(FF) + 1
    loffsetprevious = table.lastfield
    Get #FF, loffsetprevious, fieldrevious
    fieldrevious.next = lOffset
    Put #FF, loffsetprevious, fieldrevious
  End If

  table.lastfield = lOffset
  Put #FF, lTableOffset, table
  addField = lOffset

  With field
    .sign = SIGN_DB
    StringToByte .szName, szName
    .Type = lType
  End With

  Put #FF, lOffset, field
  'Call flushDB

End Function

'return table offset by table name
'if success this function return value > 0
'if failed return value is 0
Function getTableOffset(ByVal szTableName As String) As Long
  Dim i As Long
  For i = 0 To lTableCount - 1
    If tables(i).Key = szTableName Then
      getTableOffset = tables(i).lOffset
      Exit Function
    End If
  Next
End Function

'bonus function
Function EachTables() As String
  Dim i As Long
  Dim j As Long
  Dim Tmp As String
  For i = 0 To lTableCount - 1
    Tmp = Tmp & vbCrLf & _
    CStr(tables(i).lOffset) & vbTab & """" & tables(i).Key & """"
    For j = 0 To tables(i).fieldscount - 1
      Tmp = Tmp & vbCrLf & vbTab & CStr(tables(i).Fields(j).lOffset) & vbTab & """" & tables(i).Fields(j).Key & """," & vbTab & getFieldTypeName(tables(i).Fields(j).Type)
    Next
  Next
  EachTables = Tmp
End Function

'construct collection of tables for maximum navigation speed
Private Sub ReadTables()
  Dim lOffset As Long
  Dim table As table_type_2
  lTableCount = 0
  Erase tables
  lTableCountArraySize = 10
  ReDim tables(lTableCountArraySize)
  If FF = 0 Then Exit Sub
  lOffset = DB.table
  Do While lOffset > 0
    If lOffset = 0 Then Exit Do
    Get #FF, lOffset, table
    If lTableCount + 10 > lTableCountArraySize Then
      lTableCountArraySize = lTableCount * 2 + 10
      ReDim Preserve tables(lTableCountArraySize)
    End If
    With tables(lTableCount)
      .lOffset = lOffset
      .Key = ByteToString(table.szName)
    End With
    lOffset = table.next
    lTableCount = lTableCount + 1
  Loop
  Call ReadFields
End Sub

'add table to database
'if success this function return correctly table offset value > 0
'if failed return value is -1
Function addTable(ByVal szName As String, Optional ByVal mAttributes As MobDbAttributes = 0) As Long
  
  Dim lOffset As Long
  Dim loffsetprevious As Long
  Dim table As table_type_2
  Dim tableprevious As table_type_2
  
  addTable = -1
  If FF = 0 Then Exit Function
  
  If getTableOffset(szName) > 0 Then Exit Function
  lOffset = DB.table
  If lOffset = 0 Then
    lOffset = LOF(FF) + 1
    DB.table = lOffset
    DB.lasttable = lOffset
    Put #FF, 1, DB
  Else
    lOffset = LOF(FF) + 1
    loffsetprevious = DB.lasttable
    Get #FF, loffsetprevious, tableprevious
    tableprevious.next = lOffset
    Put #FF, loffsetprevious, tableprevious
    DB.lasttable = lOffset
    Put #FF, 1, DB
  End If

  addTable = lOffset
  With table
    .sign = SIGN_DB
    .attributes = mAttributes
    StringToByte .szName, szName
  End With

  Put #FF, lOffset, table

  'Put #FF, 1, DB
  'Close #FF
  'Open lpFileName For Binary As #FF
  'or
  'Call flushDB

End Function

'flush database file
Private Sub flushDB()
  'close database file if open
  Call CloseDB
  'reopen database file with reconstruct all collections
  Call OpenDB
End Sub

'close database file if open
Sub CloseDB()
  'clear tables collection
  Erase tables
  lTableCountArraySize = 0
  lTableCount = 0
  If FF <> 0 Then
    Close FF
    FF = 0
  End If
End Sub

'fill byte array with string
Private Sub StringToByte(B() As Byte, ByVal szString As String)
  Dim lLen As Long
  Dim sz() As Byte
  lLen = UBound(B) - LBound(B) + 1
  If lLen < 1 Then Exit Sub
  sz = StrConv(szString, vbFromUnicode)
  ReDim Preserve sz(lLen - 1)
  Call CopyMemory(B(0), sz(0), lLen)
End Sub

'fill string with byte array
Private Function ByteToString(B() As Byte) As String
  Dim lRet As Long
  Dim Tmp As String
  Tmp = StrConv(B, vbUnicode)
  lRet = InStr(Tmp, vbNullChar)
  If lRet > 0 Then Tmp = VBA.Left$(Tmp, lRet - 1)
  ByteToString = Tmp
End Function

'open database function if close
Public Sub OpenDB()
  If FF <> 0 Then Call CloseDB
  FF = FreeFile
  Open lpFileName For Binary As #FF
  Get #FF, 1, DB
  With DB
    'if not correctly db file then create
    'correctly db structure
    If .sign <> SIGN_DB Then
      .sign = SIGN_DB
      Erase .szName
      .table = 0
      .lasttable = 0
      .lastupdate.datetime = Now
      .lastupdate.fieldoffset = 0
      .lastupdate.rowoffset = 0
      .lastupdate.tableoffset = 0
      .uniqueid = 0
      Put #FF, 1, DB
    End If
  End With
  Call ReadTables
End Sub

'get database file
Public Property Get Filename() As String
  Call flushDB
  Filename = lpFileName
  Call OpenDB
End Property

'set database file
Public Property Let Filename(ByVal vNewValue As String)
  Call CloseDB
  lpFileName = vNewValue
  Call OpenDB
End Property

'return value of table by correctly offsets
'if fail then return Empty
Public Property Get Value(ByVal lOffsetTable As Long, _
                          ByVal lOffsetField As Long, _
                          ByVal lOffsetRow As Long) As Variant

  Dim field As field_type_2
  Dim lType As Long
  Dim Data As data_type_2
  Dim lOffsetData As Long
  Dim lSize As Long

  Dim B() As Byte
  Dim lng As Long
  Dim dbl As Double
  Dim bt As Byte
  Dim dt As Date
  Dim it As Integer
  Dim v
  
  'âîçâðàùàåì òèï vbEmpty
  Value = v 'loffsetdata

  If lOffsetField < 1 Then Exit Property
  Get #FF, lOffsetField, field
  If field.sign <> SIGN_DB Then Exit Property
  lType = field.Type

  lOffsetData = getDataByRowOffset(lOffsetRow, lOffsetField)
  If lOffsetData = 0 Then Exit Property

  Get #FF, lOffsetData, Data
  
  If Data.sign <> SIGN_DATA Then Exit Property
  lType = Data.Type

  Select Case lType
  Case fldJpegFileBytes
    If Data.sizereal > 0 Then
      ReDim B(Data.sizereal - 1)
      Get #FF, Data.lOffset, B
    End If
    Value = B
  Case fldByteArray
    If Data.sizereal > 0 Then
      ReDim B(Data.sizereal - 1)
      Get #FF, Data.lOffset, B
    End If
    Value = B
  Case fldString
    Erase B
    If Data.sizereal > 0 Then
      ReDim B(Data.sizereal - 1)
      Get #FF, Data.lOffset, B
    End If
    Value = B
  Case fldLong
    ReDim B(Data.sizereal - 1)
    Get #FF, Data.lOffset, lng
    Value = lng
  Case fldDouble
    ReDim B(Data.sizereal - 1)
    Get #FF, Data.lOffset, dbl
    Value = dbl
  Case fldByte
    ReDim B(Data.sizereal - 1)
    Get #FF, Data.lOffset, bt
    Value = bt
  Case fldDate
    ReDim B(Data.sizereal - 1)
    Get #FF, Data.lOffset, dt
    Value = dt
  Case fldInteger
    ReDim B(Data.sizereal - 1)
    Get #FF, Data.lOffset, it
    Value = it
  End Select

End Property

'write value to table by correctly offsets
Public Property Let Value(ByVal lOffsetTable As Long, _
                          ByVal lOffsetField As Long, _
                          ByVal lOffsetRow As Long, _
                           vNewValue)

  Dim field As field_type_2
  Dim lType As Integer
  Dim Data As data_type_2
  Dim lOffsetData As Long
  Dim lSize As Long
  Dim i As Long
  Dim j As Long

  Dim B() As Byte
  Dim lng As Long
  Dim dbl As Double
  Dim bt As Byte
  Dim dt As Date
  Dim it As Integer

  If lOffsetField < 1 Then Exit Property
  
  Get #FF, lOffsetField, field
  If field.sign <> SIGN_DB Then Exit Property
  lType = field.Type
  
  Select Case lType
  Case fldJpegFileBytes
    B = vNewValue
    lSize = UBound(B) + 1
  Case fldByteArray
    B = vNewValue
    lSize = UBound(B) + 1
  Case fldString
    B = CStr(vNewValue)
    lSize = UBound(B) + 1
  Case fldLong
    lng = CLng(Val(vNewValue))
    lSize = 4
  Case fldDouble
    dbl = CDbl(vNewValue)
    lSize = 8
  Case fldByte
    bt = CByte(Val(vNewValue))
    lSize = 1
  Case fldDate
    dt = CDate(vNewValue)
    lSize = 8
  Case fldInteger
    it = CInt(Val(vNewValue))
    lSize = 2
  End Select

  lOffsetData = getDataByRowOffset(lOffsetRow, lOffsetField)
  Get #FF, lOffsetData, Data

  If Data.lOffset < 1 Then Data.lOffset = LOF(FF) + 1

  If lSize > Data.size Then
    Data.size = lSize
    Data.sizereal = lSize
    Data.lOffset = LOF(FF) + 1
  Else
    Data.sizereal = lSize
  End If

  Select Case lType
    Case fldJpegFileBytes: Put #FF, Data.lOffset, B
    Case fldByteArray:     Put #FF, Data.lOffset, B
    Case fldString:        Put #FF, Data.lOffset, B
    Case fldLong:          Put #FF, Data.lOffset, lng
    Case fldDouble:        Put #FF, Data.lOffset, dbl
    Case fldByte:          Put #FF, Data.lOffset, bt
    Case fldDate:          Put #FF, Data.lOffset, dt
    Case fldInteger:       Put #FF, Data.lOffset, it
  End Select

  Data.sign = SIGN_DATA
  Data.Type = lType

  Put #FF, lOffsetData, Data

  With DB.lastupdate
    .datetime = Now
    .tableoffset = lOffsetTable
    .rowoffset = lOffsetRow
  End With

  RaiseEvent Change(lOffsetTable, lOffsetRow, lOffsetField, lOffsetData)

End Property

'return data offset by correctly offsets
'if success this function return value > 0
'if failed return value is 0
Function getDataByRowOffset(ByVal lOffsetRow As Long, ByVal lOffsetField As Long, _
                            Optional ByVal CreateIfNotExists As Boolean = True) As Long

  Dim row As row_type_2
  Dim lOffsetData As Long
  Dim Data As data_type_2
  Dim loffsetlastdata As Long
  Dim dataadd As data_type_2
  
  If lOffsetRow < 1 Then Exit Function
  Get #FF, lOffsetRow, row

  lOffsetData = row.Data
  
  Do While lOffsetData > 0
    loffsetlastdata = lOffsetData
    Get #FF, lOffsetData, Data
    If Data.fieldoffset = lOffsetField Then
      getDataByRowOffset = lOffsetData
      Exit Function
    End If
    lOffsetData = Data.next
  Loop
  
  If Not CreateIfNotExists Then Exit Function
  
  lOffsetData = LOF(FF) + 1
  dataadd.fieldoffset = lOffsetField
  Put #FF, lOffsetData, dataadd
  If row.Data = 0 Then
    row.Data = lOffsetData
    Put #FF, lOffsetRow, row
  Else
    Get #FF, loffsetlastdata, Data
    Data.next = lOffsetData
    Put #FF, loffsetlastdata, Data
  End If
  getDataByRowOffset = lOffsetData

End Function

'Export table in other file
'Excellent function for database compression
'
'Ýêñïîðò òàáëèöû âî âíåøíèé ôàéë
'Ýòîé æå ôóíêöèåé ìîæíî âîñïîëüçîâàòüñÿ äëÿ óïîðÿäî÷èâàíèÿ(ñæàòèÿ) áàçû äàííûõ
Function ExportTable(ByVal lTableOffset As Long, ByVal szPath As String) As Long

  Dim FF As Long
  Dim pvDB As cMobileDB
  Dim lRow As Long
  Dim lField As Long
  Dim lData As Long
  Dim ltable As Long
  Dim Fields As Collection
  Dim ExportFields As Collection
  Dim lExportRow As Long
  Dim i As Long
  Dim szNewTableName As String
  Dim lCount As Long
  Dim lCountAll As Long
  Dim lRowLast As Long

  Set pvDB = New cMobileDB
  Set Fields = New Collection
  Set ExportFields = New Collection

  pvDB.Filename = szPath
  szNewTableName = Me.getTableName(lTableOffset)

  Call pvDB.addTable(szNewTableName)
  ltable = pvDB.getTableOffset(szNewTableName)
  lField = Me.getFirstField(lTableOffset)
  
  Do While lField > 0
    Call pvDB.addField(ltable, Me.getFieldName(lField), Me.getFieldType(lField))
    Call Fields.Add(lField)
    lField = Me.getNextField(lField)
  Loop
  
  lField = pvDB.getFirstField(ltable)
  Do While lField > 0
    Call ExportFields.Add(lField)
    lField = pvDB.getNextField(lField)
  Loop

  lRow = Me.getFirstTableRow(lTableOffset)
  lRowLast = Me.getLastTableRow(lTableOffset)
  lCountAll = Me.RecordCount(lTableOffset)

  Do While lRow > 0
    RaiseEvent ExportProgress(lCountAll, lCount)
    lExportRow = pvDB.Add(ltable)
    For i = 1 To Fields.Count
      pvDB.Value(ltable, ExportFields(i), lExportRow) = Me.Value(lTableOffset, Fields(i), lRow)
    Next
    Call pvDB.Update
    lCount = lCount + 1
    RaiseEvent ExportProgress(lCountAll, lCount)
    If lRow = lRowLast Then Exit Do
    lRow = Me.getNextRow(lRow)
  Loop

  Call pvDB.CloseDB
  Set pvDB = Nothing
  Set Fields = Nothing
  Set ExportFields = Nothing
  
  ExportTable = lCount
  
  RaiseEvent ExportComplete
  
End Function

