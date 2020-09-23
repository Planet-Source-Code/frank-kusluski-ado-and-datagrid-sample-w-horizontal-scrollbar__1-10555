VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "ADO and DataGrid Example"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Select all that match"
      Height          =   372
      Left            =   2040
      TabIndex        =   37
      Top             =   6240
      Width           =   1692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "Telephone"
      DataSource      =   "Adodc1"
      Height          =   252
      Left            =   4800
      TabIndex        =   35
      Top             =   6120
      Width           =   3012
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   12
      Format          =   "###-###-####"
      Mask            =   "###-###-####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   252
      Left            =   1680
      TabIndex        =   34
      Top             =   3600
      Width           =   492
   End
   Begin VB.TextBox Text5 
      Height          =   288
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   33
      Top             =   3600
      Width           =   372
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Insert New Record"
      Height          =   372
      Left            =   2040
      TabIndex        =   31
      Top             =   5880
      Width           =   1692
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Delete Record"
      Height          =   372
      Left            =   360
      TabIndex        =   30
      Top             =   5880
      Width           =   1692
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Add New Record"
      Height          =   372
      Left            =   2040
      TabIndex        =   29
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Restore w/Rollback"
      Height          =   372
      Left            =   360
      TabIndex        =   28
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Restore w/Cancel"
      Height          =   372
      Left            =   2040
      TabIndex        =   27
      Top             =   5160
      Width           =   1692
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Edit name in rec 2"
      Height          =   372
      Left            =   360
      TabIndex        =   26
      Top             =   5160
      Width           =   1692
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Go bookmark"
      Height          =   372
      Left            =   2040
      TabIndex        =   25
      Top             =   4800
      Width           =   1692
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Set bookmark"
      Height          =   372
      Left            =   360
      TabIndex        =   24
      Top             =   4800
      Width           =   1692
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Reset Sort"
      Height          =   372
      Left            =   7080
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Refresh Records"
      Height          =   372
      Left            =   2040
      TabIndex        =   22
      Top             =   4440
      Width           =   1692
   End
   Begin VB.CommandButton Command13 
      Caption         =   "OK"
      Height          =   252
      Left            =   7920
      TabIndex        =   21
      Top             =   4440
      Width           =   492
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   4800
      TabIndex        =   20
      Text            =   "city ASC, name DESC"
      Top             =   4440
      Width           =   3132
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Put records in array"
      Height          =   372
      Left            =   360
      TabIndex        =   18
      Top             =   4440
      Width           =   1692
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Find Next"
      Height          =   372
      Left            =   6000
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Reset Filter"
      Height          =   372
      Left            =   4800
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton Command9 
      Caption         =   "OK"
      Height          =   252
      Left            =   7920
      TabIndex        =   15
      Top             =   4080
      Width           =   492
   End
   Begin VB.CommandButton Command8 
      Caption         =   "OK"
      Height          =   252
      Left            =   7920
      TabIndex        =   14
      Top             =   3720
      Width           =   492
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   4800
      TabIndex        =   13
      Text            =   "name LIKE 's*'"
      Top             =   4080
      Width           =   3132
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   4800
      TabIndex        =   12
      Text            =   "name LIKE 's*' OR name LIKE 't*'"
      Top             =   3720
      Width           =   3132
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   372
      Left            =   600
      TabIndex        =   9
      Top             =   3000
      Value           =   1
      Width           =   3132
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Go Backward a page"
      Height          =   372
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Width           =   1692
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Go Forward a page"
      Height          =   372
      Left            =   2040
      TabIndex        =   7
      Top             =   4080
      Width           =   1692
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">|"
      Height          =   372
      Left            =   3720
      TabIndex        =   6
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton Command2 
      Caption         =   "|<"
      Height          =   372
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   252
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   288
      Left            =   4800
      TabIndex        =   2
      Top             =   5520
      Width           =   3012
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Full View"
      Height          =   372
      Left            =   3840
      TabIndex        =   1
      Top             =   4920
      Width           =   972
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   2160
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Publishers"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2412
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   7932
      _ExtentX        =   13996
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin VB.Label Label3 
      Caption         =   "Data bound masked-edit text box"
      Height          =   252
      Left            =   4800
      TabIndex        =   36
      Top             =   5880
      Width           =   2412
   End
   Begin VB.Label Label7 
      Caption         =   "Page Size"
      Height          =   252
      Left            =   360
      TabIndex        =   32
      Top             =   3600
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "Sort"
      Height          =   252
      Left            =   4320
      TabIndex        =   19
      Top             =   4440
      Width           =   372
   End
   Begin VB.Label Label5 
      Caption         =   "Find"
      Height          =   252
      Left            =   4320
      TabIndex        =   11
      Top             =   4080
      Width           =   612
   End
   Begin VB.Label Label4 
      Caption         =   "Filter"
      Height          =   252
      Left            =   4320
      TabIndex        =   10
      Top             =   3720
      Width           =   612
   End
   Begin VB.Label Label2 
      Height          =   252
      Left            =   4680
      TabIndex        =   4
      Top             =   3000
      Width           =   3612
   End
   Begin VB.Label Label1 
      Caption         =   "Data bound text box"
      Height          =   252
      Left            =   4800
      TabIndex        =   3
      Top             =   5280
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bkMark As Variant
Dim WithEvents objConn As ADODB.Connection
Attribute objConn.VB_VarHelpID = -1
Dim WithEvents objRecd As ADODB.Recordset
Attribute objRecd.VB_VarHelpID = -1
'Dim objComm As ADODB.Command
Dim intPos As Integer
Dim blnFnd As Boolean

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command10_Click()
Adodc1.Recordset.Filter = ""
'Adodc1.Recordset.Filter = adFilterNone 'does same thing
Adodc1.Recordset.Requery  'refresh entire recordset, use resync to refresh a subset
Command10.Visible = False
End Sub

Private Sub Command11_Click()
If Adodc1.Recordset.EOF Then
   Exit Sub
End If
Adodc1.Recordset.Find Text3.Text, 1, adSearchForward, adBookmarkCurrent
intPos = Adodc1.Recordset.AbsolutePosition
If intPos < 1 Then Exit Sub
Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
HScroll1.Value = intPos - 1
End Sub

Private Sub Command12_Click()
Dim arr As Variant
Dim varFlds(0 To 2) As Variant
Dim I As Byte
Dim n As Byte
varFlds(0) = 1  'varFlds(0) = "name"
varFlds(1) = 2
varFlds(2) = 3
arr = Adodc1.Recordset.GetRows(4, adBookmarkFirst, varFlds)
'show contents of arr()
'Note: First dimension of array is column, 2nd is the row
For I = LBound(arr, 2) To UBound(arr, 2)
    For n = LBound(arr, 1) To UBound(arr, 1)
        If IsNull(arr(n, I)) Then  'prevents error: "Invalid use of null"
           MsgBox "null"
        Else
           MsgBox arr(n, I)  'MsgBox "" & arr(n, i)
        End If
    Next
Next
End Sub

Private Sub Command13_Click()
Adodc1.Recordset.Sort = Text4.Text
Command15.Visible = True
End Sub

Private Sub Command14_Click()
Adodc1.Recordset.Requery
intPos = 1
Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
HScroll1.Value = 0
End Sub

Private Sub Command15_Click()
Adodc1.Recordset.Sort = ""
Adodc1.Recordset.Requery
Command15.Visible = False
End Sub

Private Sub Command16_Click()
bkMark = Adodc1.Recordset.Bookmark
End Sub

Private Sub Command17_Click()
Adodc1.Recordset.Bookmark = bkMark
intPos = Adodc1.Recordset.AbsolutePosition
HScroll1.Value = intPos - 1
Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub Command18_Click()
On Error GoTo UpdateErr
objConn.BeginTrans   'begin transaction
Adodc1.Recordset.Move 1, adBookmarkFirst
Adodc1.Recordset!Name = "Bill Gates"
'Adodc1.Recordset.Fields("Name") = "Bill Gates"
'Adodc1.Recordset.Fields(1) = "Bill Gates"
'Adodc1.Recordset!Name.Value = "Bill Gates"
objConn.CommitTrans   'save changes
Exit Sub

UpdateErr:

MsgBox "An error occurred. Changes will be restored!"
objConn.RollbackTrans
End Sub

Private Sub Command19_Click()
Adodc1.Recordset.Move 1, adBookmarkFirst
Adodc1.Recordset!Name = "John Lennon"

MsgBox "Look at the changed record. When you click OK it will revert back to what it was"

Adodc1.Recordset.CancelUpdate

End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveFirst
intPos = 1
Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
HScroll1.Value = 0
End Sub

Private Sub Command20_Click()
objConn.BeginTrans   'begin transaction

objConn.Execute "Update Publishers SET name = 'Paul McCartney' WHERE name = 'Bill Gates'"

'Refresh entire recordset to show changes
Adodc1.Recordset.Requery

MsgBox "Look at the changed record. When you click OK it will revert back to what it was"

objConn.RollbackTrans  'restore data to original state

Adodc1.Recordset.Move 1
'Refresh only the current record to show changes
Adodc1.Recordset.Resync adAffectCurrent, adResyncAllValues

End Sub

Private Sub Command21_Click()
'add blank record
'Adodc1.Recordset.AddNew
'add record and values
Adodc1.Recordset.AddNew "name", "John Doe"
Adodc1.Recordset.AddNew Array("name", "state"), Array("Sally Smith", "NY")
Adodc1.Recordset.Requery
End Sub

Private Sub Command22_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
   Adodc1.Recordset.MoveLast
End If
Label2.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub Command23_Click()
Dim objComm As ADODB.Command
Set objComm = New ADODB.Command
objComm.ActiveConnection = objConn
objComm.CommandText = "INSERT INTO Publishers(name,city) " & _
   "VALUES('Frank Jones','Toronto')"
objComm.CommandType = adCmdText
objComm.Execute
'Move to the last modified record
'Adodc1.Recordset.Move 0, 1
'Refresh only the current record to show changes
Adodc1.Recordset.Resync adAffectCurrent, adResyncAllValues
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.PageSize = Text5.Text
End Sub

Private Sub Command4_Click()
Dim n As Integer
n = 0
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find Text3.Text, 0, adSearchForward, adBookmarkFirst
If blnFnd Then 'clear bookmarks
   Call Command14_Click
   Command4.Caption = "Select all that match"
   blnFnd = False
Else 'add bookmarks
   DataGrid1.SelBookmarks.Add (Adodc1.Recordset.Bookmark)
   Do While True
      Adodc1.Recordset.Find Text3.Text, 1, adSearchForward, adBookmarkCurrent
      If Adodc1.Recordset.EOF Then
         Exit Do
      End If
      DataGrid1.SelBookmarks.Add (Adodc1.Recordset.Bookmark)
      n = n + 1
   Loop
   If n Then
      Call Command2_Click
      Command4.Caption = "Clear all that match"
      blnFnd = True
      MsgBox CStr(n) & " records selected!"
      'MsgBox CStr(DataGrid1.SelBookmarks.Count - 1) & " records selected!"
      'Loop through bookmarks
      'For n = 0 To DataGrid1.SelBookmarks.Count - 1
      '    MsgBox DataGrid1.SelBookmarks.Item(n)  'display rec #
      'Next
   End If
End If
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast
intPos = Adodc1.Recordset.AbsolutePosition
Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
HScroll1.Value = Adodc1.Recordset.RecordCount - 1
End Sub

Private Sub Command6_Click()
Dim intPage As Integer
intPage = Adodc1.Recordset.AbsolutePage
intPage = intPage + 1
If intPage <= Adodc1.Recordset.PageCount Then Adodc1.Recordset.AbsolutePage = intPage
HScroll1.Value = Adodc1.Recordset.AbsolutePosition - 1
Label2.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub Command7_Click()
Dim intPage As Integer
intPage = Adodc1.Recordset.AbsolutePage
intPage = intPage - 1
If intPage > 0 Then Adodc1.Recordset.AbsolutePage = intPage
HScroll1.Value = Adodc1.Recordset.AbsolutePosition - 1
Label2.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub Command8_Click()
Adodc1.Recordset.Filter = Text2.Text
Command10.Visible = True
End Sub

Private Sub Command9_Click()
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find Text3.Text, 0, adSearchForward, adBookmarkFirst
Command11.Visible = True
intPos = Adodc1.Recordset.AbsolutePosition
Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
HScroll1.Value = intPos - 1
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Shift = 0 Then 'Delete key pressed
   'MsgBox DataGrid1.Columns(DataGrid1.Col).DataField
   If DataGrid1.Col = 0 Then 'Prevent column data from being deleted
      MsgBox "Can't delete data in this column."
      KeyCode = 0
   ElseIf DataGrid1.Col = -1 Then  'Entire row or column is selected
      Dim intResponse As Integer
      intResponse = MsgBox("Are you sure you want to delete selected data?", vbYesNo + vbQuestion)
      If intResponse = vbNo Then
         KeyCode = 0
      End If
   End If
End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Len(LastRow) Then
   intPos = Adodc1.Recordset.AbsolutePosition
   If intPos < 1 Then Exit Sub
   HScroll1.Value = intPos - 1
   Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
End If
End Sub

Private Sub Form_Load()
bkMark = Adodc1.Recordset.Bookmark
Label2.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
Text5.Text = Adodc1.Recordset.PageSize
HScroll1.Max = Adodc1.Recordset.RecordCount - 1
Set objConn = Adodc1.Recordset.ActiveConnection 'Get connection object
End Sub

Private Sub HScroll1_Change()
'Adodc1.Recordset.Move HScroll1.Value, 1
intPos = HScroll1.Value + 1
Adodc1.Recordset.AbsolutePosition = intPos
Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub HScroll1_Scroll()
'Adodc1.Recordset.Move HScroll1.Value, 1
intPos = HScroll1.Value + 1
Adodc1.Recordset.AbsolutePosition = intPos
Label2.Caption = "Record " & intPos & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5.Text) Then
   Text5.Text = ""
End If
End Sub
