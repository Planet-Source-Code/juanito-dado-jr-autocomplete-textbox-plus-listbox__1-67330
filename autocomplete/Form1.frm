VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      DataSource      =   "Adodc1"
      Height          =   6690
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   9480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API call to listbox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByRef lParam As Any _
) As Long
Const LB_FINDSTRING = &H18F
Dim rs As ADODB.Recordset
Dim DelKey As Boolean
Dim bNoClick As Boolean

Private Sub Form_Load()
Dim a As String

'recordset
Set rs = New ADODB.Recordset
OpenConnection

DoEvents
'query
a = "SELECT sample FROM TblSample"
rs.Open a, adoc, adOpenDynamic, adLockOptimistic
rs.Requery

If rs.RecordCount > 1 Then
    Do Until rs.EOF
        List1.AddItem rs!sample
        rs.MoveNext
    Loop
End If

Set rs = Nothing
End Sub


Private Sub List1_Click()
'to prevent executing the click event
 If bNoClick Then Exit Sub
    Text1.Text = List1.Text
End Sub

Private Sub List1_GotFocus()
    SendKeys "{DOWN}"
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    List1_Click
End Sub

Private Sub Text1_Change()
'autocomplete feature
Dim strt As Long, nIndex As Long
Dim nLen As Long, sText As String
Const LB_GETTEXTLEN As Long = &H18A
Const LB_GETTEXT As Long = &H189
Static blnBusy As Boolean

If blnBusy Then
   Exit Sub
End If
     
     bNoClick = True
     blnBusy = True
    
    'Retrieve the item's listindex
    List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
    
    If Not DelKey Then


    If List1.ListIndex <> -1 Then
        strt = Len(Text1.Text)
        Text1.Text = List1.List(List1.ListIndex)
        Text1.SelStart = strt
        Text1.SelLength = Len(Text1.Text) - strt
    Else
    
    End If
    End If
       DelKey = False
       blnBusy = False
       bNoClick = False

End Sub

Private Sub Text1_Click()
    'select all the text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
    'for delete and backspace
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
    
    DelKey = True
    Exit Sub

    ElseIf KeyCode = vbKeyDown Then
        List1.SetFocus
    End If
End Sub



