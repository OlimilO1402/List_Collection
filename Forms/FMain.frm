VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain CCollection"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17550
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   17550
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnColLBoundSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox LstLBound 
      Height          =   6180
      Left            =   8040
      TabIndex        =   13
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton BtnTestLBound 
      Caption         =   "Option Base X"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnObjsSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   13800
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnStrsSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnDecsSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   6135
      Left            =   12720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   11
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton BtnObjsCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   12840
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox LstStrs 
      Height          =   6180
      Left            =   3360
      TabIndex        =   7
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton BtnStrsCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnDecsCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox LstDecs 
      Height          =   6180
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label LblLBound 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   255
      Left            =   8160
      TabIndex        =   15
      Top             =   600
      Width           =   4380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   255
      Left            =   12840
      TabIndex        =   10
      Top             =   600
      Width           =   4380
   End
   Begin VB.Label LblStrs 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   4380
   End
   Begin VB.Label LblDecs 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click to select any element "
      Top             =   600
      Width           =   3105
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_n   As Long
Private m_ColOfDecs As CCollection
Private m_ColOfStrs As CCollection
Private m_ColOfObjs As CCollection
Private m_ColBase   As CCollection

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision & " - " & App.FileDescription
    Randomize Timer
    m_n = 1000
    BtnDecsCreate.ToolTipText = "Create " & m_n & " decimal numbers"
    BtnDecsSort.ToolTipText = "Sort the " & m_n & " decimal numbers"
    BtnStrsCreate.ToolTipText = "Create " & m_n & " names as string"
    BtnStrsSort.ToolTipText = "Sort the " & m_n & " names as string"
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = LstDecs.Left
    Dim t As Single: t = LstDecs.Top
    Dim W As Single: W = LstDecs.Width
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then LstDecs.Move L, t, W, H
    L = L + W: W = LstStrs.Width
    If W > 0 And H > 0 Then LstStrs.Move L, t, W, H
    L = L + W: W = LstLBound.Width
    If W > 0 And H > 0 Then LstLBound.Move L, t, W, H
    L = L + W: W = Me.ScaleWidth - L
    If W > 0 And H > 0 Then Text1.Move L, t, W, H
End Sub

Private Sub BtnDecsCreate_Click()
    Dim mr As VbMsgBoxResult: mr = MsgBox("Create " & m_n & " decimals" & vbCrLf & "Doubleclick one element to edit.", vbInformation Or vbOKCancel)
    If mr = vbCancel Then Exit Sub
    Set m_ColOfDecs = MNew.CCollection(False, Name:="ColOfDecs")
    Dim i As Long
    Dim d: d = CDec(1.23457890123457E+18)
    For i = 0 To m_n - 1
        m_ColOfDecs.Add CDec(CDec(Rnd) * d)
    Next
    m_ColOfDecs.ToListBox LstDecs
    LblDecs.Caption = m_ColOfDecs.ToStr(True)
End Sub
Private Sub BtnDecsSort_Click()
    If m_ColOfDecs Is Nothing Then Exit Sub
    If m_ColOfDecs.Count = 0 Then Exit Sub
    m_ColOfDecs.Sort
    m_ColOfDecs.ToListBox LstDecs
End Sub

Private Sub LstDecs_DblClick()
    Dim li As Long: li = LstDecs.ListIndex
    Dim d: d = m_ColOfDecs.Item(li + 1)
    Dim s As String: s = InputBox("Edit", "Edit", d)
    If StrPtr(s) = 0 Then Exit Sub
    d = CDec(s)
    m_ColOfDecs.Item(li + 1) = d
    LstDecs.List(li) = d
End Sub
Private Sub LblDecs_Click()
    If m_ColOfDecs Is Nothing Then Exit Sub
    If m_ColOfDecs.Count = 0 Then Exit Sub
    Dim s As String: s = InputBox("Index?", "Index", CLng(Rnd * m_n))
    If StrPtr(s) = 0 Then Exit Sub
    Dim i As Long: i = CLng(s)
    Dim d: d = m_ColOfDecs.Item(i)
    MsgBox d
End Sub

Private Sub BtnStrsCreate_Click()
    Dim mr As VbMsgBoxResult: mr = MsgBox("Create " & m_n & " strings." & vbCrLf & "Doubleclick one element to edit.", vbInformation Or vbOKCancel)
    If mr = vbCancel Then Exit Sub
    Set m_ColOfStrs = MNew.CCollection(False, Name:="ColOfStrs")
    Dim i As Long
    Dim nam As String
    For i = 0 To m_n - 1
        nam = GetRandomName
        m_ColOfStrs.Add nam
    Next
    m_ColOfStrs.ToListBox LstStrs
    LblStrs.Caption = m_ColOfStrs.ToStr(True)
End Sub
Private Sub BtnStrsSort_Click()
    If m_ColOfStrs Is Nothing Then Exit Sub
    If m_ColOfStrs.Count = 0 Then Exit Sub
    m_ColOfStrs.Sort
    m_ColOfStrs.ToListBox LstStrs
End Sub
Private Sub LblStrs_Click()
    If m_ColOfStrs Is Nothing Then Exit Sub
    If m_ColOfStrs.Count = 0 Then Exit Sub
    Dim s As String: s = InputBox("Index?", "Index", CLng(Rnd * m_n))
    If StrPtr(s) = 0 Then Exit Sub
    Dim i As Long: i = CLng(s)
    s = m_ColOfStrs.Item(i)
    MsgBox s
End Sub
Private Sub LstStrs_DblClick()
    Dim li As Long: li = LstStrs.ListIndex
    Dim s As String: s = m_ColOfStrs.Item(li + 1)
    s = InputBox("Edit", "Edit", s)
    If StrPtr(s) = 0 Then Exit Sub
    m_ColOfStrs.Item(li + 1) = s
    LstStrs.List(li) = s
End Sub

Private Sub BtnObjsCreate_Click()
    Set m_ColOfObjs = MNew.CCollection(True, Name:="ColOfObjs")
    Dim i As Long
    Dim nam As String
    Dim Obj As CCollection
    For i = 1 To 20
        nam = GetRandomName
        Set Obj = MNew.CCollection(True, , nam)
        m_ColOfObjs.Add Obj, Obj.Name
    Next
    Label1.Caption = m_ColOfObjs.ToStr(True)
    Text1.Text = m_ColOfObjs.Data_ToStr
End Sub

Private Sub BtnObjsSort_Click()
    If m_ColOfObjs Is Nothing Then Exit Sub
    If m_ColOfObjs.Count = 0 Then Exit Sub
    m_ColOfObjs.Sort
    Dim s As String
    'in VBA only the following line will work:
    s = m_ColOfObjs.Data_ToStr
    'in VBC you can also use this:
'    Dim v, Obj As Object
'    For Each v In Col
'        Set Obj = v
'        s = s & Obj.Name & vbCrLf
'    Next
    Text1.Text = s
End Sub

Private Sub TestLBound1()
    Dim col As CCollection: Set col = MNew.CCollection(False, , , -1)
    col.Add 123456
    col.Add 234567
    col.Add 345678
    Dim i As Long
    Dim ll As Long: ll = col.LLBound
    Dim uu As Long: uu = col.UUBound
    Debug.Print "Col.LBound = " & ll & "; Col.UBound = " & uu
    For i = ll To uu
        Debug.Print " i =" & i & " col.Item(i) = " & col.Item(i)
    Next
End Sub

Private Sub BtnTestLBound_Click()
    Dim slb As String: slb = InputBox("Define LBound of CCollection: ", "LBound?", -2)
    If StrPtr(slb) = 0 Then Exit Sub
    If Not IsNumeric(slb) Then MsgBox "Numeric!!": Exit Sub
    Dim lb As Long: lb = CLng(slb)
    Dim n As Long: n = 100
    Set m_ColBase = MNew.CCollection(False, Name:="ColBaseX", OptionBaseLBound:=lb)
    Dim mr As VbMsgBoxResult: mr = MsgBox("Create " & n & " strings. LBound now is " & lb & vbCrLf & "Blick one element to edit.", vbInformation Or vbOKCancel)
    If mr = vbCancel Then Exit Sub
    Dim i As Long
    For i = 1 To n
        m_ColBase.Add GetRandomName
    Next
    LstLBound.Clear
    For i = m_ColBase.LLBound To m_ColBase.UUBound
        LstLBound.AddItem i & vbTab & m_ColBase.Item(i)
    Next
    LblLBound.Caption = m_ColBase.ToStr(True)
End Sub

Private Sub LstLBound_Click()

    If m_ColBase Is Nothing Then Exit Sub
    Dim li As Long: li = LstLBound.ListIndex
    Dim i As Long: i = m_ColBase.LLBound + li
    Dim s As String: s = m_ColBase.Item(i)
    s = InputBox("Edit", "Edit", s)
    If StrPtr(s) = 0 Then Exit Sub
    m_ColBase.Item(i) = s
    LstLBound.List(li) = i & vbTab & s
'
'    If m_ColOfStrs.Count = 0 Then Exit Sub
'    Dim s As String: s = InputBox("Index?", "Index", CLng(Rnd * m_n))
'
'    Dim i As Long: i = m_ColBase.LLBound + LstLBound.ListIndex
'    MsgBox "The element with index " & i & " is:" & vbCrLf & m_ColBase.Item(i)
End Sub

Private Sub BtnColLBoundSort_Click()
    m_ColBase.Sort
    Dim i As Long
    LstLBound.Clear
    For i = m_ColBase.LLBound To m_ColBase.UUBound
        LstLBound.AddItem i & vbTab & m_ColBase.Item(i)
    Next
End Sub


