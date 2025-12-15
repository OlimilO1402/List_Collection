VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14655
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
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnObjsSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnStrsSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   4320
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
      Left            =   8040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   11
      Top             =   960
      Width           =   6615
   End
   Begin VB.CommandButton BtnObjsCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   8040
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
      Left            =   3360
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   600
      Width           =   4500
   End
   Begin VB.Label LblStrs 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   4500
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

Private Sub Form_Load()
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
    L = L + W: W = Me.ScaleWidth - L
    If W > 0 And H > 0 Then Text1.Move L, t, W, H
End Sub

Private Sub BtnDecsCreate_Click()
    Set m_ColOfDecs = MNew.CCollection(False)
    Dim i As Long
    Dim d: d = CDec(1.23457890123457E+18)
    For i = 0 To m_n - 1
        m_ColOfDecs.Add CDec(CDec(Rnd) * d)
    Next
    m_ColOfDecs.ToListBox LstDecs
    LblDecs.Caption = m_ColOfDecs.ToStr
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
    Set m_ColOfStrs = MNew.CCollection(False)
    Dim i As Long
    Dim nam As String
    For i = 0 To m_n - 1
        nam = GetRandomName
        m_ColOfStrs.Add nam
    Next
    m_ColOfStrs.ToListBox LstStrs
    LblStrs.Caption = m_ColOfStrs.ToStr
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
    Set m_ColOfObjs = MNew.CCollection(True, , "Col")
    Dim i As Long
    Dim nam As String
    Dim Obj As CCollection
    For i = 1 To 20
        nam = GetRandomName
        Set Obj = MNew.CCollection(True, , nam)
        m_ColOfObjs.Add Obj, Obj.Name
    Next
    Label1.Caption = m_ColOfObjs.ToStr
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

