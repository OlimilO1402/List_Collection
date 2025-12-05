VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnStrsSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnDecsSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   6735
      Left            =   8040
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   4575
   End
   Begin VB.CommandButton BtnTestNewEnum 
      Caption         =   "TestNewEnum"
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox LstStrs 
      Height          =   6690
      Left            =   3360
      TabIndex        =   4
      Top             =   600
      Width           =   4575
   End
   Begin VB.CommandButton BtnStrsCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnDecsCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox LstDecs 
      Height          =   6690
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label LblStrs 
      Caption         =   "Strings"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   600
   End
   Begin VB.Label LblDecs 
      Caption         =   "Decimals"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Click to select any element "
      Top             =   120
      Width           =   780
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

Private Sub Form_Load()
    Randomize Timer
    m_n = 1000
    BtnDecsCreate.ToolTipText = "Create " & m_n & " decimal numbers"
    BtnDecsSort.ToolTipText = "Sort the " & m_n & " decimal numbers"
    BtnStrsCreate.ToolTipText = "Create " & m_n & " names as string"
    BtnStrsSort.ToolTipText = "Sort the " & m_n & " names as string"
End Sub

Private Sub BtnDecsCreate_Click()
    Set m_ColOfDecs = MNew.CCollection(False)
    Dim i As Long
    Dim d: d = CDec(1.23457890123457E+18)
    For i = 0 To m_n - 1
        m_ColOfDecs.Add CDec(CDec(Rnd) * d)
    Next
    m_ColOfDecs.ToListBox LstDecs
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
    'm_ColOfDecs.ToListBox LstDecs
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
    m_ColOfStrs.ToListBox LstStrs
End Sub

Private Sub BtnTestNewEnum_Click()
    Dim Col As CCollection: Set Col = MNew.CCollection(True)
    Dim i As Long
    Dim nam As String
    For i = 1 To 20
        nam = GetRandomName
        Col.Add nam, nam
    Next
    Col.Sort
    Dim v, s As String
    For Each v In Col
        s = s & CStr(v) & vbCrLf
    Next
    Text1.Text = s
End Sub

