VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add tip..."
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   3120
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Tip"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   2640
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   4875
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4875
   End
   Begin VB.Label Label2 
      Caption         =   "Code"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Text1 = "" Then
        MsgBox "A description is required", vbExclamation, "Add tip..."
        Text1.SetFocus
        Exit Sub
    End If
    If Text2 = "" Then
        MsgBox "There's no code to add!", vbExclamation, "Add tip..."
        Text2.SetFocus
        Exit Sub
    End If
    
Dim x As Integer
Dim bThere As Boolean
    bThere = False
    For x = 0 To UBound(sTips)
        If sTips(x) > "" Then
            If UCase(Text1) = UCase(GetTipTitle(sTips(x))) Then
                bThere = True
                Exit For
            End If
        End If
    Next x
    If bThere Then
        MsgBox "That description already exists, please choose another", vbExclamation, "Add tip..."
        Text1.SetFocus
        Exit Sub
    End If
    
Dim sIn As String
Dim fNum As Integer
    fNum = FreeFile
    sIn = FileText(App.Path & "\_mastertips.txt")
    sIn = "[TIP]" & Text1 & vbCrLf & Text2 & vbCrLf & sIn
    Open App.Path & "\_mastertips.txt" For Output As fNum
    Print #fNum, sIn
    Close fNum
    If MsgBox("Tip added" & vbCrLf & vbCrLf & "Do you wish to add another?", vbInformation + vbYesNo, "Add tip...") = vbYes Then
        Text1 = ""
        Text2 = ""
    Else
        Form1.LoadList sIn
        Unload Me
    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = Form1.Left + 300
    Me.Top = Form1.Top + 300
    
End Sub
