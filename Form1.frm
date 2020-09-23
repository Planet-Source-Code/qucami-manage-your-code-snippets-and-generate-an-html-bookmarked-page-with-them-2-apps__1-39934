VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Tips..."
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   7710
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20108
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      Picture         =   "Form1.frx":0000
      TabIndex        =   7
      ToolTipText     =   "Add new code"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit application"
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   10740
      Picture         =   "Form1.frx":043E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Copy Code"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   10080
      TabIndex        =   4
      ToolTipText     =   "Enter search partial and press [enter]"
      Top             =   1200
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ATM Display"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2040
      Width           =   11295
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      Height          =   1620
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search..."
      Height          =   195
      Left            =   10080
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1800
      Width           =   11295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Clipboard.Clear
    
    Clipboard.SetText IIf(Text1.SelText > "", Text1.SelText, Text1)
    SB.Panels(1).Text = "Code sent to windows clipboard"
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Command3_Click()
    Form2.Show 1
    
End Sub

Private Sub Form_Load()
    ReadTips
End Sub
Sub ReadTips()
Dim sIn As String
Dim sFile
Dim fName As String
    Me.Show
    fName = App.Path & "\_mastertips.txt"
    If FileExists(fName) Then
        sIn = FileText(fName)
    Else
        SB.Panels(1).Text = "Tips file (" & fName & ")not found"
        Exit Sub
    End If

    LoadList sIn
    
    List1.Visible = True
End Sub
Sub LoadList(ByVal sIn As String)
Dim x As Integer
    List1.Clear
    sTips = Split(sIn, "[TIP]")
    sIn = ""
    For x = 0 To UBound(sTips)
        If sTips(x) > "" Then
            List1.AddItem GetTipTitle(sTips(x))
        End If
    Next x
    List1.Refresh
    SB.Panels(1).Text = UBound(sTips) + 1 & " tips loaded"
End Sub


Private Sub List1_Click()
    ShowTip List1.Text
End Sub
Sub ShowTip(ByVal sIn As String)
Dim x As Integer
Dim sTitle As String
Dim sTip As String
    For x = 0 To UBound(sTips)
        If UCase(sIn) = Left(UCase(sTips(x)), Len(sIn)) Then
            sTip = sIn
            sTitle = Mid(sTips(x), InStr(sTips(x), vbCrLf) + 2)
            Text1 = sTitle
            Label1 = sTip
            Exit For
        End If
    Next x
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim x As Integer
    If KeyAscii = 13 Then
        List1.Clear
        For x = 0 To UBound(sTips)
            If sTips(x) > "" Then
                If InStr(UCase(sTips(x)), UCase(Text2)) > 0 Then
                    List1.AddItem GetTipTitle(sTips(x))
                End If
            End If
        Next x
    End If
End Sub
