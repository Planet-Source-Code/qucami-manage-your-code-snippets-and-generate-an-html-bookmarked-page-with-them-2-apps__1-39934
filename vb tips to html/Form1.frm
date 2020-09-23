VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Convert VB Tips file to HTML"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Body"
      Height          =   3195
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   2340
      Width           =   4875
      Begin VB.Frame Frame3 
         Caption         =   "Code"
         Height          =   1335
         Left            =   180
         TabIndex        =   18
         Top             =   1680
         Width           =   4455
         Begin VB.CheckBox Check2 
            Caption         =   "Bold"
            Height          =   195
            Left            =   3360
            TabIndex        =   21
            Top             =   960
            Width           =   675
         End
         Begin VB.ComboBox cmbTipFontSize 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   20
            Text            =   "Combo1"
            Top             =   540
            Width           =   630
         End
         Begin VB.ComboBox cmbTipFonts 
            Height          =   315
            Left            =   420
            Sorted          =   -1  'True
            TabIndex        =   19
            Text            =   "Combo1"
            Top             =   540
            Width           =   2835
         End
         Begin VB.Label Label4 
            Caption         =   "Size"
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   23
            Top             =   300
            Width           =   435
         End
         Begin VB.Label Label3 
            Caption         =   "Font"
            Height          =   195
            Index           =   2
            Left            =   420
            TabIndex        =   22
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Code Title"
         Height          =   1335
         Left            =   180
         TabIndex        =   12
         Top             =   240
         Width           =   4455
         Begin VB.ComboBox cmbTipTitleFonts 
            Height          =   315
            Left            =   420
            Sorted          =   -1  'True
            TabIndex        =   15
            Text            =   "Combo1"
            Top             =   540
            Width           =   2835
         End
         Begin VB.ComboBox cmbTipTitleFontSize 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   14
            Text            =   "Combo1"
            Top             =   540
            Width           =   630
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Bold"
            Height          =   195
            Left            =   3360
            TabIndex        =   13
            Top             =   960
            Value           =   1  'Checked
            Width           =   675
         End
         Begin VB.Label Label3 
            Caption         =   "Font"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   17
            Top             =   300
            Width           =   435
         End
         Begin VB.Label Label4 
            Caption         =   "Size"
            Height          =   195
            Index           =   1
            Left            =   3360
            TabIndex        =   16
            Top             =   300
            Width           =   435
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Title"
      Height          =   1635
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   660
      Width           =   4875
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   180
         TabIndex        =   25
         Text            =   "Visual Basic Tips"
         Top             =   1200
         Width           =   4395
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Bold"
         Height          =   195
         Left            =   3960
         TabIndex        =   10
         Top             =   540
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.ComboBox cmbTitleSize 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   540
         Width           =   630
      End
      Begin VB.ComboBox cmbTitleFonts 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   540
         Width           =   2835
      End
      Begin VB.Label Label5 
         Caption         =   "Caption"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Size"
         Height          =   195
         Index           =   0
         Left            =   3120
         TabIndex        =   8
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Font"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO!"
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   300
      Width           =   435
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Text            =   "c:\VB_Tips.htm"
      Top             =   5880
      Width           =   4875
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Text            =   "H:\vb code\vb tips\_mastertips.txt"
      Top             =   300
      Width           =   4875
   End
   Begin VB.Label Label2 
      Caption         =   "HTML file"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tips file"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim fNum As Integer
Dim sFileIn As String
Dim sLine() As String
Dim x As Long
Dim c As Long
Dim sTip() As String
    Command1.Enabled = False
    fNum = FreeFile
    sFileIn = FileText(Text1)
    sLine = Split(sFileIn, vbCrLf)
    ReDim sTip(0)
    For x = 0 To UBound(sLine)
        If Left(UCase(sLine(x)), 5) = "[TIP]" Then
            ReDim Preserve sTip(UBound(sTip) + 1)
            sTip(UBound(sTip) - 1) = Mid(sLine(x), 6)
        End If
    Next x
    Open Text2 For Output As fNum
    Print #fNum, "<html>"
    Print #fNum, "<head>"
    Print #fNum, "<meta http-equiv=" & Chr(34) & "Content-Language" & Chr(34) & " content=" & Chr(34) & "en-gb" & Chr(34) & ">"
    Print #fNum, "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=windows-1252" & Chr(34) & ">"
    Print #fNum, "<title>VB Tips</title>"
    Print #fNum, "</head>"
    Print #fNum, "<body>"
    Print #fNum, "<p><font face=""" & cmbTitleFonts.Text & """ size=""" & _
                 cmbTitleSize.Text & _
                 """>" & "<a name=""Top"">" & _
                 IIf(chkOptions.Value = vbChecked, "<b>", "") & _
                 Text3 & _
                 IIf(chkOptions.Value = vbChecked, "</b>", "") & _
                 "</a></font></p>"

    Print #fNum, "<p><font face=""" & cmbTipTitleFonts.Text & """ size=""" & cmbTipTitleFontSize.Text & """>" & IIf(Check1.Value = vbChecked, "<b>", "")
    For x = 0 To UBound(sTip) - 1
        Print #fNum, "<a href=" & Chr(34) & "#Tip" & Format(x, "0000") & Chr(34) & ">" & sTip(x) & "</a><br>"
    Next x
    Print #fNum, "</p>" & IIf(Check1.Value = vbChecked, "</b>", "") & "</font>"
    c = 0
    For x = 0 To UBound(sLine)
        If Left(sLine(x), 5) = "[TIP]" Then
            Print #fNum, IIf(Check2.Value = vbChecked, "</b>", "") & "</font><br>"
            Print #fNum, _
            "<font face=""" & cmbTipTitleFonts.Text & """ size=""" & cmbTipTitleFontSize.Text & """><b>" & _
            "<a href=""#Top"">Top</a> &nbsp;" & _
            "<a name=" & Chr(34) & "Tip" & Format(c, "0000") & Chr(34) & ">" & _
            IIf(Check1.Value = vbChecked, "<b>", "") & _
            UCase(Mid(sLine(x), 6)) & _
            IIf(Check1.Value = vbChecked, "</b>", "") & _
            "</a></b></font><br>"
            Print #fNum, "<font face=""" & cmbTipFonts.Text & """ size=""" & cmbTipFontSize.Text & """>" & IIf(Check2.Value = vbChecked, "<b>", "")
            c = c + 1
        Else
            Print #fNum, Replace(sLine(x), " ", "&nbsp;") & "<br>"
        End If
    Next x
    Print #fNum, "</font>"
    Print #fNum, "</body>"
    Print #fNum, "</html>"

    Close fNum
    Command1.Enabled = True
End Sub
Function FileText(ByVal filename As String) As String
    Dim handle As Integer
    
    If Len(Dir$(filename)) = 0 Then
        Err.Raise 53   ' File not found
    End If
    
    handle = FreeFile
    Open filename$ For Binary As #handle
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function

Private Sub Form_Load()
Dim x As Integer
    For x = 0 To Printer.FontCount
        If Printer.Fonts(x) > "" Then
            cmbTitleFonts.AddItem Printer.Fonts(x)
            cmbTipTitleFonts.AddItem Printer.Fonts(x)
            cmbTipFonts.AddItem Printer.Fonts(x)
        End If
        cmbTitleFonts.Text = "Arial"
        cmbTipTitleFonts.Text = "Arial"
        cmbTipFonts.Text = "Courier New"
    Next x
    For x = 1 To 9
        cmbTitleSize.AddItem x
        cmbTipTitleFontSize.AddItem x
        cmbTipFontSize.AddItem x
    Next x
    cmbTitleSize.Text = "3"
    cmbTipTitleFontSize.Text = "2"
    cmbTipFontSize.Text = "2"
End Sub
