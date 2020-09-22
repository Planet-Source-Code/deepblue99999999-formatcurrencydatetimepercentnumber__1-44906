VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format example by DeepBlue99999999@yahoo.com"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOutput 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Height          =   360
      Left            =   2625
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   5100
      Width           =   3540
   End
   Begin VB.Frame Frame2 
      Caption         =   "Syntax:"
      Height          =   1065
      Left            =   150
      TabIndex        =   2
      Top             =   975
      Width           =   6090
      Begin VB.Label lblSyntax 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   150
         TabIndex        =   21
         Top             =   360
         Width           =   5790
      End
      Begin VB.Label lblSyntax2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         ForeColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   150
         TabIndex        =   22
         Top             =   375
         Width           =   5790
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arguments:"
      Height          =   2790
      Left            =   150
      TabIndex        =   3
      Top             =   2175
      Width           =   6090
      Begin VB.TextBox txtExpression 
         Height          =   390
         Left            =   3075
         TabIndex        =   4
         Text            =   "-123456789"
         Top             =   300
         Width           =   2115
      End
      Begin VB.Frame fraOthers 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   1890
         Left            =   75
         TabIndex        =   11
         Top             =   750
         Width           =   5340
         Begin VB.TextBox txtNumDigits 
            Height          =   360
            Left            =   2985
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "2"
            Top             =   75
            Width           =   390
         End
         Begin VB.ComboBox cboSettings 
            Height          =   360
            Index           =   2
            ItemData        =   "FormatFunctions.frx":0000
            Left            =   2985
            List            =   "FormatFunctions.frx":000F
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1425
            Width           =   1590
         End
         Begin VB.ComboBox cboSettings 
            Height          =   360
            Index           =   1
            ItemData        =   "FormatFunctions.frx":0032
            Left            =   2985
            List            =   "FormatFunctions.frx":0041
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   975
            Width           =   1590
         End
         Begin VB.ComboBox cboSettings 
            Height          =   360
            Index           =   0
            ItemData        =   "FormatFunctions.frx":0064
            Left            =   2985
            List            =   "FormatFunctions.frx":0073
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   525
            Width           =   1590
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "NumDigitsAfterDecimal:"
            Height          =   240
            Left            =   885
            TabIndex        =   15
            Top             =   75
            Width           =   2025
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "GroupDigits:"
            Height          =   240
            Left            =   1860
            TabIndex        =   14
            Top             =   1425
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "UseParensForNegativeNumbers:"
            Height          =   240
            Left            =   150
            TabIndex        =   13
            Top             =   975
            Width           =   2760
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "IncludeLeadingDigit:"
            Height          =   240
            Left            =   1185
            TabIndex        =   12
            Top             =   525
            Width           =   1725
         End
      End
      Begin VB.Frame fraDate 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Caption         =   "Argument:"
         Height          =   1890
         Left            =   150
         TabIndex        =   16
         Top             =   750
         Width           =   5340
         Begin VB.ComboBox cboDateFormats 
            Height          =   360
            ItemData        =   "FormatFunctions.frx":0096
            Left            =   2925
            List            =   "FormatFunctions.frx":00A9
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   75
            Width           =   1815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "NamedFormat:"
            Height          =   240
            Left            =   1575
            TabIndex        =   18
            Top             =   75
            Width           =   1290
         End
      End
      Begin VB.Label lblExpression 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Expression:"
         Height          =   240
         Left            =   2025
         TabIndex        =   10
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdFormatIt 
      Caption         =   "Format it!"
      Default         =   -1  'True
      Height          =   390
      Left            =   2775
      TabIndex        =   1
      Top             =   450
      Width           =   1815
   End
   Begin VB.ComboBox cboFunctions 
      Height          =   360
      ItemData        =   "FormatFunctions.frx":00EE
      Left            =   150
      List            =   "FormatFunctions.frx":00FE
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   450
      Width           =   2415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Formatted output:"
      Height          =   240
      Left            =   975
      TabIndex        =   19
      Top             =   5100
      Width           =   1560
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Select a function:"
      Height          =   240
      Left            =   150
      TabIndex        =   9
      Top             =   150
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Syntax(0 To 4) As String  'To display Syntax
Private GroupDigits As Integer      'One of these values - > vbTrue, vbFalse, vbUseDefault
Private UseParens As Integer         'One of these values - > vbTrue, vbFalse, vbUseDefault
Private LeadingDigit As Integer     'One of these values - > vbTrue, vbFalse, vbUseDefault
Private DateFormat As Integer      'One of these values - > vbGeneralDate, vbLongDate,
                                                         'vbShortDate, vbLongTime, vbShortTime

Private Sub cboDateFormats_Click()
      DateFormat = cboDateFormats.ItemData(cboDateFormats.ListIndex)
End Sub

Private Sub cboFunctions_Click()
      Dim i As Integer
      
      txtOutput.Text = ""  'Clear output.
      
      'Get which function's selected.
      i = cboFunctions.ItemData(cboFunctions.ListIndex)
      
      'Display function's syntax
      lblSyntax.Caption = Syntax(i)
      lblSyntax2.Caption = Syntax(i)  'Shade effect.
      
      If i = 1 Then  'FormatDateTime function
            txtExpression.Text = Now                  'Assign default value
            lblExpression.Caption = "Date:"
            fraDate.Visible = True
            fraOthers.Visible = False
      Else  'Other functions
            'txtExpression.Text = "-123456789"    'Assign default value
            lblExpression.Caption = "Expression:"
            fraDate.Visible = False
            fraOthers.Visible = True
      End If
End Sub

Private Sub cboSettings_Click(Index As Integer)
      Dim i As Integer
      
      i = cboSettings(Index).ListIndex
      
      Select Case Index
      'Include leading digit
      Case 0
            'Get one of these values -> vbTrue, vbFalse, vbUseDefault, -1, 0, -2 respectively.
            LeadingDigit = cboSettings(Index).ItemData(i)
            
      'Use parens for negative numbers
      Case 1
            'Get one of these values -> vbTrue, vbFalse, vbUseDefault, -1, 0, -2 respectively.
            UseParens = cboSettings(Index).ItemData(i)
            
      'Group digits
      Case 2
            'Get one of these values -> vbTrue, vbFalse, vbUseDefault, -1, 0, -2 respectively.
            GroupDigits = cboSettings(Index).ItemData(i)
      End Select
End Sub

Private Sub cmdFormatIt_Click()
      Dim NumDigits As Integer
      Dim result As String
      
      NumDigits = Val(txtNumDigits.Text)
      
      On Error GoTo ErrorHandler
      
      Select Case cboFunctions.ListIndex
      Case 0 'FormatCurrency
            result = FormatCurrency(txtExpression, NumDigits, LeadingDigit, UseParens, GroupDigits)
            
      Case 1 'FormatDateTime
            result = FormatDateTime(txtExpression, DateFormat)
            
      Case 2 'FormatNumber
            result = FormatNumber(txtExpression, NumDigits, LeadingDigit, UseParens, GroupDigits)
            
      Case 3 'FormatPercent
            result = FormatPercent(txtExpression, NumDigits, LeadingDigit, UseParens, GroupDigits)
            
      End Select
      
      txtOutput.Text = result
      Exit Sub
ErrorHandler:
      MsgBox "Something's wrong, check your input again", vbInformation
End Sub

Private Sub Form_Initialize()
      Syntax(0) = "FormatCurrency(Expression,NumDigitsAfterDecimal, IncludeLeadingDigit ,UseParensForNegativeNumbers ,GroupDigits)"
      Syntax(1) = "FormatDateTime(Date, NamedFormat)"
      Syntax(2) = "FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)"
      Syntax(3) = "FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)"
      
      Dim ctl As Control
      
      'Select the first item.
      For Each ctl In Me.Controls
            If TypeName(ctl) = "ComboBox" Then
                  ctl.ListIndex = 0
            End If
      Next ctl
      
      Dim i As Integer
      
      'Populate ItemData.
      For i = 0 To 4
            If i < 4 Then cboFunctions.ItemData(i) = i
            cboDateFormats.ItemData(i) = i
      Next i
      
      LeadingDigit = -1  'vbTrue
      UseParens = -1      'vbTrue
      GroupDigits = -1    'vbTrue
End Sub
