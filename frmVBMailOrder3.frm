VERSION 5.00
Begin VB.Form frmVBMailOrder 
   Caption         =   "VB Mail Order"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEmployee 
      Caption         =   "Employee"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   3615
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtTotalHours 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name of Employee"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Hours Worked"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Form"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtTotalSales 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblBonus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Employee's Bonus"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   1290
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Store Sales"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmVBMailOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB Mail Order Chapter 3
' by Ethan Hampton
' 2/9/2021
' This program calculates employee bonuses

Private Sub Form_Load()
    ' Disable print
    cmdPrint.Enabled = False
    
    ' Declarations
    Dim curStoreSales As Currency
    Dim intTotalHours As Integer
    Dim curBonus As Currency
End Sub

Private Sub cmdCalc_Click()
    ' Check for empty fields
    If txtTotalSales.Text <> "" Then
        If txtTotalHours.Text <> "" Then
            ' Grab input
            If CCur(txtTotalHours.Text) > 160 Then
                txtTotalHours.Text = "160"
            End If
            curStoreSales = CCur(txtTotalSales.Text)
            intTotalHours = CInt(txtTotalHours.Text)
            
            ' Calculate the given employee's bonus based on total store sales and their total hours worked
            curBonus = (curStoreSales * intTotalHours) * 0.02
            
            ' Display the output
            lblBonus.Caption = FormatCurrency(curBonus)
            
            ' Enable Print button
            cmdPrint.Enabled = True
        End If
    End If
End Sub

Private Sub cmdClear_Click()
    ' Clear text values
    txtName.Text = ""
    txtTotalHours.Text = ""
    
    ' Clear output labels
    lblBonus.Caption = ""
    
    ' Disable print button
    cmdPrint.Enabled = False
End Sub

Private Sub cmdExit_Click()
    ' Exit the form
    End
End Sub

Private Sub cmdPrint_Click()
    ' Print the form. Disabled to cause less issues.
    ' PrintForm
End Sub
