VERSION 5.00
Begin VB.Form frmVBMailOrder 
   Caption         =   "VB Mail Order"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      ToolTipText     =   "Close the form."
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "&Clear"
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      ToolTipText     =   "Clear the form."
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Form"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      ToolTipText     =   "Print the form."
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox chkNewCustomer 
      Caption         =   "New Customer"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtPartNum 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      ToolTipText     =   "Part number of the product in the catalog."
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtPageNum 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      ToolTipText     =   "Page number of the product in the catalog."
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCatCode 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Code of the product from the catalog."
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame fraPayment 
      Caption         =   "Payment Type"
      Height          =   1455
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
      Begin VB.OptionButton optMoneyOrder 
         Caption         =   "Money Order"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optCOD 
         Caption         =   "COD"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optCharge 
         Caption         =   "Charge"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraShipping 
      Caption         =   "Shipping"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
      Begin VB.OptionButton optGround 
         Caption         =   "Ground"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optExpress 
         Caption         =   "Express"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Part Number"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Page Number"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Catalog Code"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmVBMailOrder2.frx":0000
      ToolTipText     =   "Do not click this."
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmVBMailOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB Mail Order Chapter 2
' by Ethan Hampton
' 2/7/2021
' This program prints a mail order form.

Private Sub cmdExit_Click()
    End ' Exit.
End Sub

Private Sub cmdPrint_Click()
    ' Print the form. Commented out to cause less issues.
    ' PrintForm
End Sub

