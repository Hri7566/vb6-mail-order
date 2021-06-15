VERSION 5.00
Begin VB.Form frmSummary 
   Caption         =   "Summary"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   Picture         =   "frmSummary.frx":0000
   ScaleHeight     =   2790
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Summary"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame fraSummary 
      Caption         =   "Summary"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sales tax"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Shipping and handling"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total amount due"
         Height          =   195
         Left            =   555
         TabIndex        =   12
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label lblSalesTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label lblShipping 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblDollarAmount 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dollar amount due"
         Height          =   195
         Left            =   495
         TabIndex        =   7
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label lblDollarAmountTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblTotalTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblShippingTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label lblSalesTaxTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nontaxable"
         Height          =   195
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Taxable"
         Height          =   195
         Left            =   4080
         TabIndex        =   1
         Top             =   360
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
