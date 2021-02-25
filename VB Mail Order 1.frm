VERSION 5.00
Begin VB.Form frmVBMailOrder 
   Caption         =   "VB Mail Order"
   ClientHeight    =   2385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdShipping 
      Caption         =   "Shipping"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOP 
      Caption         =   "Order Processing"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdMarketing 
      Caption         =   "Marketing"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCR 
      Caption         =   "Customer Relations"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblPhone 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone Number"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmVBMailOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB Mail Order Chapter 1
' by Ethan Hampton
' 2/6/2021
' This program displays the name and phone number of different departments of a company.

Private Sub cmdCR_Click()
    ' Display contact information for Customer Relations
    lblName.Caption = "Tricia Mills"
    lblPhone.Caption = "500-1111"
End Sub

Private Sub cmdMarketing_Click()
    ' Display contact information for Marketing
    lblName.Caption = "Michelle Rigner"
    lblPhone.Caption = "500-2222"
End Sub

Private Sub cmdOP_Click()
    ' Display contact information for Order Processing
    lblName.Caption = "Kenna DeVoss"
    lblPhone.Caption = "500-3333"
End Sub

Private Sub cmdShipping_Click()
    ' Display contact information for Shipping
    lblName.Caption = "Eric Andrews"
    lblPhone.Caption = "500-4444"
End Sub

Private Sub cmdExit_Click()
    ' Exit the form and end the process
    End
End Sub
