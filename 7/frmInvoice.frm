VERSION 5.00
Begin VB.Form frmInvoice 
   Caption         =   "Invoice"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboCatalog 
      Height          =   315
      Left            =   120
      MouseIcon       =   "frmInvoice.frx":0000
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
    Debug.Print frmVBMailOrder.txtName.Text
End Sub
