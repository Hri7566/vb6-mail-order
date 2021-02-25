VERSION 5.00
Begin VB.Form frmVBMailOrder 
   Caption         =   "VB Mail Order"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSummary 
      Caption         =   "Summary"
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   5415
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Taxable"
         Height          =   195
         Left            =   4080
         TabIndex        =   38
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nontaxable"
         Height          =   195
         Left            =   2280
         TabIndex        =   37
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblSalesTaxTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   36
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label lblShippingTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label lblTotalTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblDollarAmountTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dollar amount due"
         Height          =   195
         Left            =   495
         TabIndex        =   32
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label lblDollarAmount 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblShipping 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label lblSalesTax 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total amount due"
         Height          =   195
         Left            =   555
         TabIndex        =   27
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Shipping and handling"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sales tax"
         Height          =   195
         Left            =   1200
         TabIndex        =   25
         Top             =   960
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6960
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   6960
      TabIndex        =   22
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Summary"
      Height          =   495
      Left            =   5640
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdNextItem 
      Caption         =   "&Next Item"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   20
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame fraItem 
      Caption         =   "Item"
      Height          =   1935
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtItemPrice 
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
         Left            =   1200
         TabIndex        =   15
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtItemWeight 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtItemQuantity 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtItemDesc 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Price"
         Height          =   195
         Left            =   660
         TabIndex        =   19
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Weight"
         Height          =   195
         Left            =   510
         TabIndex        =   18
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Quantity"
         Height          =   195
         Left            =   435
         TabIndex        =   17
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Description"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Frame fraCustomer 
      Caption         =   "Customer"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtZIP 
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
         Left            =   960
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&ZIP"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&State"
         Height          =   195
         Left            =   2760
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&City"
         Height          =   195
         Left            =   525
         TabIndex        =   7
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Address"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Name"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmVBMailOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB Mail Order Chapter 4
' by Ethan Hampton
' 2/10/2021
' This program calculates mail orders for items from a catalog

' Declarations
Dim mcurDollarAmount As Currency
Dim mintTotalWeight As Integer
Dim mintTotalItems As Integer
Dim mcurSalesTax As Currency
Dim mcurShipping As Currency
Dim mcurHandling As Currency
Dim mcurTotal As Currency

Private Sub cmdExit_Click()
    ' Exit the form
    End
End Sub

Private Sub cmdNextItem_Click()
    ' Declarations
    Dim intQuantity As Integer
    Dim intWeight As Integer
    Dim curPrice As Currency
    Dim boolGood As Boolean
    
    ' Check the values of the text boxes to make sure they're all right, and if not, display a message box
    boolGood = True ' Start fine and change if anything goes wrong
    If Not IsNumeric(txtItemQuantity.Text) Then ' Check if entered quantity is right
        msg ("Item quantity has to be a number") ' If not, then display a message box
        boolGood = False ' Do not continue to calculate
    End If
    
    If Not IsNumeric(txtItemWeight.Text) Then ' Check if entered weight is right
        msg ("Item weight has to be a number") ' If not, then display a message box
        boolGood = False ' Do not continue to calculate
    End If
    
    If Not IsNumeric(txtItemPrice.Text) Then ' Check if entered price is right
        msg ("Item price has to be a number") ' If not, then display a message box
        boolGood = False ' Do not continue to calculate
    End If
    
    ' If everything was entered correctly, calculate the total dollar amount and weight
    If boolGood Then
        ' Store the quantity, weight, and price
        intQuantity = Int(txtItemQuantity.Text)
        intWeight = Int(txtItemWeight.Text)
        curPrice = CCur(txtItemPrice.Text)
        
        ' Add them to the total base price and weight
        mcurDollarAmount = mcurDollarAmount + (curPrice * intQuantity)
        mintTotalWeight = mintTotalWeight + (intWeight * intQuantity)
        
        ' Display the base amount due
        lblDollarAmount.Caption = FormatCurrency(mcurDollarAmount)
        lblDollarAmountTax.Caption = FormatCurrency(mcurDollarAmount)
        
        ' Reset the item text boxes
        txtItemDesc.Text = ""
        txtItemQuantity.Text = ""
        txtItemWeight.Text = ""
        txtItemPrice.Text = ""
    End If
End Sub

Private Sub msg(ByVal str As String)
    ' This function just makes it easier to generate message boxes
    MsgBox str, vbOKOnly, "Message"
End Sub

Private Sub cmdPrint_Click()
    ' Print the form (disabled)
    ' PrintForm
End Sub

Private Sub cmdUpdate_Click()
    ' Declarations
    Dim curShippingCharge As Currency
    Dim curHandlingCharge As Currency
    
    ' Calculate shipping charge
    curShippingCharge = FormatCurrency(0.25 * mintTotalWeight)
    
    ' Calculate handling charge
    If mintTotalWeight < 10 Then
        curHandlingCharge = FormatCurrency(1)
    ElseIf mintTotalWeight < 100 Then
        curHandlingCharge = FormatCurrency(3)
    Else
        curHandlingCharge = FormatCurrency(5)
    End If
    
    ' Display the current summary status
    lblDollarAmount.Caption = FormatCurrency(mcurDollarAmount)
    lblSalesTax.Caption = FormatCurrency(0) ' Since this is the nontaxable section, this should be zero
    lblShipping.Caption = FormatCurrency(curShippingCharge + curHandlingCharge)
    lblTotal.Caption = FormatCurrency(mcurDollarAmount + curHandlingCharge + curShippingCharge)
    
    ' Display again in the taxable section
    lblDollarAmountTax.Caption = FormatCurrency(mcurDollarAmount)
    lblSalesTaxTax.Caption = FormatCurrency(mcurDollarAmount * 0.08) ' This time, display it
    lblShippingTax.Caption = FormatCurrency(curShippingCharge + curHandlingCharge)
    lblTotalTax.Caption = FormatCurrency(mcurDollarAmount + (mcurDollarAmount * 0.08) + curHandlingCharge + curShippingCharge)
End Sub

Private Sub txtState_Change()
    ' This limits the input to only have a max of 2 characters
    If Len(txtState.Text) > 2 Then
        txtState.Text = Mid(txtState.Text, 1, 2)
    End If
End Sub

Private Sub checkItemFilled()
    ' This subroutine checks if the current item's parameters are filled out with enough info to determine if the Next Item button should be enabled
    If checkFilled(txtItemQuantity) And checkFilled(txtItemWeight) And checkFilled(txtItemPrice) Then
        cmdNextItem.Enabled = True
    Else
        cmdNextItem.Enabled = False
    End If
End Sub

Private Function checkFilled(ByRef txt As TextBox) As Boolean
    ' This function checks if a text box is filled
    Dim Ret As Boolean
    If Len(txt.Text) = 0 Then
        Ret = False
    Else
        Ret = True
    End If
    checkFilled = Ret
End Function

Private Sub txtItemWeight_Change()
    ' This checks if the item is completed on text change
    Call checkItemFilled
End Sub

Private Sub txtItemQuantity_Change()
    ' This checks if the item is completed on text change
    Call checkItemFilled
End Sub

Private Sub txtItemPrice_Change()
    ' This checks if the item is completed on text change
    Call checkItemFilled
End Sub
