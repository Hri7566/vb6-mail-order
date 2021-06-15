VERSION 5.00
Begin VB.Form frmVBMailOrder 
   Caption         =   "Form1"
   ClientHeight    =   2610
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   2400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
      Begin VB.Label lblShippingRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Shipping rate"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraInput 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtZone 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtWeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Zone"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight"
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmVBMailOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub msg(ByVal str As String)
    ' This function just makes it easier to generate message boxes
    MsgBox str, vbOKOnly, "Message"
End Sub

Private Sub checkAllFilled()
    Dim boolContinue As Boolean
    boolContinue = False
    
    If checkFilled(txtWeight) And checkFilled(txtZone) Then
        boolContinue = True
    End If
    
    If boolContinue = True Then
        cmdCalculate.Enabled = True
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

Private Sub cmdCalculate_Click()
    Dim curWeight As Currency
    Dim strZone As String
    Dim curRate As Currency
    Dim boolContinue As Boolean
    boolContinue = True
    
    If Not IsNumeric(txtWeight.Text) Then
        msg ("Weight must be a number.")
        boolContinue = False
    End If
    
    If LCase(txtZone.Text) <> "a" And LCase(txtZone.Text) <> "n" And LCase(txtZone.Text) <> "c" And LCase(txtZone.Text) <> "d" Then
        msg ("Zone can only be A, B, C, or D.")
        boolContinue = False
    End If
    
    If boolContinue Then
        curWeight = CCur(txtWeight.Text)
        If curWeight < 0 Then
            curWeight = 0
        End If
        
        strZone = LCase(txtZone.Text)
        
        Select Case curWeight
            Case Is <= 1
                Select Case strZone
                    Case Is = "a"
                        curRate = 1#
                    Case Is = "b"
                        curRate = 1.5
                    Case Is = "c"
                        curRate = 1.65
                    Case Is = "d"
                        curRate = 1.85
                End Select
            Case Is <= 3
                Select Case strZone
                    Case Is = "a"
                        curRate = 1.58
                    Case Is = "b"
                        curRate = 2#
                    Case Is = "c"
                        curRate = 2.4
                    Case Is = "d"
                        curRate = 3.05
                End Select
            Case Is <= 5
                Select Case strZone
                    Case Is = "a"
                        curRate = 1.71
                    Case Is = "b"
                        curRate = 2.52
                    Case Is = "c"
                        curRate = 3.1
                    Case Is = "d"
                        curRate = 4#
                End Select
            Case Is <= 10
                Select Case strZone
                    Case Is = "a"
                        curRate = 2.04
                    Case Is = "b"
                        curRate = 3.12
                    Case Is = "c"
                        curRate = 4#
                    Case Is = "d"
                        curRate = 5.01
                End Select
            Case Else
                Select Case strZone
                    Case Is = "a"
                        curRate = 2.52
                    Case Is = "b"
                        curRate = 3.75
                    Case Is = "c"
                        curRate = 5.1
                    Case Is = "d"
                        curRate = 7.25
                End Select
        End Select
        
        lblShippingRate.Caption = curRate
    End If
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub txtWeight_Change()
    Call checkAllFilled
End Sub

Private Sub txtZone_Change()
    Call checkAllFilled
End Sub
