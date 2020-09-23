VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUSR 
   Caption         =   "Solving XOR problem using neural networks"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   2520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3720
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Train"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   480
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label16 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      Top             =   390
      Width           =   255
   End
   Begin VB.Label Label14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2130
      TabIndex        =   14
      Top             =   3863
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1410
      TabIndex        =   12
      Top             =   2190
      Width           =   3975
   End
   Begin VB.Label Label12 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1410
      TabIndex        =   11
      Top             =   1590
      Width           =   3975
   End
   Begin VB.Label Label11 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1410
      TabIndex        =   10
      Top             =   990
      Width           =   3975
   End
   Begin VB.Label Label10 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1410
      TabIndex        =   9
      Top             =   390
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   810
      TabIndex        =   8
      Top             =   2183
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   330
      TabIndex        =   7
      Top             =   2183
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   330
      TabIndex        =   6
      Top             =   1583
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   810
      TabIndex        =   5
      Top             =   383
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   810
      TabIndex        =   4
      Top             =   1583
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   810
      TabIndex        =   3
      Top             =   983
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   330
      TabIndex        =   2
      Top             =   983
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   330
      TabIndex        =   1
      Top             =   383
      Width           =   255
   End
   Begin VB.Menu menu 
      Caption         =   "Main Menu"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu df 
         Caption         =   "-"
      End
      Begin VB.Menu load 
         Caption         =   "Load Net"
      End
      Begin VB.Menu save 
         Caption         =   "Save Net"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmUSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Private Sub about_Click()
a = MsgBox("Made By : Paras Chopra, paraschopra@lycos.com, http://naramcheez.netfirms.com. Do you want to see the readme", vbYesNo)
If a = vbYes Then
    Shell ("start " & App.Path & "\readme.txt")
End If
End Sub

Private Sub Command1_Click()
Dim inputdata1(2) As Double
Dim outputdata1(1) As Double
Dim inputdata2(2) As Double
Dim outputdata2(1) As Double
Dim inputdata3(2) As Double
Dim outputdata3(1) As Double
Dim inputdata4(2) As Double
Dim outputdata4(1) As Double

'Dim myInput(2) As Double
'Dim myOutput(5) As Double

inputdata1(1) = 1
inputdata1(2) = 0
outputdata1(1) = 1


inputdata2(1) = 0
inputdata2(2) = 1
outputdata2(1) = 1

inputdata3(1) = 1
inputdata3(2) = 1
outputdata3(1) = 0

inputdata4(1) = 0
inputdata4(2) = 0
outputdata4(1) = 0

a = InputBox("How many Itterations?", , "1500")
If a <> vbCancel And a <> "" And IsNumeric(a) = True Then
For i = 1 To CInt(a)
DoEvents
myInput = Array(Int(Rnd + 0.5), Int(Rnd + 0.5)) ' del
myOutput = Array(myInput(1) Xor myInput(2), _
                     myInput(1) And myInput(2), _
                     myInput(1) Or myInput(2), _
                     Abs(myInput(1) > myInput(2)), _
                     Abs(myInput(1) < myInput(2)) _
                  ) ' del
Call SupervisedTrain(myInput, myOutput) 'del
'Call SupervisedTrain(inputdata1, outputdata1)
'Call SupervisedTrain(inputdata2, outputdata2)
'Call SupervisedTrain(inputdata3, outputdata3)
'Call SupervisedTrain(inputdata4, outputdata4)
Label14 = Int((i / a) * 100) & "%"
Next i
MsgBox "Training complete"
End If
End Sub

Private Sub Command2_Click()

'Dim x(5) As Double

x = Run(Array(1, 0))
Label10 = x(1)
Label9 = Round(x(1), 0)
x = Run(Array(1, 1))
Label11 = x(1)
Label15 = Round(x(1), 0)
x = Run(Array(0, 1))
Label12 = x(1)
Label16 = Round(x(1), 0)
x = Run(Array(0, 0))
Label13 = x(1)
Label17 = Round(x(1), 0)


'del frm here

Debug.Print "Printing the results"

x = Run(Array(1, 0))
Debug.Print "Results of 1 Xor 0 : " & Round(x(1))
Debug.Print "Results of 1 And 0 : " & Round(x(2))
Debug.Print "Results of 1 Or 0 : " & Round(x(3))
Debug.Print "Results of 1 > 0 : " & Round(x(4))
Debug.Print "Results of 1 < 0 : " & Round(x(5))
Debug.Print

x = Run(Array(0, 1))
Debug.Print "Results of 0 Xor 1 : " & Round(x(1))
Debug.Print "Results of 0 And 1 : " & Round(x(2))
Debug.Print "Results of 0 Or 1 : " & Round(x(3))
Debug.Print "Results of 0 > 1 : " & Round(x(4))
Debug.Print "Results of 0 < 1 : " & Round(x(5))
Debug.Print

x = Run(Array(1, 1))
Debug.Print "Results of 1 Xor 1 : " & Round(x(1))
Debug.Print "Results of 1 And 1 : " & Round(x(2))
Debug.Print "Results of 1 Or 1 : " & Round(x(3))
Debug.Print "Results of 1 > 1 : " & Round(x(4))
Debug.Print "Results of 1 < 1 : " & Round(x(5))
Debug.Print

x = Run(Array(0, 0))
Debug.Print "Results of 0 Xor 0 : " & Round(x(1))
Debug.Print "Results of 0 And 0 : " & Round(x(2))
Debug.Print "Results of 0 Or 0 : " & Round(x(3))
Debug.Print "Results of 0 > 0 : " & Round(x(4))
Debug.Print "Results of 0 < 0 : " & Round(x(5))
Debug.Print
End Sub

Private Sub Command3_Click()
Call SaveNet("d:\paras\temp.net")
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
Call CreateNet(2.5, Array(2, 7, 5))
End Sub

Private Sub Form_Unload(Cancel As Integer)
EraseNetwork
End Sub

Private Sub load_Click()
Dlg.Filter = "Neural nets |*.nn"
Dlg.ShowOpen
If Dlg.FileName <> "" Then
    LoadNet (Dlg.FileName)
    Command2_Click
End If
End Sub

Private Sub save_Click()
Dlg.Filter = "Neural nets |*.nn"
Dlg.ShowSave
If Dlg.FileName <> "" Then
   SaveNet (Dlg.FileName)
End If
End Sub
