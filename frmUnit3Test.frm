VERSION 5.00
Begin VB.Form frmUnit3Test 
   Caption         =   "Unit 3 Structured Programming Project"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComputerTotal 
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtYourTotal 
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtDiceThree 
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtDiceFour 
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtDiceTwo 
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtDiceOne 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
   End
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   5175
   End
   Begin VB.Label lblComputerTotal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computer Total"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblYourTotal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Your Total"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblComputerDice 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computer Dice Rolls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblYourDice 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Your Dice Rolls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label lblPrompt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click below to generate "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblDiceSimulator 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dice Simulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmUnit3Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Tanav Majithia
'Date: December 16 2021
'Purpose: Unit 3 Test Structured Programming Dice Simulator
Option Explicit

Private Sub cmdGenerate_Click()
'Declaration
Dim intNumber As Integer
Dim intYourDiceOne As Integer
Dim intYourDiceTwo As Integer
Dim intComputerDiceOne As Integer
Dim intComputerDiceTwo As Integer
Dim intYourTotal As Integer
Dim intComputerTotal As Integer
'Initilization
intNumber = 0
intYourDiceOne = 0
intComputerDiceOne = 0
intYourDiceTwo = 0
intComputerDiceTwo = 0
'Processing/Output
Do While intNumber < 10
    intYourDiceOne = Int(6 * Rnd) + 1
    intComputerDiceOne = Int(6 * Rnd) + 1
    intYourDiceTwo = Int(6 * Rnd) + 1
    intComputerDiceTwo = Int(6 * Rnd) + 1
    intYourTotal = intYourDiceOne + intYourDiceTwo
    intComputerTotal = intComputerDiceOne + intComputerDiceTwo
    intNumber = intNumber + 1
Loop
txtComputerTotal = intComputerTotal
txtYourTotal = intYourTotal
'Output
txtDiceOne.Text = (intYourDiceOne)
txtDiceThree.Text = (intComputerDiceOne)
txtDiceTwo.Text = (intYourDiceTwo)
txtDiceFour.Text = (intComputerDiceTwo)
lstOutput.AddItem "First Roll: " & " The Computer has a total of " & intComputerTotal
lstOutput.AddItem "First Roll: " & " Your Total is " & intYourTotal


End Sub

Private Sub Label1_Click()

End Sub

