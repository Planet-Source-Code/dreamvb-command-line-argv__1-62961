VERSION 5.00
Begin VB.Form frmArgTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Arguments Test"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   180
      TabIndex        =   1
      Top             =   105
      Width           =   2910
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   465
      Left            =   225
      TabIndex        =   0
      Top             =   2460
      Width           =   840
   End
End
Attribute VB_Name = "frmArgTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A Little more simpler way to get command line parms for
' Your VB programs
' By Dreamvb

Private WithEvents CmdArgv As clsMain
Attribute CmdArgv.VB_VarHelpID = -1

Sub CmdArgv_Main(argc As Integer, argv As Variant)

    'argc returns the amount of Arguments
    'argv we use the hold the Arguments
    
    'You can also use Arguments with quotes take the following example
    ' also you can use two kinds of quotes " and '
    
    ' CmdArg.exe "he llo" "world" would return 3 Arguments
    ' argv(0) = prog.exe
    ' argv(1) = he llo
    ' argv(2) = world
    
    ' CmdArg.exe he llo world would return 4
    ' argv(0) = prog.exe
    ' argv(1) = he
    ' argv(2) = llo
    ' argv(3) = world
    
    'Test demo
    For x = 0 To argc
        List1.AddItem argv(x)
    Next
    
End Sub

Private Sub Command2_Click()
    Unload frmArgTest
End Sub

Private Sub Form_Load()
    Set CmdArgv = New clsMain
    CmdArgv.sCommandLine = Command$
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmArgTest = Nothing
End Sub
