VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_CmdLine As String
Event Main(argc As Integer, argv As Variant)

Public Property Get sCommandLine() As String
    sCommandLine = m_CmdLine
End Property

Public Property Let sCommandLine(ByVal vNewCmdL As String)
    m_CmdLine = vNewCmdL
    Call PhaseCommandLine(vNewCmdL)
End Property

Private Sub PhaseCommandLine(CmdLine As String)
Dim ArgLst As Variant
Dim sCmd As String
    
    'Make sure we also append on the exe name as well to the parms
    sCmd = App.EXEName & ".exe"
    sCmd = sCmd & " " & CmdLine
    ArgLst = SplitFx(sCmd, " ") 'Build the list up

    RaiseEvent Main(UBound(ArgLst), ArgLst)
    Erase ArgLst
    sCmd = ""
    
End Sub

Private Sub RevChar(lpStr, char As String)
Dim x As Integer
    'Only used to remove a char form a string
    For x = 1 To Len(lpStr)
        If Mid(lpStr, x, 1) = char Then
            p = p & ""
        Else
            p = p & Mid(lpStr, x, 1)
        End If
    Next x
    
    lpStr = p
    p = ""
End Sub

Function SplitFx(Expression As Variant, Delimiter As Variant) As Variant
Dim x As Long
Dim ch As String * 1
Dim iStart As Integer, sTemp As Variant
Dim vHold() As String, iSize As Integer, QuoteFound As Boolean, CheckChr As Integer

    iSize = -1 'We start our array size from zero
    QuoteFound = False
    
    If Not Right(Expression, 1) = sDet Then
        'add Delimiter to end of the string if it's not found
        Expression = Left(Expression, Len(Expression)) & Delimiter
    End If
    
    Do While x < Len(Expression) 'Loop this the expression string
        x = x + 1 'update counter
        
        ch = Mid$(Expression, x, 1) 'Extract a char from the expression
        
        QuoteFound = (ch = Chr(34)) Or (ch = Chr(39)) ' check for strings inside " or '
        
        If QuoteFound Then
            CheckChr = Asc(ch)
            If (ch = Chr(CheckChr)) Then 'if found
                    If InStr(x + 1, Expression, Chr(CheckChr), vbBinaryCompare) = 0 Then
                    'Next quote was not found but first was
                    x = x + 1
                Else
                    x = InStr(x + 1, Expression, Chr(CheckChr), vbBinaryCompare)
                    'Keep looking
                End If
            End If
        End If
        
        If (ch = Delimiter) Then
            'if we have found the Delimiter extract the data
            sTemp = Trim(Mid(Expression, iStart + 1, (x - iStart) - 1))
            RevChar sTemp, Chr(34) 'Get rid of the quotes
            RevChar sTemp, Chr(39) ' Get rid of the quotes
            If Len(sTemp) <> 0 Then
                iSize = iSize + 1 'update array size
                ReDim Preserve vHold(iSize) 'resize the array to hold data
                
                vHold(iSize) = sTemp 'store the data
                sTemp = "" 'free the temp holder
            End If
            
            iStart = x 'Get the next position of out counter
            End If
        Loop
        
        x = 0
       
        If iSize = -1 Then
            'this was added to stop any errors if the string is empty
            'just sends back an array with the size of zero and a nullchar
            ReDim vHold(0)
            vHold(0) = vbNullChar
        End If
        
        SplitFx = vHold ' send the array back
        Erase vHold 'erase temp array
    
End Function

Public Sub ClearArgs()
    m_CmdLine = ""
End Sub
