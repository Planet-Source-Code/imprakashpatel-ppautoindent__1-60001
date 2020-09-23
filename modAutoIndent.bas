Attribute VB_Name = "modAutoIndent"
'***************************************************************************
'* Copyright (c) 2005 by Prakah Patel
'*
'* This software is the proprietary information of Pd Systems.
'* Use is subject to license terms.
'*
'* @author  Prakash Patel
'* @version 1.0
'* @date    31 March 2004
'*
'***************************************************************************

Option Explicit


Public Sub IndentCode(mCode As CodeModule)
    Dim nIndent As Integer
    Dim nLine As Long
    Dim strNewLine As String
    
    For nLine = 1 To mCode.CountOfLines
        ' Get next line.
        strNewLine = mCode.Lines(nLine, 1)
        ' Remove leading space.
        strNewLine = LTrim$(strNewLine)
        
        If IsBlockEnd(strNewLine) Then nIndent = nIndent - 1
        If nIndent < 0 Then nIndent = 0
        
        ' Put back new line.
        mCode.ReplaceLine nLine, Space$(nIndent * 4) & strNewLine
        
        If IsBlockStart(strNewLine) Then nIndent = nIndent + 1
        
    Next nLine
End Sub

Private Function IsBlockStart(strLine As String) As Boolean
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    
    strTemp = Left$(strLine, nPos)
    
    Select Case strTemp
    Case "With", "For", "Do", "While", "Select", "Case", "Else", "Else:", "#Else", "#Else:", "Sub", "Function", "Property", "Enum", "Type"
        bOK = True
    Case "If", "#If", "ElseIf", "#ElseIf"
        bOK = (Len(strLine) = (InStr(1, strLine, " Then") + 4))
    Case "Private", "Public", "Friend"
        nPos = InStr(1, strLine, " Static ")
        If nPos Then
            nPos = InStr(nPos + 7, strLine, " ")
        Else
            nPos = InStr(Len(strTemp) + 1, strLine, " ")
        End If
        Select Case Mid$(strLine, nPos + 1, InStr(nPos + 1, strLine, " ") - nPos - 1)
        Case "Sub", "Function", "Property", "Enum", "Type"
            bOK = True
        End Select
    End Select
    
    IsBlockStart = bOK
End Function

Private Function IsBlockEnd(strLine As String) As Boolean
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    
    strTemp = Left$(strLine, nPos)
    
    Select Case strTemp
    Case "Next", "Loop", "Wend", "End Select", "Case", "Else", "#Else", "Else:", "#Else:", "ElseIf", "#ElseIf", "End If", "#End If"
        bOK = True
    Case "End"
        bOK = (Len(strLine) > 3)
    End Select
    
    IsBlockEnd = bOK
End Function


