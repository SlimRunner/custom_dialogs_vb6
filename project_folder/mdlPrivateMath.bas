Attribute VB_Name = "mdlPrivateMath"
Option Explicit

Public Function IsCharsetValid(ByVal sourceText As String, ByVal ruleSet As String) As Boolean
    Dim uniCode() As Byte, i As Long
    Dim retVal As Boolean
    'IsCharsetValid = InStr(1, SlctChrt, Chr(KeyAscii)) <> 0
    
    retVal = True
    
    If ruleSet = "" Then
        IsCharsetValid = retVal
        Exit Function
    End If
    
    uniCode = sourceText
    
    For i = 0 To UBound(uniCode) Step 2
        If InStr(1, ruleSet, Chr(uniCode(i))) = 0 Then _
            retVal = False
    Next i
    
    IsCharsetValid = retVal
End Function

'Takes the type of lowerRange
Public Function IsRangeValid(ByVal sourceText As String, ByVal lowerRange As Variant, ByVal upperRange As Variant) As Boolean
    Dim srcNumber As Double
    
    srcNumber = Val(sourceText)
    
    Select Case VarType(lowerRange)
    Case vbByte, vbInteger, vbLong
        IsRangeValid = CLng(srcNumber) >= lowerRange And CLng(srcNumber) <= upperRange
        
    Case vbSingle, vbDouble
        IsRangeValid = srcNumber >= lowerRange And srcNumber <= upperRange
        
    End Select
End Function

Public Function Get_Greater(ByVal num1 As Long, ByVal num2 As Long) As Long
    If num1 > num2 Then
        Get_Greater = num1
    Else
        Get_Greater = num2
    End If
End Function

Public Function Get_RightLoc(ObjControl As Object) As Long
    Get_RightLoc = ObjControl.Left + ObjControl.width
End Function

Public Function Get_BottomLoc(ObjControl As Object) As Long
    Get_BottomLoc = ObjControl.Top + ObjControl.Height
End Function

Public Function Get_RightMostControl(sourceBox As Form) As Long
    Dim i_each As Object, x As Long, rightMost As Long
    Dim firstIter As Boolean
    
    firstIter = True
    
    For Each i_each In sourceBox.Controls
        x = i_each.Left + i_each.width

        If firstIter Then
            rightMost = x
            firstIter = False
        End If
        
        If rightMost < x And firstIter = False Then
            rightMost = x
        End If
        
    Next i_each
    
    Get_RightMostControl = rightMost
End Function

Public Function Get_LowestControl(sourceBox As Form) As Long
    Dim i_each As Object, y As Long, lowest As Long
    Dim firstIter As Boolean
    
    firstIter = True
    
    For Each i_each In sourceBox.Controls
        y = i_each.Top + i_each.Height
        
        If firstIter Then
            lowest = y
            firstIter = False
        End If
        
        If lowest < y And firstIter = False Then
            lowest = y
        End If
        
    Next i_each
    
    Get_LowestControl = lowest
End Function
