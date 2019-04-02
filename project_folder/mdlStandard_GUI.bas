Attribute VB_Name = "mdlStandard_GUI"
Option Explicit

'Icons
Private Const IDI_HAND As Long = 32513&
Private Const IDI_ERROR As Long = IDI_HAND
Private Const IDI_QUESTION As Long = 32514&
Private Const IDI_EXCLAMATION As Long = 32515&
Private Const IDI_ASTERISK As Long = 32516&
Private Const IDI_INFORMATION As Long = IDI_ASTERISK

'MessageBeep
Private Const MB_ICONHAND As Long = &H10&
Private Const MB_ICONSTOP As Long = MB_ICONHAND
Private Const MB_ICONERROR As Long = MB_ICONHAND
Private Const MB_ICONQUESTION As Long = &H20&
Private Const MB_ICONEXCLAMATION As Long = &H30&
Private Const MB_ICONWARNING As Long = MB_ICONEXCLAMATION
Private Const MB_ICONASTERISK As Long = &H40&
Private Const MB_ICONINFORMATION As Long = MB_ICONASTERISK

Private Declare Function MessageBeep Lib "user32.dll" (ByVal wType As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long

Public Sub DrawOEM_Icon(ByVal hInstance As Long, dial_icon As eDLG_ICON)
    Dim hIcon As Long 'handle(pointer) to the icon
    
    Select Case dial_icon
        Case DLG_CRITICAL
            hIcon = LoadIcon(ByVal 0&, IDI_ERROR)
        Case DLG_QUESTION
            hIcon = LoadIcon(ByVal 0&, IDI_QUESTION)
        Case DLG_EXCLAMATION
            hIcon = LoadIcon(ByVal 0&, IDI_EXCLAMATION)
        Case DLG_INFORMATION
            hIcon = LoadIcon(ByVal 0&, IDI_INFORMATION)
        Case Else
            Exit Sub
    End Select
    
    Call DrawIcon(hInstance, 0, 0, hIcon)
    Call DestroyIcon(hIcon)
End Sub

Public Sub DialogBeep(ByVal oem_sound As eOEM_SOUND)
    Select Case oem_sound
        Case OEM_BEEP
            Beep
        Case OEM_CRITICAL
            Call MessageBeep(MB_ICONERROR)
        Case OEM_QUESTION
            Call MessageBeep(MB_ICONQUESTION)
        Case OEM_EXCLAMATION
            Call MessageBeep(MB_ICONEXCLAMATION)
        Case OEM_INFORMATION
            Call MessageBeep(MB_ICONINFORMATION)
        Case Else
            Beep
    End Select
End Sub

Public Sub Draw_Stroke(ByRef SourceForm As Form, ByVal x1 As Long, _
                           ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, _
                           ByVal strokeColor As Long, ByVal borderColor As Long)
    SourceForm.Line (x1, y1)-(x2, y2), strokeColor, BF
    SourceForm.Line (x1, y1 - 1)-(x2, y1 - 1), borderColor
End Sub

'Determines if the object being passed is any of the ones passed in checkList
'checkList: Obj1,Obj2,...,Objn
Public Function IsObject(ByRef myObj As Object, ByVal checkList As String) As Boolean
    Dim splitSTR As Variant
    Dim str_each As Variant
    
    splitSTR = Split(checkList, LOWER_DELIM)
    
    For Each str_each In splitSTR
        IsObject = IsObject Or (TypeName(myObj) = str_each)
    Next str_each
End Function

'Fills a list or combo box with items in listText separated by Delimiter
Public Sub PopulateList(ByRef targetList As Object, ByVal listText As String, ByVal withData As Boolean, Optional ByVal Delimiter As String = LOWER_DELIM)
    Dim splitSTR As Variant, i As Long
    
    If IsObject(targetList, "ListBox,ComboBox") Then
        splitSTR = Split(listText, Delimiter)
        
        targetList.Clear
        
        If withData Then
            For i = 0 To UBound(splitSTR) Step 2
                targetList.AddItem splitSTR(i)
                targetList.ItemData(i \ 2) = splitSTR(i + 1)
            Next i
        Else
            For i = 0 To UBound(splitSTR)
                targetList.AddItem splitSTR(i)
                'targetList(i).ItemData = 0
            Next i
        End If
        
    End If
End Sub
