Attribute VB_Name = "mdlSTDErrorHandler"
Option Explicit

Private Const WIDNOWS_DLL_ERROBJECT = &H10000

Private Const LANG_NEUTRAL = &H0
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100

Public Enum eErrLevelConsts
    ELC_SOURCECHAIN = 0
    ELC_TOPLEVEL = 1
End Enum

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, _
                                                                              ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
                                                                              ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Sub RaiseUserError(ByVal usrNumber As Long, ByVal usrDescription As String)
    Err.Raise vbObjectError + usrNumber, "", usrDescription
End Sub

Public Sub RaiseDllError(ByVal lastDllErr As Long)
    Const BUFF_SIZE As Long = 512
    Dim retBuffer As String
    
    'reverves a buffer for the API output
    retBuffer = Space(BUFF_SIZE)
    'Gets the DLL error description
    If FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lastDllErr, LANG_NEUTRAL, retBuffer, BUFF_SIZE, ByVal 0&) = 0 Then
        retBuffer = "Description failed to be retrieved."
    End If
    'trims the buffer and adds a header
    retBuffer = "Windows DLL Error: 0x" & Hex(lastDllErr) & vbCrLf & Replace(Trim$(retBuffer), vbCrLf, "")
    
    Err.Raise vbObjectError + lastDllErr + WIDNOWS_DLL_ERROBJECT, "", retBuffer
End Sub

Public Sub ErrorLevelManager(ByVal errLevel As eErrLevelConsts, ByVal newSource As String, _
                              ByVal errDesc As String, chainErrNum As Long, ByVal chainSource As String)
                              
    Static ProjectName As String
    
    'Catches the source project name
    If ProjectName = "" Then ProjectName = "  [" & App.Title & "]"
    
    Select Case errLevel
    Case eErrLevelConsts.ELC_SOURCECHAIN 're-raises error
        If chainSource = "" Or chainSource = App.Title Then
            'When the error is the first in the chain
            Err.Raise chainErrNum, "• " & newSource & ProjectName
        Else
            'when the error is a link in the chain but not the last (top level)
            Err.Raise chainErrNum, chainSource & vbCrLf & "    › " & newSource & ProjectName, errDesc
        End If
        
    Case eErrLevelConsts.ELC_TOPLEVEL 'shows error on msgbox
        If chainSource <> "" And chainSource <> App.Title Then
            'If the error is the last of a chain of 2 or more
            If chainErrNum < 0 Then
                'Takes care of user defined errors
                MsgBox "Object Error: " & (chainErrNum - vbObjectError) & " + vbObjectError" & _
                       vbCrLf & errDesc & vbCrLf & vbCrLf & chainSource & vbCrLf & "    » " & _
                       newSource & ProjectName, vbCritical Or vbOKOnly, "Error Handler"
            Else
                'Takes care of VB6 standard errors
                MsgBox "Standard Error: " & chainErrNum & vbCrLf & errDesc & vbCrLf & vbCrLf & _
                       chainSource & vbCrLf & "    » " & newSource & ProjectName, vbCritical Or vbOKOnly, "Error Handler"
            End If
        Else
            'if the error is just one
            If chainErrNum < 0 Then
                'Takes care of user defined errors
                MsgBox "Object Error: " & (chainErrNum - vbObjectError) & " + vbObjectError" & _
                       vbCrLf & errDesc & vbCrLf & vbCrLf & "• " & newSource & ProjectName, vbCritical Or vbOKOnly, "Error Handler"
            Else
                'Takes care of VB6 standard errors
                MsgBox "Standard Error: " & chainErrNum & vbCrLf & errDesc & vbCrLf & vbCrLf & _
                       "• " & newSource & ProjectName, vbCritical Or vbOKOnly, "Error Handler"
            End If
        End If
        
    End Select
End Sub
