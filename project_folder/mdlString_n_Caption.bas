Attribute VB_Name = "mdlString_n_Caption"
Option Explicit

Public Const DIALOG_STRING_OFFSET As Long = 7000 'needs to be added not disjoined bitwise (OR'd)
Public Const UPPER_DELIM = ";"
Public Const LOWER_DELIM = ","

'Field Type String
Public Const FTS_TEXT As String = "text"
Public Const FTS_INTEGER As String = "integer"
Public Const FTS_FLOAT As String = "floating"

'Needs a reboot. Complex implemetation using LIKE operator in development
Private Const CHRSET_TEXT As String = ""
Private Const CHRSET_INTEGER As String = "0123456789-" '"CHARSET{0123456789-}LIKE{-[0-9]}"
Private Const CHRSET_FLOAT As String = "0123456789.-"

Public Const FA_LEFT As String = "left"
Public Const FA_CENTER As String = "center"
Public Const FA_RIGHT As String = "right"

Public Enum eBUTTON_STRING
    DBTN_OK = DIALOG_STRING_OFFSET + vbOK
    DBTN_CANCEL = DIALOG_STRING_OFFSET + vbCancel
    DBTN_ABORT = DIALOG_STRING_OFFSET + vbAbort
    DBTN_RETRY = DIALOG_STRING_OFFSET + vbRetry
    DBTN_IGNORE = DIALOG_STRING_OFFSET + vbIgnore
    DBTN_YES = DIALOG_STRING_OFFSET + vbYes
    DBTN_NO = DIALOG_STRING_OFFSET + vbNo
End Enum

Public Enum eINVALID_FIELD_MESSAGE
    IFM_PROMPTMESSAGE = 1000
    IFM_FIELD = 1001
    IFM_TEXT_SET = 1101
    IFM_INTEGER_RANGE = 1102
    IFM_INTRANGE_SUFFIX = IFM_INTEGER_RANGE + 1
    IFM_FLOAT_RANGE = 1104
    IFM_FLRANGE_SUFFIX = IFM_FLOAT_RANGE + 1
    IFM_LIST_EMPTY = 1201
End Enum

Public Function FontExists(ByVal fontName As String) As Boolean
    Dim objFont As New StdFont
    
    objFont.Name = fontName
    FontExists = StrComp(fontName, objFont.Name, vbTextCompare) = 0
End Function

Public Function GetString_stdButtonName(ByVal ButtonStyle As eBUTTON_STRING) As String
    GetString_stdButtonName = VB.LoadResString(ButtonStyle)
End Function

Public Function GetTypeCharset(ByVal fieldType As String) As String
        Select Case fieldType
        Case FTS_TEXT
            GetTypeCharset = CHRSET_TEXT
        Case FTS_INTEGER
            GetTypeCharset = CHRSET_INTEGER
        Case FTS_FLOAT
            GetTypeCharset = CHRSET_FLOAT
    End Select
End Function

Public Function Get_EnumTypeFromString(ByVal fieldType As String) As String
        Select Case fieldType
        Case FTS_TEXT
            Get_EnumTypeFromString = eFIELD_TYPE.FT_TEXT
        Case FTS_INTEGER
            Get_EnumTypeFromString = eFIELD_TYPE.FT_INTEGER
        Case FTS_FLOAT
            Get_EnumTypeFromString = eFIELD_TYPE.FT_FLOAT
    End Select
End Function

Public Function Get_StringTypeFromEnum(ByVal fieldType As eFIELD_TYPE) As String
    Select Case fieldType
        Case eFIELD_TYPE.FT_TEXT
            Get_StringTypeFromEnum = FTS_TEXT
        Case eFIELD_TYPE.FT_INTEGER
            Get_StringTypeFromEnum = FTS_INTEGER
        Case eFIELD_TYPE.FT_FLOAT
            Get_StringTypeFromEnum = FTS_FLOAT
    End Select
End Function

Public Function Get_InvalidFieldMessage(ByVal stringCode As eINVALID_FIELD_MESSAGE) As String
    Dim retVal As String
    
    Get_InvalidFieldMessage = VB.LoadResString(stringCode)
End Function

Public Function Get_DefaultLabelOnEmpty(ByVal fldText As String, ByVal fldType As String) As String
    Dim retVal As String
    Select Case fldType
        Case FTS_TEXT
            retVal = VB.LoadResString(201)
        Case FTS_INTEGER
            retVal = VB.LoadResString(202)
        Case FTS_FLOAT
            retVal = VB.LoadResString(203)
    End Select
    
    If fldText = "" Then
        Get_DefaultLabelOnEmpty = retVal
    Else
        Get_DefaultLabelOnEmpty = fldText
    End If
End Function

Public Function Get_DefaultComboLabel(ByVal fldText As String) As String
    If fldText = "" Then
        Get_DefaultComboLabel = VB.LoadResString(301)
    Else
        Get_DefaultComboLabel = fldText
    End If
End Function

Public Function Get_AlignString(ByVal alignNum As AlignmentConstants) As String
    Select Case alignNum
        Case AlignmentConstants.vbLeftJustify
            Get_AlignString = FA_LEFT
        Case AlignmentConstants.vbCenter
            Get_AlignString = FA_CENTER
        Case AlignmentConstants.vbRightJustify
            Get_AlignString = FA_RIGHT
    End Select
End Function

Public Function Get_AlignValue(ByVal alignNum As String) As Long
    Select Case alignNum
        Case FA_LEFT
            Get_AlignValue = AlignmentConstants.vbLeftJustify
        Case FA_CENTER
            Get_AlignValue = AlignmentConstants.vbCenter
        Case FA_RIGHT
            Get_AlignValue = AlignmentConstants.vbRightJustify
    End Select
End Function
