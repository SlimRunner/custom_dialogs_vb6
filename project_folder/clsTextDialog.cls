VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains an object with properties to call a custom text input dialog."
Option Explicit

'******************************************************************
'******************************************************************
'   Mario D. Flores (aka SlimRunner or SlimNkName)
'
'   This class creates and manages a custom text dialog that
'   can be programmatically adjusted to different needs by
'   adding at will an arbitrary number of buttons, text boxes,
'   combo boxes, and/or check boxes. It can also display
'   standard system icons as well as play their corresponding
'   sounds.
'
'******************************************************************
'******************************************************************

Private Const DEF_FIELD_SIZE As Long = 120 'px
Private Const DEF_BUTTON_SIZE As Long = 73

Public Enum eFIELD_TYPE
    FT_TEXT = 0
    FT_INTEGER = 1
    FT_FLOAT = 2
    'FT_DATE = 3 'Not implemented yet
    'FT_CURRENCY = 4 'Not implemented yet
    'FT_CUSTOM = 99 'Not implemented yet
End Enum

Public Enum eDLG_RET_VAL
    RET_OK = VbMsgBoxResult.vbOK
    RET_CANCEL = VbMsgBoxResult.vbCancel
    RET_ABORT = VbMsgBoxResult.vbAbort
    RET_RETRY = VbMsgBoxResult.vbRetry
    RET_IGNORE = VbMsgBoxResult.vbIgnore
    RET_YES = VbMsgBoxResult.vbYes
    RET_NO = VbMsgBoxResult.vbNo
End Enum

Public Enum eDLG_BUTTON_STYLE
    DLG_OK_ONLY = 0
    DLG_OK_CANCEL = 1
    DLG_ABORT_RETRY_IGNORE = 2
    DLG_YES_NO_CANCEL = 3
    DLG_YES_NO = 4
    DLG_RETRY_CANCEL = 5
    DLG_CUSTOM = 99
End Enum

Public Enum eDLG_FIELD_STYLE
    DLG_TEXTBOX = 0
    DLG_COMBOBOX = 1
End Enum

Public Enum eDLG_ICON
    DLG_NONE = 0
    DLG_CRITICAL = 1
    DLG_QUESTION = 2
    DLG_EXCLAMATION = 3
    DLG_INFORMATION = 4
End Enum

Public Enum eOEM_SOUND
    OEM_BEEP = 0
    OEM_CRITICAL = 1
    OEM_QUESTION = 2
    OEM_EXCLAMATION = 3
    OEM_INFORMATION = 4
End Enum

Public Enum eCheck_Alignment
    CHKA_LEFT = AlignmentConstants.vbLeftJustify
    CHKA_RIGHT = AlignmentConstants.vbRightJustify
End Enum

Public Type tInputBoxProp
    ibTitle As String
    ibPrompt As String
    ibFont As String
    ibButtonStyle As eDLG_BUTTON_STYLE
    ibFieldSet As Collection
    ibButtonSet As Collection
    ibCheckSet As Collection
    ibIcon As eDLG_ICON
    ibSound As eOEM_SOUND
    ibFieldWidth As Long
    ibSilent As Boolean
    ibDefaultButton As Integer
    ibEnableESC As Boolean
End Type

Public Type tField_Format
    ffType As eDLG_FIELD_STYLE
    ffFlags As String
        'Flags for textboxes
            '<caption>;<type>;<alignment>;<charset/range>;<max_length>;<default_text>
            'type: can be "text" (String), "integer" (Long), "floating" (Double)
            'alignment: can be left, center or right
                'charset: characters accepted to be typed or validated (text)
                'range: "lower,upper" valid range when the field is numeric (numeric)
            'max_length: maximum length of characters that can be typed in the field
            'default_text: text to be preloaded into the text box
                
        'Flags for comboboxes
            '<caption>;<list>;<contains_data>;<default_item>
            'list: "1,2,3,...,n" or "C1,D1,C2,D2,C3,D3,...,Cn,Dn"
            'contains_data: true or false determines if there is item data paired with the list
            'default_item: can be either text (e.g "item") or index (e.g. "/index/123")
        'None of these are case sensitve, and they have a parsing method for both directions
End Type

Public Type tButton_Format
    bfCaption As String
    bfWidth As Long
    bfRetVal As Long
    bfValidate As Boolean
    'Under construction
End Type

Public Type tCheck_Format
    cfCaption As String
    cfCheck_Loc As eCheck_Alignment
    cfDefault_State As CheckBoxConstants
End Type


'Property variables (P: property, F: function, S: sub)
Private dlgTitle As String              'P: InboxTitle
Private dlgPrompt As String             'P: PromptMessage
Private dialFont As String              'P: FontDialog
Private dlgStyle As eDLG_BUTTON_STYLE   'P: DialogStyle
Private fieldSet As Collection          'P/S/S/S: FieldItems, AddFields, DeleteFields, ClearFields
Private buttonSet As Collection         'P/S/S/S: ButtonItems, AddButtons, DeleteButtons, ClearButtons
Private checkSet As Collection          'P/S/S/S: CheckItems, AddChecks, DeleteChecks, ClearChecks
Private dialIcon As eDLG_ICON           'P: IconMessage
Private dialSound As eOEM_SOUND         'P: SoundMessage
Private inputWidth As Long              'P: FieldWidth
Private silentPrompt As Boolean         'P: SilentMode
Private defaultButton As Integer        'S: SetDefaultButton
Private enableCancel As Boolean         'P: EnableCancelKey
Private fieldOutput() As String         'P[ro]/S: OutputFields, ClearReturn
Private checkOutput() As Long           'P[ro]/S: OutputChecks, ClearReturn
Private lastReturn As eDLG_RET_VAL      'P[ro]/S: DialogReturn, ClearReturn
'
'NOTE: buttonSet is ignored when style is not custom. checkSet is set to nothing when checkboxes are not wanted

Public Property Get OutputFields() As String()
Attribute OutputFields.VB_Description = "Returns a String-type array filled with the field text values of the dialog."
    OutputFields = fieldOutput
End Property

Public Property Get OutputChecks() As CheckBoxConstants()
Attribute OutputChecks.VB_Description = "Returns a Long-type array filled with the check box values of the dialog."
    OutputChecks = checkOutput
End Property

Public Property Get EnableCancelKey() As Boolean
Attribute EnableCancelKey.VB_Description = "Enables or disables the Escape Key to exit the dialog. Should only be used with custom dialogs, otherwise unexpected output may be retrieved."
    EnableCancelKey = enableCancel
End Property

Public Property Let EnableCancelKey(ByVal enable As Boolean)
    enableCancel = enable
End Property

Public Property Get CheckItems(index As Integer) As tCheck_Format
Attribute CheckItems.VB_Description = "Retrieves or edits an individual item from the check box set."
    CheckItems = checkSet(index)
End Property

Public Property Let CheckItems(index As Integer, item As tCheck_Format)
    checkSet(index) = item
End Property

Public Property Get ButtonItems(index As Integer) As tButton_Format
Attribute ButtonItems.VB_Description = "Retrieves or edits an individual item from the button set."
    ButtonItems = buttonSet(index)
End Property

Public Property Let ButtonItems(index As Integer, item As tButton_Format)
    buttonSet(index) = item
End Property

Public Property Get FontDialog() As String
Attribute FontDialog.VB_Description = "Sets or gets the font of the dialog."
    FontDialog = dialFont
End Property

Public Property Let FontDialog(newFont As String)
    If FontExists(newFont) Then
        dialFont = newFont
    Else
        dialFont = ""
    End If
End Property

Public Property Get FieldWidth() As Long
Attribute FieldWidth.VB_Description = "Sets or gets the width of the fields in the dialog."
    FieldWidth = inputWidth
End Property

Public Property Let FieldWidth(ByVal width As Long)
    If width > 9 Then
        inputWidth = width
    Else
        inputWidth = 10
    End If
End Property

Public Property Get SoundMessage() As eOEM_SOUND
Attribute SoundMessage.VB_Description = "Sets or gets a constant that determines the sound played when the dialog is shown."
    SoundMessage = dialSound
End Property

Public Property Let SoundMessage(ByVal setSound As eOEM_SOUND)
    dialSound = setSound
End Property

Public Property Get FieldItems(index As Integer) As tField_Format
Attribute FieldItems.VB_Description = "Retrieves or edits an individual item from the field set."
    FieldItems = fieldSet(index)
End Property

Public Property Let FieldItems(index As Integer, item As tField_Format)
    fieldSet(index) = item
End Property

Public Property Get IconMessage() As eDLG_ICON
Attribute IconMessage.VB_Description = "Sets or gets a constant that determines the icon displayed in the dialog."
    IconMessage = dialIcon
End Property

Public Property Let IconMessage(ByVal setIcon As eDLG_ICON)
    dialIcon = setIcon
End Property

'Read-only
Public Property Get DialogReturn() As eDLG_RET_VAL
Attribute DialogReturn.VB_Description = "Gets the return value from the dialog."
    DialogReturn = lastReturn
End Property

Public Property Get InboxTitle() As String
Attribute InboxTitle.VB_Description = "Sets or gets the title of the dialog."
    InboxTitle = dlgTitle
End Property

Public Property Let InboxTitle(ByVal title As String)
    dlgTitle = title
End Property

Public Property Get PromptMessage() As String
Attribute PromptMessage.VB_Description = "Gets or sets the text that is displayed at the top of the dialog."
    PromptMessage = dlgPrompt
End Property

Public Property Let PromptMessage(ByVal textInput As String)
    dlgPrompt = textInput
End Property

Public Property Get SilentMode() As Boolean
Attribute SilentMode.VB_Description = "Gets or sets the sound mode of the dialog."
    SilentMode = silentPrompt
End Property

Public Property Let SilentMode(ByVal enabled As Boolean)
    silentPrompt = enabled
End Property

Public Property Get DialogStyle() As eDLG_BUTTON_STYLE
Attribute DialogStyle.VB_Description = "Gets or sets the dialog button style."
    DialogStyle = dlgStyle
End Property

Public Property Let DialogStyle(ByVal style As eDLG_BUTTON_STYLE)
    dlgStyle = style
End Property

'Constructor... or something similar
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Sets the default values of the properties of this class."
    'Default constructor - - - - - - -
    dlgTitle = VB.LoadResString(101)
    dlgPrompt = ""
    dialFont = "" 'dialog determines if it exists
    dlgStyle = DLG_OK_ONLY
    Set fieldSet = New Collection
    Set buttonSet = New Collection
    Set checkSet = New Collection
    dialIcon = DLG_NONE
    dialSound = OEM_BEEP
    inputWidth = DEF_FIELD_SIZE
    silentPrompt = True
    defaultButton = 0
    enableCancel = False
    lastReturn = -1 'means no dialog has been called yet
    '- - - - - - - - - - - - - - - - -
End Sub

Private Sub Class_Terminate()
    Set fieldSet = Nothing
    Set buttonSet = Nothing
    Set checkSet = Nothing
End Sub

Public Sub SetDefaultButton(ByVal index As Integer)
Attribute SetDefaultButton.VB_Description = "Sets the default button of the dialog. To disable pass 0 (zero) as index."
    defaultButton = index
End Sub

Public Sub AddFields(ByRef item As tField_Format)
Attribute AddFields.VB_Description = "Adds either text boxes or combination boxes to the dialog by using a parsing text. Use the respective Set property for correct parsing format."
    Call fieldSet.Add(item)
End Sub

Public Sub DeleteFields(ByVal index As Integer)
Attribute DeleteFields.VB_Description = "Deletes the specified field from the field set."
    Call fieldSet.Remove(index)
End Sub

Public Sub AddChecks(ByRef item As tCheck_Format)
Attribute AddChecks.VB_Description = "Adds a check box to the dialog. If you don't want check boxes just it leave empty or clear it."
    Call checkSet.Add(item)
End Sub

Public Sub DeleteChecks(ByRef index As Integer)
Attribute DeleteChecks.VB_Description = "Deletes the specified check box from the check box set."
    Call checkSet.Remove(index)
End Sub

Public Sub AddButtons(ByRef item As tButton_Format)
Attribute AddButtons.VB_Description = "Adds a button to the dialog. It is ignored unless style is set to custom."
    Call buttonSet.Add(item)
End Sub

Public Sub DeleteButtons(ByRef index As Integer)
Attribute DeleteButtons.VB_Description = "Deletes the specified button from the button set."
    Call buttonSet.Remove(index)
End Sub

Public Sub ClearFields()
Attribute ClearFields.VB_Description = "Empties the field set."
    Set fieldSet = Nothing
    Set fieldSet = New Collection
End Sub

Public Sub ClearChecks()
Attribute ClearChecks.VB_Description = "Empties the check box set."
    Set checkSet = Nothing
    Set checkSet = New Collection
End Sub

Public Sub ClearButtons()
Attribute ClearButtons.VB_Description = "Empties the button set."
    Set buttonSet = Nothing
    Set buttonSet = New Collection
End Sub

Public Sub ClearReturn()
Attribute ClearReturn.VB_Description = "Resets the return value of the class to the default value (-1), and also erases the fields and checks output."
    lastReturn = -1
    Erase fieldOutput
    Erase checkOutput
End Sub

'==============================================================================================
'==============================================================================================
'==============================================================================================
'General functions and methods (not related directly to properties)

Public Function Get_TextBoxFlags(ByVal fieldCaption As String, ByVal inputType As eFIELD_TYPE, _
                                 Optional ByVal fieldAlign As AlignmentConstants = AlignmentConstants.vbLeftJustify, _
                                 Optional ByVal validRuleset As String = "", _
                                 Optional ByVal maxLength As Long = 0, _
                                 Optional ByVal defaultText As String = "") As tField_Format
Attribute Get_TextBoxFlags.VB_Description = "Returns field object for a text box assigned with the arguments passed. Formatting flags are generated automatically."
    Dim interVal As tField_Format
    
    interVal.ffType = DLG_TEXTBOX
    interVal.ffFlags = fieldCaption & UPPER_DELIM & Get_StringTypeFromEnum(inputType) & UPPER_DELIM & _
                       Get_AlignString(fieldAlign) & UPPER_DELIM & validRuleset & UPPER_DELIM & _
                       Str(maxLength) & ";" & defaultText
    
    Get_TextBoxFlags = interVal
End Function

Public Function Get_ComboBoxFlags(ByVal fieldCaption As String, ByVal inputList As String, _
                                  ByVal listWithData As Boolean, Optional ByVal defaultItem As Variant = -1) As tField_Format
Attribute Get_ComboBoxFlags.VB_Description = "Returns field object for a combination box assigned with the arguments passed. Formatting flags are generated automatically."
    Dim retVal As tField_Format
    
    retVal.ffType = DLG_COMBOBOX
    retVal.ffFlags = fieldCaption & UPPER_DELIM & inputList & UPPER_DELIM & Trim(listWithData)
    
    If VarType(defaultItem) <> vbString Then
        retVal.ffFlags = retVal.ffFlags & UPPER_DELIM & "/index/" & Trim(Int(defaultItem))
    Else
        retVal.ffFlags = retVal.ffFlags & UPPER_DELIM & Trim(defaultItem)
    End If
    
    'NOTE: converting a boolean to string makes it language specific to the computer it is running on
    
    Get_ComboBoxFlags = retVal
End Function

Public Function Get_CheckBoxItem(ByVal checkCaption As String, _
                                 Optional ByVal checkDirection As eCheck_Alignment = CHKA_LEFT, _
                                 Optional ByVal defaultValue As CheckBoxConstants = vbUnchecked) As tCheck_Format
Attribute Get_CheckBoxItem.VB_Description = "Returns check box object assigned with the arguments passed."
    Dim retVal As tCheck_Format
    
    retVal.cfCaption = checkCaption
    retVal.cfCheck_Loc = checkDirection
    retVal.cfDefault_State = defaultValue
    
    Get_CheckBoxItem = retVal
End Function

Public Function Get_ButtonItem(ByVal buttonCaption As String, ByVal returnValue As eDLG_RET_VAL, _
                               Optional ByVal buttonWidth As Long = DEF_BUTTON_SIZE, _
                               Optional ByVal validateFields As Boolean = True) As tButton_Format
Attribute Get_ButtonItem.VB_Description = "Returns button object assigned with the arguments passed."
    Dim retVal As tButton_Format
    
    retVal.bfCaption = buttonCaption
    retVal.bfRetVal = returnValue
    retVal.bfWidth = buttonWidth
    retVal.bfValidate = validateFields
    
    Get_ButtonItem = retVal
End Function

Public Function Show() As eDLG_RET_VAL
Attribute Show.VB_Description = "Shows the dialog and retuns the value assigned to the button pressed."
    Dim text_Input As frmTextDialog
    Dim dlgParams As tInputBoxProp
    
    '- - - - - - - - - - - - - - - - -
    Set text_Input = New frmTextDialog
    
    With text_Input
        dlgParams.ibTitle = dlgTitle
        dlgParams.ibPrompt = dlgPrompt
        dlgParams.ibFont = dialFont
        dlgParams.ibButtonStyle = dlgStyle
        Set dlgParams.ibFieldSet = fieldSet
        Set dlgParams.ibButtonSet = buttonSet
        Set dlgParams.ibCheckSet = checkSet
        dlgParams.ibIcon = dialIcon
        dlgParams.ibSound = dialSound
        dlgParams.ibFieldWidth = inputWidth
        dlgParams.ibSilent = silentPrompt
        dlgParams.ibDefaultButton = defaultButton
        dlgParams.ibEnableESC = enableCancel
        
        .Initialize_Parameters dlgParams
        .Show 1
        
        If .data_Validation Then
            fieldOutput = .Get_FieldValues
            checkOutput = .Get_CheckValues
        End If
        lastReturn = .Dialog_Return
        Show = lastReturn
    End With
    
    Unload text_Input
    Set text_Input = Nothing
    '- - - - - - - - - - - - - - - - -
End Function
