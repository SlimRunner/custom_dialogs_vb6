VERSION 5.00
Begin VB.Form frmTextDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Placeholder"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   Icon            =   "frmTextDialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkChecks 
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMessageIcon 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   480
      Left            =   1200
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ComboBox cmbInputField 
      Height          =   315
      Index           =   0
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton btnDialButtons 
      BackColor       =   &H80000005&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtInputField 
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblInputTag 
      AutoSize        =   -1  'True
      Caption         =   "placeholder: "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Placeholder"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "frmTextDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'THE DISPOSAL OF THIS FORM SHOULD BE HANDLED BY THE CALLER

'Standard button names can be pulled from resouce file by adding 7000 plus the VB default msgbox button value
Private Const GAP_SPACING As Long = 8
Private Const ICON_SIZE  As Long = 32
Private Const CHECK_SIZE As Long = 12 'checkbox size is 13x13
Private Const DEF_BUTTON_SIZE As Long = 73

Private Const COLOR_INVALID_RED As Long = &H241CED
'vbButtonText used here

'Parameters
Private my_Params As tInputBoxProp
Public data_Validation As Boolean 'determines if the button pressed validated the data.

'Button return value
Public Dialog_Return As eDLG_RET_VAL

'Control variables
Private fieldBag As New Collection 'stores all fields in top-down order
Private tagBag As New Collection

'Flow control
Private loadFinished As Boolean

'Fetches properties from the caller function. Executes before LOAD event.
Public Sub Initialize_Parameters(dlgParams As Variant)
    my_Params = dlgParams
    
    If my_Params.ibButtonStyle = DLG_CUSTOM And my_Params.ibButtonSet.Count = 0 Then
        my_Params.ibButtonStyle = DLG_OK_ONLY
    End If
End Sub

'Collects and returns the text property of every field and returns it into a string array to the caller
Public Function Get_FieldValues() As String()
    Dim i As Long, iCount As Long
    
    iCount = fieldBag.Count
    If iCount > 0 Then
        ReDim retVal(0 To iCount - 1) As String
        
        For i = 1 To iCount 'Iterating 1 based collection
            If IsObject(fieldBag(i), "TextBox") Or fieldBag(i).Tag = Trim(False) Then
                retVal(i - 1) = fieldBag(i).Text
            Else
                If fieldBag(i).ListIndex >= 0 Then _
                    retVal(i - 1) = fieldBag(i).ItemData(fieldBag(i).ListIndex)
            End If
        Next
        
        Get_FieldValues = retVal
    End If
End Function

'Collects and returns the checked value of every checkbox and retruns it into a Long-type array to the caller
Public Function Get_CheckValues() As CheckBoxConstants()
    Dim i As Long, iCount As Long
    
    iCount = chkChecks.Count
    If iCount > 0 Then
        ReDim retVal(0 To iCount - 1) As CheckBoxConstants
        
        For i = 0 To iCount - 1 'Iterating 0 based array
            retVal(i) = chkChecks(i).Value
        Next
        
        Get_CheckValues = retVal
    End If
End Function

'Returns an increasing sequential number starting at 0
Private Function TabIndexer() As Long
    Static i As Long
    
    TabIndexer = i
    i = i + 1
End Function

'Returns the return value of the given button (as determined by index and style)
'Custom styles DO use this function
Private Function Get_ButtonReturn(ByVal stl_button As eDLG_BUTTON_STYLE, ByVal sub_index As Integer) As eDLG_RET_VAL
    Select Case stl_button
    Case DLG_OK_ONLY:               Get_ButtonReturn = RET_OK
    Case DLG_OK_CANCEL:             Get_ButtonReturn = Choose(sub_index + 1, RET_OK, RET_CANCEL)
    Case DLG_ABORT_RETRY_IGNORE:    Get_ButtonReturn = Choose(sub_index + 1, RET_ABORT, RET_RETRY, RET_IGNORE)
    Case DLG_YES_NO_CANCEL:         Get_ButtonReturn = Choose(sub_index + 1, RET_YES, RET_NO, RET_CANCEL)
    Case DLG_YES_NO:                Get_ButtonReturn = Choose(sub_index + 1, RET_YES, RET_NO)
    Case DLG_RETRY_CANCEL:          Get_ButtonReturn = Choose(sub_index + 1, RET_RETRY, RET_CANCEL)
    Case DLG_CUSTOM:                Get_ButtonReturn = my_Params.ibButtonSet(sub_index + 1).bfRetVal
    End Select
End Function

'Retruns the name of a button from a default GUI by index
Private Function Get_ButtonName(ByVal stl_button As eDLG_BUTTON_STYLE, ByVal sub_index As Integer) As String
    Select Case stl_button
    Case eDLG_BUTTON_STYLE.DLG_OK_ONLY
        Get_ButtonName = Choose(sub_index + 1, GetString_stdButtonName(DBTN_OK))
    Case eDLG_BUTTON_STYLE.DLG_OK_CANCEL
        Get_ButtonName = Choose(sub_index + 1, GetString_stdButtonName(DBTN_OK), _
                                               GetString_stdButtonName(DBTN_CANCEL))
    Case eDLG_BUTTON_STYLE.DLG_ABORT_RETRY_IGNORE
        Get_ButtonName = Choose(sub_index + 1, GetString_stdButtonName(DBTN_ABORT), _
                                               GetString_stdButtonName(DBTN_RETRY), _
                                               GetString_stdButtonName(DBTN_IGNORE))
    Case eDLG_BUTTON_STYLE.DLG_YES_NO_CANCEL
        Get_ButtonName = Choose(sub_index + 1, GetString_stdButtonName(DBTN_YES), _
                                               GetString_stdButtonName(DBTN_NO), _
                                               GetString_stdButtonName(DBTN_CANCEL))
    Case eDLG_BUTTON_STYLE.DLG_YES_NO
        Get_ButtonName = Choose(sub_index + 1, GetString_stdButtonName(DBTN_YES), _
                                               GetString_stdButtonName(DBTN_NO))
    Case eDLG_BUTTON_STYLE.DLG_RETRY_CANCEL
        Get_ButtonName = Choose(sub_index + 1, GetString_stdButtonName(DBTN_RETRY), _
                                               GetString_stdButtonName(DBTN_CANCEL))
    End Select
End Function

Private Function Get_ValidationAuth(ByVal stl_button As eDLG_BUTTON_STYLE, ByVal sub_index As Integer) As Boolean
    Select Case stl_button
    Case eDLG_BUTTON_STYLE.DLG_OK_ONLY
        Get_ValidationAuth = Choose(sub_index + 1, True)
        
    Case eDLG_BUTTON_STYLE.DLG_OK_CANCEL
        Get_ValidationAuth = Choose(sub_index + 1, True, False)
        
    Case eDLG_BUTTON_STYLE.DLG_ABORT_RETRY_IGNORE
        Get_ValidationAuth = Choose(sub_index + 1, False, True, False)
        
    Case eDLG_BUTTON_STYLE.DLG_YES_NO_CANCEL
        Get_ValidationAuth = Choose(sub_index + 1, True, False, False)
        
    Case eDLG_BUTTON_STYLE.DLG_YES_NO
        Get_ValidationAuth = Choose(sub_index + 1, True, False)
        
    Case eDLG_BUTTON_STYLE.DLG_RETRY_CANCEL
        Get_ValidationAuth = Choose(sub_index + 1, True, False)
        
    Case eDLG_BUTTON_STYLE.DLG_CUSTOM
        Get_ValidationAuth = my_Params.ibButtonSet(sub_index + 1).bfValidate
        
    End Select
End Function
'Retruns the number of buttons for default GUI styles
'Custom styles do NOT use this function
Private Function Determine_ButtonsQC(ByVal inProp As eDLG_BUTTON_STYLE) As Long
    Select Case inProp
    Case eDLG_BUTTON_STYLE.DLG_OK_ONLY
        Determine_ButtonsQC = 1
    Case eDLG_BUTTON_STYLE.DLG_OK_CANCEL
        Determine_ButtonsQC = 2
    Case eDLG_BUTTON_STYLE.DLG_ABORT_RETRY_IGNORE
        Determine_ButtonsQC = 3
    Case eDLG_BUTTON_STYLE.DLG_YES_NO_CANCEL
        Determine_ButtonsQC = 3
    Case eDLG_BUTTON_STYLE.DLG_YES_NO
        Determine_ButtonsQC = 2
    Case eDLG_BUTTON_STYLE.DLG_RETRY_CANCEL
        Determine_ButtonsQC = 2
    End Select
End Function

Private Function AreFieldsValid() As Boolean
    Dim strTag As String, i As Long
    Dim strRange As Variant, retVal As Boolean
    Dim i_valid As Boolean, strErrorMessage As String
    Dim i_message As String
    
    retVal = True
    
    'PASS AN ARRAY STRING CONTAINING ALL THE ERROR MESSAGES TO THE CALLER (Overall message) (Only one call, but requires a message handler at the caller)
    For i = 1 To fieldBag.Count
        If IsObject(fieldBag(i), "TextBox") Then
            'lbl Tag ->type
            'txt Tag ->charset/range
            
            strTag = fieldBag(i).Tag
            
            Select Case Get_EnumTypeFromString(lblInputTag(i - 1).Tag)
            Case eFIELD_TYPE.FT_TEXT
                i_valid = IsCharsetValid(fieldBag(i).Text, strTag)
                
                i_message = Get_InvalidFieldMessage(IFM_FIELD) & i & Get_InvalidFieldMessage(IFM_TEXT_SET) & "'" & strTag & "'."
                
            Case eFIELD_TYPE.FT_INTEGER
                strRange = Split(strTag, LOWER_DELIM)
                If UBound(strRange) = 1 Then
                    i_valid = IsRangeValid(fieldBag(i).Text, CLng(Val(strRange(0))), CLng(Val(strRange(1))))
                Else
                    i_valid = True   'default to true when RANGE is empty or not valid
                End If
                
                i_message = Get_InvalidFieldMessage(IFM_FIELD) & i & Get_InvalidFieldMessage(IFM_INTEGER_RANGE) & strRange(0) & Get_InvalidFieldMessage(IFM_INTRANGE_SUFFIX) & strRange(1)
                
            Case eFIELD_TYPE.FT_FLOAT
                strRange = Split(strTag, LOWER_DELIM)
                If UBound(strRange) = 1 Then
                    i_valid = IsRangeValid(fieldBag(i).Text, Val(strRange(0)), Val(strRange(1)))
                Else
                    i_valid = True   'default to true when empty or not valid
                End If
                
                i_message = Get_InvalidFieldMessage(IFM_FIELD) & i & Get_InvalidFieldMessage(IFM_FLOAT_RANGE) & strRange(0) & Get_InvalidFieldMessage(IFM_FLRANGE_SUFFIX) & strRange(1)
                
            End Select
            
        Else 'is Combo
            'TODO: true unless it's empty (-1)
            i_valid = fieldBag(i).ListIndex >= 0 Or fieldBag(i).ListCount = 0 '<- this last expression is a safe check
            i_message = Get_InvalidFieldMessage(IFM_FIELD) & i & Get_InvalidFieldMessage(IFM_LIST_EMPTY)
            
        End If
        
        If Not i_valid Then
            tagBag(i).ForeColor = COLOR_INVALID_RED
            'tagBag(i).ToolTipText = strErrorMessage 'for some reason it DOESN'T work
            strErrorMessage = strErrorMessage & i_message & vbCrLf
        End If
        
        retVal = retVal And i_valid
    Next i
    
    If Not retVal And data_Validation Then MsgBox Get_InvalidFieldMessage(IFM_PROMPTMESSAGE) & vbCrLf & vbCrLf & strErrorMessage, vbExclamation, Me.Caption
    AreFieldsValid = retVal
End Function

'Resets the collection of button styles and sets the default ones into it
Private Sub Get_DefaultButtons()
    Dim i As Long
    Dim defButton As tButton_Format
    
    'Clears the button set collection
    Set my_Params.ibButtonSet = Nothing
    Set my_Params.ibButtonSet = New Collection
    
    'Populates collection with the default items
    For i = 0 To Determine_ButtonsQC(my_Params.ibButtonStyle) - 1
        defButton.bfCaption = Get_ButtonName(my_Params.ibButtonStyle, i)
        defButton.bfWidth = DEF_BUTTON_SIZE
        defButton.bfValidate = Get_ValidationAuth(my_Params.ibButtonStyle, i)
        
        Call my_Params.ibButtonSet.Add(defButton)
    Next
End Sub

Private Sub Parse_FieldProperties(fieldItem As Object, labelItem As Object, ByVal propFlags As String)
    Dim strFlags As Variant
    Dim index As Integer
    
    strFlags = Split(propFlags, UPPER_DELIM)
    
    If IsObject(fieldItem, "TextBox") Then
        'Parsing format
        '<caption>;<type>;<alignment>;<charset/range>;<max_length>;<default_text>
        '    0        1        2            3              4             5
        
        labelItem.Caption = Get_DefaultLabelOnEmpty(strFlags(0), strFlags(1))
        labelItem.Tag = strFlags(1) 'used in txtInputField and AreFieldsValid
        fieldItem.Alignment = Get_AlignValue(strFlags(2))
        fieldItem.Tag = strFlags(3) 'used in txtInputField and AreFieldsValid
        fieldItem.maxLength = CLng(Val(strFlags(4)))
        fieldItem.Text = strFlags(5)
        
    Else 'IS A COMBO BOX
        'Parsing format
        '<caption>;<list>;<contains_data>;<default_item>
        '    0        1          2              3
        
        labelItem.Caption = Get_DefaultComboLabel(strFlags(0))
        fieldItem.Tag = strFlags(2) 'used in Get_FieldValues and AreFieldsValid
        Call PopulateList(fieldItem, strFlags(1), strFlags(2), LOWER_DELIM)
        If fieldItem.ListCount > 0 Then
            
            'Determines if flag passed is an index or a text item
            If InStr(1, strFlags(3), "/index/") Then
                index = CInt(Val(Replace(strFlags(3), "/index/", "")))
                If index < fieldItem.ListCount Then fieldItem.ListIndex = index
            Else
                On Error Resume Next 'To avoid error if the element does not exist in the list
                fieldItem.Text = CStr(strFlags(3))
                On Error GoTo 0
            End If
        End If
        
    End If
End Sub

'Loads all fields in proper order and stores them in a global collection
Private Sub Load_Fields()
    Dim eachField As Variant, i_fld As Long, i_lbl As Long
    Dim frstIter As Boolean, frst_Combo As Boolean, frst_Text As Boolean
    
    'frst_Combo and frst_Text are used because the visible property does not refresh _
    during the duration of this Sub procedure. Also when they are true means that the _
    first element in their respective array is in use already.
    
    'Two counters are used below because the respective index of the label and field _
    are bound to not be parallel since the field can be either of two controls.
    
    For Each eachField In my_Params.ibFieldSet
        
        'Prevents attempting to re-load the first element
        If frstIter Then
            i_lbl = lblInputTag.Count 'determines the current label
            Load lblInputTag(i_lbl)
            
            'Determines if the next element to load is a TextBox or a ComboBox
            If eachField.ffType = eDLG_FIELD_STYLE.DLG_TEXTBOX And frst_Text Then
                i_fld = txtInputField.Count 'determines the current field
                Load txtInputField(i_fld)
                
            ElseIf eachField.ffType = eDLG_FIELD_STYLE.DLG_COMBOBOX And frst_Combo Then
                i_fld = cmbInputField.Count
                Load cmbInputField(i_fld)
                
            Else
                i_fld = 0
                
            End If
            
        Else
            i_lbl = 0
            i_fld = 0
            
            frstIter = True
        End If
        
        'Setting up properties
        lblInputTag(i_lbl).Visible = True
        lblInputTag(i_lbl).TabIndex = TabIndexer()
        
        If eachField.ffType = eDLG_FIELD_STYLE.DLG_TEXTBOX Then
            frst_Text = True
            txtInputField(i_fld).Visible = True
            txtInputField(i_fld).TabIndex = TabIndexer()
            
            Call Parse_FieldProperties(txtInputField(i_fld), lblInputTag(i_lbl), eachField.ffFlags)
            Call fieldBag.Add(txtInputField(i_fld), Str(txtInputField(i_fld).hWnd))
            Call tagBag.Add(lblInputTag(i_lbl), Str(txtInputField(i_fld).hWnd))
            
        ElseIf eachField.ffType = eDLG_FIELD_STYLE.DLG_COMBOBOX Then
            frst_Combo = True
            cmbInputField(i_fld).Visible = True
            cmbInputField(i_fld).TabIndex = TabIndexer()
            
            Call Parse_FieldProperties(cmbInputField(i_fld), lblInputTag(i_lbl), eachField.ffFlags)
            Call fieldBag.Add(cmbInputField(i_fld), Str(cmbInputField(i_fld).hWnd))
            Call tagBag.Add(lblInputTag(i_lbl), Str(cmbInputField(i_fld).hWnd))
            
        End If
    Next eachField
End Sub

Private Sub Populate_Dialog()
    Dim frstIter As Boolean
    Dim i_each As Variant, i As Long
    
    'Loads all TextBoxes and/or ComboBoxes in their respective array, but _
    stores them sequentially into a collection for ease of management.
    Load_Fields
    
    'Loads all the Checkboxes in an array
    For Each i_each In my_Params.ibCheckSet
        
        'Prevents attempting to re-load the first element
        If frstIter Then
            i = chkChecks.Count
            Load chkChecks(i)
        Else
            i = 0
            frstIter = True
        End If
        
        chkChecks(i).Visible = True
        chkChecks(i).TabIndex = TabIndexer()
        chkChecks(i).Caption = i_each.cfCaption
        chkChecks(i).Alignment = i_each.cfCheck_Loc
        chkChecks(i).Value = i_each.cfDefault_State
    Next i_each
    frstIter = False
    
    'Loads all the Buttons in an array
    For Each i_each In my_Params.ibButtonSet
        
        'Prevents attempting to re-load the first element
        If frstIter Then
            i = btnDialButtons.Count
            Load btnDialButtons(i)
        Else
            i = 0
            frstIter = True
        End If
        
        btnDialButtons(i).Visible = True
        If i + 1 = my_Params.ibDefaultButton Then _
            btnDialButtons(i).Default = True
        btnDialButtons(i).TabIndex = TabIndexer()
        btnDialButtons(i).Caption = i_each.bfCaption
        btnDialButtons(i).Tag = Str(i_each.bfRetVal) 'This may be unnecessary/unused
        btnDialButtons(i).width = i_each.bfWidth
    Next i_each
    frstIter = False
    
    'This needs to be separate because when the first copy of the original _
    button is loaded it inherits the Default property, and since it's also _
    true, it overrides the one for the original button.
    If my_Params.ibDefaultButton = 1 Then btnDialButtons(0).Default = True
End Sub

'Adjusts every textbox height to match its font size
Private Sub Adjust_TextBoxes_Height()
    Dim ascii_charset As String * 256
    Dim char_heigth As Long, gap_height As Long
    Dim i As Long
    
    For i = 0 To 255
        ascii_charset = ascii_charset & Chr(i)
    Next i
    
    char_heigth = Me.TextHeight(ascii_charset)
    gap_height = Int(char_heigth * 0.2 + 0.5)
    
    For i = 0 To txtInputField.Count - 1
        txtInputField(i).Height = char_heigth + gap_height
    Next i
End Sub

'Adjusts every checkbox size (w/h)
Private Sub Adjust_CheckBoxes_Size()
    Dim i As Long, noAmpersand As String
    
    For i = 0 To chkChecks.UBound
        noAmpersand = Replace(chkChecks(i).Caption, "&", "")
        chkChecks(i).width = Me.TextWidth(noAmpersand) + GAP_SPACING + CHECK_SIZE
        chkChecks(i).Height = Me.TextHeight(noAmpersand) + GAP_SPACING
    Next i
End Sub

'Adjusts this form's width and height to fit its dynamic contents correctly (assumes controls are sorted)
Private Sub Resize_ThisForm()
    Dim frame_height As Long
    Dim frame_width As Long
    Dim rightEdge As Long
    Dim lowestEdge As Long
    
    frame_width = Me.width - VB.Screen.TwipsPerPixelX * Me.ScaleWidth
    frame_height = Me.Height - VB.Screen.TwipsPerPixelY * Me.ScaleHeight
    
    rightEdge = (Get_RightMostControl(Me) + GAP_SPACING) * VB.Screen.TwipsPerPixelX
    lowestEdge = (Get_LowestControl(Me) + GAP_SPACING) * VB.Screen.TwipsPerPixelY
    
    Me.width = frame_width + rightEdge
    Me.Height = frame_height + lowestEdge
End Sub

'Rearranges all controls in an organized fashion (assumes correct width has been determined)
Private Sub Arrange_GUILayout()
    Dim i As Long, x As Long, y As Long
    Dim rightMost As Long 'stores the RIGHT edge of the farthest spanning label
    Dim curY As Long 'stores the lowest member of the previous step (reference)
    
    'At this point all controls should be loaded already and their text propeties set
    '- - - - - - - - - - - - - - - - - - - - - - - - - - -
    'Moving icon
    If my_Params.ibIcon = DLG_NONE Then
        picMessageIcon.Move -ICON_SIZE, 0, ICON_SIZE, ICON_SIZE
    Else
        picMessageIcon.Move GAP_SPACING * 1.5, GAP_SPACING * 1.5, ICON_SIZE, ICON_SIZE
    End If
    '- - - - - - - - - - - - - - - - - - - - - - - - - - -
    'Moving prompt
    'TODO: determine if wrapping is needed
    x = Get_RightLoc(picMessageIcon) + GAP_SPACING
    y = picMessageIcon.Top + (picMessageIcon.Height - Me.TextHeight("lL01'{}")) / 2
    lblDescription.Move x, y
    
    If lblDescription.Caption = "" And my_Params.ibIcon = DLG_NONE Then
        lblDescription.Move -lblDescription.width, -lblDescription.Height
        picMessageIcon.Move -picMessageIcon.width, -picMessageIcon.Height
    End If
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - -
    'Resizing input memebers
    Adjust_TextBoxes_Height
    Adjust_CheckBoxes_Size
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - -
    'Moving labels
    If fieldBag.Count > 0 Then
        x = GAP_SPACING
        y = Get_Greater(Get_BottomLoc(picMessageIcon), Get_BottomLoc(lblDescription)) + GAP_SPACING * 1.5
        For i = 0 To lblInputTag.UBound
            lblInputTag(i).Move x, y
            rightMost = Get_Greater(rightMost, Get_RightLoc(lblInputTag(i)))
            y = y + lblInputTag(i).Height + GAP_SPACING * 2
        Next
    End If
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - -
    'Moving fields
    If fieldBag.Count > 0 Then
        x = rightMost + GAP_SPACING
        For i = 1 To fieldBag.Count
            y = lblInputTag(i - 1).Top + (lblInputTag(i - 1).Height - fieldBag(i).Height) / 2
            fieldBag(i).Move x, y, my_Params.ibFieldWidth
            'THIS MAY NEED A CHANGE
            'left should be re-determined after the form is resized
        Next i
        
        curY = Get_BottomLoc(fieldBag(i - 1))
    Else
        curY = Get_LowestControl(Me)
    End If
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - -
    'Moving check boxes
    If my_Params.ibCheckSet.Count > 0 Then
        x = rightMost + GAP_SPACING
        y = curY + GAP_SPACING * 1
        For i = 0 To chkChecks.UBound
            chkChecks(i).Move x, y
            y = y + chkChecks(i).Height
        Next i
        
        curY = Get_BottomLoc(chkChecks(i - 1))
    Else
        curY = Get_LowestControl(Me)
    End If
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - -
    'Moving buttons
    x = GAP_SPACING
    y = curY + GAP_SPACING * 2.5
    For i = 0 To btnDialButtons.UBound
        btnDialButtons(i).Move x, y
        x = x + btnDialButtons(i).width + GAP_SPACING
    Next i
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - -
    Resize_ThisForm
    
    Dim leftJust As Long 'left justify
    
    'aligning buttons to the right of the form
    leftJust = Get_RightLoc(btnDialButtons(btnDialButtons.UBound)) + GAP_SPACING - Me.ScaleWidth
    If leftJust < 0 Then
    For i = 0 To btnDialButtons.UBound
        btnDialButtons(i).Left = btnDialButtons(i).Left - leftJust
    Next i
    End If
End Sub

'Hides this form (read top comment) so that the helper function from caller can read the return variables
Private Sub Form_SafeExit(ByVal dial_ret As eDLG_RET_VAL)
    Dialog_Return = dial_ret
    Me.Hide
End Sub

'Assigns the font of the form to each control in it
Private Sub Set_FontToAll()
    Dim frmControls As Object
    
    On Error Resume Next
    For Each frmControls In Me.Controls
        frmControls.Font = Me.Font
    Next frmControls
    On Error GoTo 0
End Sub

'Moves all controls outside of the form (top-left)
Private Sub MoveToSafe_Load()
    Dim frmControls As Object
    
    For Each frmControls In Me.Controls
        frmControls.Move -frmControls.width, -frmControls.Height
    Next frmControls
End Sub

'=====================================================================================================================================
'=====================================================================================================================================
'=====================================================================================================================================
'=====================================================================================================================================
'=====================================================================================================================================

Private Sub txtInputField_Change(index As Integer)
    If loadFinished Then _
        tagBag.item(Str(txtInputField(index).hWnd)).ForeColor = vbButtonText
End Sub

Private Sub txtInputField_KeyPress(index As Integer, KeyAscii As Integer)
    Dim SlctChrt As String
    
    SlctChrt = GetTypeCharset(lblInputTag(index).Tag)
    
    If SlctChrt = "" Then
        SlctChrt = txtInputField(index).Tag
    End If
    
    If SlctChrt <> "" And Not (InStr(1, SlctChrt, Chr(KeyAscii)) <> 0 Or KeyAscii = 8) Then KeyAscii = 0
    'If Not (InStr(1, SlctChrt, Chr(KeyAscii)) <> 0 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub cmbInputField_Click(index As Integer)
    If loadFinished Then _
        tagBag.item(Str(cmbInputField(index).hWnd)).ForeColor = vbButtonText
End Sub

Private Sub btnDialButtons_Click(index As Integer)
    data_Validation = Get_ValidationAuth(my_Params.ibButtonStyle, index)
    
    'Should be false only when "data needs to be validated" and "data is NOT valid" any other case should be true
    
    'If data needs to be validated and therefore data is valid Then
    If data_Validation Imp AreFieldsValid Then
        Form_SafeExit Get_ButtonReturn(my_Params.ibButtonStyle, index)
    Else
        'AreFieldsValid should take care of error messages
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Form_QueryUnload 0, QueryUnloadConstants.vbFormControlMenu
    End If
End Sub

Private Sub Form_Load()
    
    'Safe-check for font assignment
    If FontExists(my_Params.ibFont) Then
        Me.Font = my_Params.ibFont
    Else
        my_Params.ibFont = Me.Font
    End If
    
    Set_FontToAll 'Make all controls get the same font as the form
    
    If Not my_Params.ibSilent Then
        'Select type of sound
        Call DialogBeep(my_Params.ibSound)
    End If
    
    picMessageIcon.TabIndex = TabIndexer()
    If my_Params.ibIcon <> DLG_NONE Then
        picMessageIcon.Visible = True
        'Draws icon only when required
        Call DrawOEM_Icon(picMessageIcon.hdc, my_Params.ibIcon)
    End If
    
    'Captioning and title
    Me.Caption = my_Params.ibTitle
    lblDescription.Caption = my_Params.ibPrompt
    lblDescription.TabIndex = TabIndexer()
    
    'Generates default buttons when appropriate
    If my_Params.ibButtonStyle <> DLG_CUSTOM Then
        Get_DefaultButtons
    End If
    
    MoveToSafe_Load 'moves all controls
    Populate_Dialog 'loads controls
    Arrange_GUILayout 'organizes GUI
    
    'draws a permanent white stroke at the bottom of dialog for style
    Me.AutoRedraw = True
    Draw_Stroke Me, 0, btnDialButtons(0).Top - GAP_SPACING, Me.ScaleWidth, Me.ScaleHeight, vbWindowBackground, vbButtonShadow
    Me.AutoRedraw = False
    
    loadFinished = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'makes sure that the form is unloaded properly (or hidden)
    If UnloadMode = QueryUnloadConstants.vbFormControlMenu Then
        Cancel = 1
        If (my_Params.ibButtonStyle < 6 And _
        (my_Params.ibButtonStyle Mod 2) = 1 Or _
        my_Params.ibButtonStyle = 0) Or _
        my_Params.ibEnableESC Then _
            Form_SafeExit IIf(my_Params.ibButtonStyle <> DLG_OK_ONLY, RET_CANCEL, RET_OK)
    End If
End Sub
