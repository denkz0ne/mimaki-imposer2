VERSION 5.00
Begin VB.UserForm frmMimakiImposerPanel
   Caption         =   "Mimaki Imposer v2.4.0 / 2026-10-02"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5700
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMimakiImposerPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cboLayout As MSForms.ComboBox
Private WithEvents cmdRefresh As MSForms.CommandButton
Private WithEvents optCopies As MSForms.ToggleButton
Private WithEvents optAllPages As MSForms.ToggleButton
Private WithEvents optCommonBackLast As MSForms.ToggleButton
Private WithEvents optAlternatingPairs As MSForms.ToggleButton
Private WithEvents txtCopies As MSForms.TextBox
Private WithEvents txtPageRange As MSForms.TextBox
Private WithEvents tglLandscape As MSForms.ToggleButton
Private WithEvents tglPortrait As MSForms.ToggleButton
Private WithEvents tglFront As MSForms.ToggleButton
Private WithEvents tglBack As MSForms.ToggleButton
Private WithEvents txtDocNumber As MSForms.TextBox
Private WithEvents txtStartSlot As MSForms.TextBox
Private WithEvents txtOutputFolder As MSForms.TextBox
Private WithEvents cmdOutputFolder As MSForms.CommandButton
Private WithEvents chkOpenExisting As MSForms.CheckBox
Private WithEvents chkPreserveVector As MSForms.CheckBox
Private WithEvents cmdRun As MSForms.CommandButton
Private WithEvents cmdClose As MSForms.CommandButton

Private lblLayoutInfo As MSForms.Label
Private lblStatus As MSForms.Label
Private mUpdating As Boolean
Private mSourceDoc As Document
Private mSourceSelectionShapes As Collection
Private mCachedLayoutName As String
Private mCachedCapacity As Long
Private mCachedSlotInfoLayoutName As String
Private mCachedSlotWidth As Double
Private mCachedSlotHeight As Double
Private mCachedGraphicWidth As Double
Private mCachedGraphicHeight As Double

Private Const COLOR_PANEL As Long = &HF3F1EE
Private Const COLOR_FRAME As Long = &HFBFAF8
Private Const COLOR_INFO As Long = &HEDE5D8
Private Const COLOR_PRIMARY As Long = &H2B9D59
Private Const COLOR_PRIMARY_SOFT As Long = &HE7F5EB
Private Const COLOR_SECONDARY_SOFT As Long = &HEDF3FB
Private Const COLOR_WARNING_SOFT As Long = &HE6F2FB
Private Const COLOR_MUTED As Long = &HEEE9E2
Private Const COLOR_TEXT_DARK As Long = &H2F2F2F
Private Const COLOR_TEXT_LIGHT As Long = &HFFFFFF
Private Const SETTINGS_APP As String = "MimakiImposer"
Private Const SETTINGS_SECTION As String = "Panel"

Private Sub UserForm_Initialize()
    mUpdating = True
    Set mSourceDoc = PickBestSourceDocument()
    CaptureCurrentSourceSelection
    BuildUi
    LoadLayouts
    LoadLastSettings
    ApplySourceDocumentAutoDefaults
    mUpdating = False
    ApplyFullSheetCopiesDefault
    UpdateLayoutInfo
    SetStatus "Pripravene. Rozlozit vytvori vystup cez stabilny temp workflow a panel sa po uspechu zavrie."
End Sub

Private Sub BuildUi()
    Dim fraTemplate As MSForms.Frame
    Dim fraInput As MSForms.Frame
    Dim fraOutput As MSForms.Frame
    Dim lbl As MSForms.Label

    Me.Caption = MimakiV21_GetScriptVersionLabel()
    Me.Width = 408
    Me.Height = 466
    Me.BackColor = COLOR_PANEL

    Set fraTemplate = AddFrame("fraTemplate", "Sablona", 8, 8, 382, 82)
    Set lbl = AddLabel(fraTemplate, "lblLayout", "Produkt", 8, 18, 48, 14)

    Set cboLayout = fraTemplate.Controls.Add("Forms.ComboBox.1", "cboLayout", True)
    cboLayout.Left = 60
    cboLayout.Top = 15
    cboLayout.Width = 244
    cboLayout.Style = fmStyleDropDownList
    cboLayout.ControlTipText = "Vyber produkt podla predlohovej vrstvy, ktorej nazov konci na SLOT."

    Set cmdRefresh = fraTemplate.Controls.Add("Forms.CommandButton.1", "cmdRefresh", True)
    cmdRefresh.Left = 310
    cmdRefresh.Top = 12
    cmdRefresh.Width = 60
    cmdRefresh.Height = 24

    Set lblLayoutInfo = AddLabel(fraTemplate, "lblLayoutInfo", "Nacitavam produkt...", 8, 42, 362, 38)
    lblLayoutInfo.BackColor = COLOR_INFO
    lblLayoutInfo.BorderStyle = fmBorderStyleSingle
    lblLayoutInfo.WordWrap = True

    Set fraInput = AddFrame("fraInput", "Vstupna grafika", 8, 96, 382, 150)

    Set optCopies = fraInput.Controls.Add("Forms.ToggleButton.1", "optCopies", True)
    optCopies.Left = 10
    optCopies.Top = 18
    optCopies.Width = 82
    optCopies.Height = 44

    Set optAllPages = fraInput.Controls.Add("Forms.ToggleButton.1", "optAllPages", True)
    optAllPages.Left = 102
    optAllPages.Top = 18
    optAllPages.Width = 82
    optAllPages.Height = 44

    Set optCommonBackLast = fraInput.Controls.Add("Forms.ToggleButton.1", "optCommonBackLast", True)
    optCommonBackLast.Left = 194
    optCommonBackLast.Top = 18
    optCommonBackLast.Width = 82
    optCommonBackLast.Height = 44

    Set optAlternatingPairs = fraInput.Controls.Add("Forms.ToggleButton.1", "optAlternatingPairs", True)
    optAlternatingPairs.Left = 286
    optAlternatingPairs.Top = 18
    optAlternatingPairs.Width = 82
    optAlternatingPairs.Height = 44
    Set lbl = AddLabel(fraInput, "lblCopies", "Pocet kusov", 10, 70, 78, 14)
    Set txtCopies = fraInput.Controls.Add("Forms.TextBox.1", "txtCopies", True)
    txtCopies.Left = 10
    txtCopies.Top = 86
    txtCopies.Width = 82
    txtCopies.ControlTipText = "Pocet kusov pri rezime Kopie strany."

    Set lbl = AddLabel(fraInput, "lblPageRange", "Rozsah stran", 104, 70, 86, 14)
    Set txtPageRange = fraInput.Controls.Add("Forms.TextBox.1", "txtPageRange", True)
    txtPageRange.Left = 104
    txtPageRange.Top = 86
    txtPageRange.Width = 82
    txtPageRange.ControlTipText = "Rozsah stran pri P/Z workflow, napriklad 1-158 alebo 3-20."

    Set lbl = AddLabel(fraInput, "lblOrientation", "Orientacia vstupu", 10, 118, 96, 14)

    Set tglLandscape = fraInput.Controls.Add("Forms.ToggleButton.1", "tglLandscape", True)
    tglLandscape.Left = 112
    tglLandscape.Top = 108
    tglLandscape.Width = 118
    tglLandscape.Height = 34

    Set tglPortrait = fraInput.Controls.Add("Forms.ToggleButton.1", "tglPortrait", True)
    tglPortrait.Left = 238
    tglPortrait.Top = 108
    tglPortrait.Width = 118
    tglPortrait.Height = 34

    Set lbl = AddLabel(fraInput, "lblOrientationHelp", "Vyber orientaciu podla zdrojovej grafiky.", 10, 132, 346, 14)
    lbl.ForeColor = &H6F6F6F

    Set fraOutput = AddFrame("fraOutput", "Vystup", 8, 252, 382, 100)

    Set tglFront = fraOutput.Controls.Add("Forms.ToggleButton.1", "tglFront", True)
    tglFront.Caption = "P"
    tglFront.Left = 8
    tglFront.Top = 18
    tglFront.Width = 40
    tglFront.Height = 26
    tglFront.ControlTipText = "Predna strana."

    Set tglBack = fraOutput.Controls.Add("Forms.ToggleButton.1", "tglBack", True)
    tglBack.Caption = "Z"
    tglBack.Left = 52
    tglBack.Top = 18
    tglBack.Width = 40
    tglBack.Height = 26
    tglBack.ControlTipText = "Zadna strana."

    Set lbl = AddLabel(fraOutput, "lblDocNumber", "VP", 108, 20, 20, 14)
    Set txtDocNumber = fraOutput.Controls.Add("Forms.TextBox.1", "txtDocNumber", True)
    txtDocNumber.Left = 132
    txtDocNumber.Top = 17
    txtDocNumber.Width = 88
    txtDocNumber.ControlTipText = "Cislo vyrobneho prikazu."

    Set lbl = AddLabel(fraOutput, "lblStartSlot", "Start", 228, 20, 28, 14)
    Set txtStartSlot = fraOutput.Controls.Add("Forms.TextBox.1", "txtStartSlot", True)
    txtStartSlot.Left = 260
    txtStartSlot.Top = 17
    txtStartSlot.Width = 42
    txtStartSlot.ControlTipText = "Pozicia, od ktorej sa ma zacat rozklad."

    Set chkOpenExisting = fraOutput.Controls.Add("Forms.CheckBox.1", "chkOpenExisting", True)
    chkOpenExisting.Caption = "Pouzit existujuci subor"
    chkOpenExisting.Left = 8
    chkOpenExisting.Top = 50
    chkOpenExisting.Width = 144
    chkOpenExisting.ControlTipText = "Ak vystup uz existuje, pokracovat v nom namiesto vytvorenia novej kopie."
    Set chkPreserveVector = fraOutput.Controls.Add("Forms.CheckBox.1", "chkPreserveVector", True)
    chkPreserveVector.Caption = "Zachovat vektor"
    chkPreserveVector.Left = 160
    chkPreserveVector.Top = 50
    chkPreserveVector.Width = 158
    chkPreserveVector.ControlTipText = "Nevytvara TIFF; zachova zdrojove vrstvy a vektorovu geometriu. Bitmapove objekty zostavaju bitmapami."

    Set lbl = AddLabel(fraOutput, "lblOutputFolder", "Folder", 8, 76, 40, 14)
    Set txtOutputFolder = fraOutput.Controls.Add("Forms.TextBox.1", "txtOutputFolder", True)
    txtOutputFolder.Left = 54
    txtOutputFolder.Top = 72
    txtOutputFolder.Width = 276

    Set cmdOutputFolder = fraOutput.Controls.Add("Forms.CommandButton.1", "cmdOutputFolder", True)
    cmdOutputFolder.Left = 336
    cmdOutputFolder.Top = 70
    cmdOutputFolder.Width = 30
    cmdOutputFolder.Height = 24

    Set lblStatus = AddLabel(Me, "lblStatus", "", 8, 358, 382, 34)
    lblStatus.BackColor = &HFFFFFF
    lblStatus.BorderStyle = fmBorderStyleSingle
    lblStatus.WordWrap = True

    Set cmdRun = Me.Controls.Add("Forms.CommandButton.1", "cmdRun", True)
    cmdRun.Left = 8
    cmdRun.Top = 402
    cmdRun.Width = 258
    cmdRun.Height = 32
    cmdRun.BackColor = COLOR_PRIMARY
    cmdRun.ForeColor = COLOR_TEXT_LIGHT

    Set cmdClose = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    cmdClose.Left = 274
    cmdClose.Top = 402
    cmdClose.Width = 116
    cmdClose.Height = 32

    ApplyButtonIcons
End Sub

Private Function AddFrame(ByVal controlName As String, ByVal captionText As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.Frame
    Set AddFrame = Me.Controls.Add("Forms.Frame.1", controlName, True)
    AddFrame.Caption = captionText
    AddFrame.Left = x
    AddFrame.Top = y
    AddFrame.Width = w
    AddFrame.Height = h
    AddFrame.BackColor = COLOR_FRAME
End Function

Private Function AddLabel(ByVal parentControl As Object, ByVal controlName As String, ByVal captionText As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.Label
    Set AddLabel = parentControl.Controls.Add("Forms.Label.1", controlName, True)
    AddLabel.Caption = captionText
    AddLabel.Left = x
    AddLabel.Top = y
    AddLabel.Width = w
    AddLabel.Height = h
    AddLabel.BackStyle = fmBackStyleTransparent
End Function

Private Sub ApplyButtonIcons()
    cmdRefresh.Caption = "Refresh"
    cmdRefresh.ControlTipText = "Obnovit zoznam produktov zo sablony."
    SetControlIcon cmdOutputFolder, "folder.bmp", "...", "Vybrat vystupny folder.", True, fmPicturePositionCenter
    SetControlIcon cmdRun, "run.bmp", "Be" & ChrW$(382), "Spustit rozkladanie.", False, fmPicturePositionLeftCenter
    SetControlIcon cmdClose, "close.bmp", "Zavriet", "Zavriet panel bez rozkladania.", False, fmPicturePositionLeftCenter
    SetControlIcon optCopies, "copies.bmp", "KOPIE", "Rozlozit viac kusov z aktualne oznacenej grafiky.", False, fmPicturePositionLeftCenter
    SetControlIcon optAllPages, "pages.bmp", "P ZZZ", "Prva strana je spolocny predok, dalsie strany su zadky.", False, fmPicturePositionLeftCenter
    SetControlIcon optCommonBackLast, "pages.bmp", "PPP Z", "Posledna strana je spolocny zadok, predchadzajuce strany su predky.", False, fmPicturePositionLeftCenter
    SetControlIcon optAlternatingPairs, "pages.bmp", "PZ PZ", "Strany su pary predok/zadok: 1=P, 2=Z, 3=P, 4=Z.", False, fmPicturePositionLeftCenter
    SetControlIcon tglLandscape, "landscape.bmp", "Na sirku", "Zdrojova grafika je orientovana na sirku.", False, fmPicturePositionLeftCenter
    SetControlIcon tglPortrait, "portrait.bmp", "Na vysku", "Zdrojova grafika je orientovana na vysku.", False, fmPicturePositionLeftCenter
End Sub

Private Sub SetControlIcon(ByVal ctl As Object, ByVal iconFileName As String, ByVal fallbackCaption As String, ByVal tooltipText As String, ByVal iconOnly As Boolean, ByVal picturePosition As fmPicturePosition)
    Dim iconPath As String

    ctl.ControlTipText = tooltipText
    iconPath = EnsureEmbeddedIconFile(iconFileName)
    If TryApplyControlIcon(ctl, iconPath, fallbackCaption, iconOnly, picturePosition) Then Exit Sub

    ctl.Caption = fallbackCaption
End Sub

Private Function TryApplyControlIcon(ByVal ctl As Object, ByVal iconPath As String, ByVal fallbackCaption As String, ByVal iconOnly As Boolean, ByVal picturePosition As fmPicturePosition) As Boolean
    On Error GoTo EH

    If Len(Trim$(iconPath)) = 0 Then Exit Function
    If Len(Dir$(iconPath)) = 0 Then Exit Function

    ctl.Picture = LoadPicture(iconPath)
    ctl.PicturePosition = picturePosition
    If iconOnly Then
        ctl.Caption = ""
    Else
        ctl.Caption = fallbackCaption
    End If

    TryApplyControlIcon = True
    Exit Function

EH:
    TryApplyControlIcon = False
End Function

Private Function EnsureEmbeddedIconFile(ByVal iconFileName As String) As String
    Dim iconFolder As String
    Dim iconPath As String
    Dim encoded As String

    encoded = EmbeddedIconBase64(iconFileName)
    If Len(encoded) = 0 Then Exit Function

    iconFolder = Environ$("TEMP")
    If Len(iconFolder) = 0 Then iconFolder = Environ$("TMP")
    If Len(iconFolder) = 0 Then Exit Function
    iconFolder = CombineLocalPath(iconFolder, "MimakiImposer")
    EnsureLocalFolderExists iconFolder
    iconFolder = CombineLocalPath(iconFolder, "embedded_icons")
    EnsureLocalFolderExists iconFolder

    iconPath = CombineLocalPath(iconFolder, LCase$(iconFileName))
    If Len(Dir$(iconPath)) = 0 Then WriteBase64BinaryFile iconPath, encoded
    If Len(Dir$(iconPath)) > 0 Then EnsureEmbeddedIconFile = iconPath
End Function

Private Function EmbeddedIconBase64(ByVal iconFileName As String) As String
    Dim s As String

    Select Case LCase$(Trim$(iconFileName))
        Case "refresh.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////8plw/fHr/////////////////////////////////Ozl+tjJ9bGS9bGS+tjJ/Ozl////////////"
            s = s & "////////////////////////////////////////////////////74pd6200+tjJ/////////////////vv59rqg74BO62Ml"
            s = s & "62Ml62Ml62Ml62Ml62Ml8Y9j98Wu/vv5////////////////////////////////////////////////////8plw62Ml62Ml"
            s = s & "9bGS////////+tjJ7ndA62Ml62Ml62Ml62Ml62Ml62Ml62Ml62Ml62Ml62Ml620098Wu////////////////////////////"
            s = s & "////////////////////9KiG62Ml62Ml7ndA/fHr+dTB7Gkt62Ml62Ml7XxH9a6M+tjJ/vv5////+tjJ9KiG7XI662Ml62Ml"
            s = s & "62Ml9sCn////////////////////////////////////////////9bWZ62Ml6200/Ofe++LX7Gkt62Ml7Gkt9sCn////////"
            s = s & "////////////////////////+Mm1620062Ml62Ml9bGS////////////////////////////////////////98Wu7Gkt+tjJ"
            s = s & "++LX620062Ml7Gkt+dTB////////////////////////////////////////++LX7XI662Ml74BO/vv5////////////////"
            s = s & "////////////////////+tjJ+Mm1////8pVq62Ml62Ml98Wu////////////////////////////////////////////////"
            s = s & "/Ozl74pd/Ozl/////////////////////////////////////////////////fHr7Gkt62Ml9KN/////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////9rqg"
            s = s & "62Ml62Ml+t7Q////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////74BO62Ml7XxH////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////62Ml62Ml9bGS////////////////////////"
            s = s & "/////////////////////////////////////////////////////////////////////////////////////////Ozl62Ml"
            s = s & "62Ml98Wu////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "/////////////////////////Ozl62Ml62Ml9sCn////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////7Gkt62Ml9KiG////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////74VW"
            s = s & "62Ml7XxH////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////9bGS62Ml62Ml++LX////////////////////////////////////////////////////"
            s = s & "/////////////////////////////////////////////////////////////fHr7Gkt62Ml74RW////////////////////"
            s = s & "////////////////////////////////////9KN/9sCn////////////////////////////////////////////////////"
            s = s & "9KiG62Il6mMl739M////////////////////////////////////////////////9bGS62Ml62Ml9sCn////+dTB+Mm1////"
            s = s & "/////////////////////////////////vby7Gsv62Ml62Ml7nZA/vv5////////////////////////////////////9rqg"
            s = s & "62Ml62Ml7ndA/vby++LX62009bGS////////////////////////////////////////++PY62Uo62Ml62Ml7Gkt86uK/vby"
            s = s & "/////////////////////Ozl8plw62Ml62Ml7ndA/Ozl/Ozl7ndA62Ml9KN/////////////////////////////////////"
            s = s & "////////+t/R7Gkv62Ml62Ml62Ml62Qn7ntG8Ipc9KeF9a6M74BO62Ml62Ml62Ml6200/Ofe/vby74BO62Ml62Ml8pVq////"
            s = s & "/////////////////////////////////////////////fby8IdW62Mk62Ml62Ml62Ml62Ml62Ml62Ml62Ml62Ml62Ml8plw"
            s = s & "/fHr/////Ofe74VW62Ml62Ml74VW/////////////////////////////////////////////////////////vby8qaC7XQ8"
            s = s & "62Qm62Ml62Ml62Ml74pd9a6M/Ofe////////////////////9bWZ7Gkt7ndA////////////////////////////////////"
            s = s & "/////////////////////////////////////////////////////////////////////////////////////Ofe74VW////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
        Case "template.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////9fPx7Ojk7Ojk7Ojk7Ojk7Ojk7Ojk"
            s = s & "7Ojk7Ojk7Ojk7Ojk7Ojk7Ojk7Ojk7Ojk7Ojk7Ojk7Ojk+Pb1////////////////////////////////////////////+/r6"
            s = s & "hHpzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzVUEzlIuE5N7Y////////////////"
            s = s & "////////////////////////////urGts7Cu////////////////////////////////////////////////////////////"
            s = s & "////////jYWBzcXA////////////////////////////////////////////qqCZw8C+////////////////////////////"
            s = s & "////////////////////////////////////////lI2IycK8////////////////////////////////////////////qqCZ"
            s = s & "w8C+////////////////////////////////////////////////////////////////////lI2IycK8////////////////"
            s = s & "////////////////////////////qqCZqKXPlJTylJTylJTylJTywMD3zs74lJTylJTylJTylJTyzs74wMD3lJTylJTylJTy"
            s = s & "lJTy2tr6lI2IycK8////////////////////////////////////////////qqCZlZPYrKz14uL74uL7ycn4lJTyrKz1vLz3"
            s = s & "4uL74uL7vLz3rKz1lJTyycn44uL74uL7rKz1wMD3lI2IycK8////////////////////////////////////////////qqCZ"
            s = s & "lZPYwMD3////////4uL7lJTyrKz10tL5////////0tL5rKz1lJTy4uL7////////wMD3wMD3lI2IycK8////////////////"
            s = s & "////////////////////////////qqCZlZPYwMD3////////4uL7lJTyrKz10tL5////////0tL5rKz1lJTy4uL7////////"
            s = s & "wMD3wMD3lI2IycK8////////////////////////////////////////////qqCZlZPYwMD3////////4uL7lJTyrKz10tL5"
            s = s & "////////0tL5rKz1lJTy4uL7////////wMD3wMD3lI2IycK8////////////////////////////////////////////qqCZ"
            s = s & "lZPYlJTywMD3wMD3rKz1lJTyrKz1oKDzwMD3wMD3oKDzrKz1lJTyrKz1wMD3wMD3lJTywMD3lI2IycK8////////////////"
            s = s & "////////////////////////////qqCZqqfOdHTwdHTwdHTwdHTwubn2y8v4dHTwdHTwdHTwdHTwy8v4ubn2dHTwdHTwdHTw"
            s = s & "dHTw3Nz6lI2IycK8////////////////////////////////////////////qqCZlZPYrKz14uL74uL7ycn4lJTyrKz1vLz3"
            s = s & "4uL74uL7vLz3rKz1lJTyycn44uL74uL7rKz1wMD3lI2IycK8////////////////////////////////////////////qqCZ"
            s = s & "lZPYwMD3////////4uL7lJTyrKz10tL5////////0tL5rKz1lJTy4uL7////////wMD3wMD3lI2IycK8////////////////"
            s = s & "////////////////////////////qqCZlZPYwMD3////////4uL7lJTyrKz10tL5////////0tL5rKz1lJTy4uL7////////"
            s = s & "wMD3wMD3lI2IycK8////////////////////////////////////////////qqCZlZPYwMD3////////4uL7lJTyrKz10tL5"
            s = s & "////////0tL5rKz1lJTy4uL7////////wMD3wMD3lI2IycK8////////////////////////////////////////////qqCZ"
            s = s & "lZPYlJTywMD3wMD3rKz1lJTyrKz1oKDzwMD3wMD3oKDzrKz1lJTyrKz1wMD3wMD3lJTywMD3lI2IycK8////////////////"
            s = s & "////////////////////////////qqCZq6jNk5Pzk5Pzk5Pzk5PzwsL30ND5k5Pzk5Pzk5Pzk5Pz0ND5wsL3k5Pzk5Pzk5Pz"
            s = s & "k5Pz3t76lI2IycK8////////////////////////////////////////////qqCZlZPYlJTywMD3wMD3rKz1lJTyrKz1oKDz"
            s = s & "wMD3wMD3oKDzrKz1lJTyrKz1wMD3wMD3lJTywMD3lI2IycK8////////////////////////////////////////////qqCZ"
            s = s & "lZPYwMD3////////4uL7lJTyrKz10tL5////////0tL5rKz1lJTy4uL7////////wMD3wMD3lI2IycK8////////////////"
            s = s & "////////////////////////////qqCZlZPYwMD3////////4uL7lJTyrKz10tL5////////0tL5rKz1lJTy4uL7////////"
            s = s & "wMD3wMD3lI2IycK8////////////////////////////////////////////qqCZlZPYwMD3////////4uL7lJTyrKz10tL5"
            s = s & "////////0tL5rKz1lJLx4dTn/u/k/u/kv7Tqv7Tqr6ScycK8////////////////////////////////////////////qqCZ"
            s = s & "lZPYrKz14uL74uL7ycn4lJTyrKz1vLz34uL74uL7vLz3rKz1lJHwybnk4dDg4dDgq5/nv7rwlI2IycK8////////////////"
            s = s & "////////////////////////////qqCZtLHIlJTylJTylJTylJTy0tL53t76lJTylJTylJTylJTy3t760s70k4npk4npk4np"
            s = s & "k4/v6ur8lI2IycK8////////////////////////////////////////////qqCZw8C+////////////////////////////"
            s = s & "//////////////////r3/urb/urb/vfy////////lI2IycK8////////////////////////////////////////////qqCZ"
            s = s & "w8C+//////////////////////////////////////////////r3/urb/vfy////////////lI2I0MrF////////////////"
            s = s & "////////////////////////////4NzagHVv3NrZ4+Li4+Li4+Li4+Li4+Li4+Li4+Li4+Li4+Li4+Li6uTg7eXf4+Li4+Li"
            s = s & "4+Li1NLRdmpk8O7t////////////////////////////////////////////////1dDNlIyGi4F7i4F7i4F7i4F7i4F7i4F7"
            s = s & "i4F7i4F7i4F7i4F7i4F7i4F7i4F7i4F7i4F7nZWQ4Nza////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
        Case "folder.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////mLjPKXaoJnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSo"
            s = s & "JnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSoJnSoY5y/+Pv7////////////////jLLMCnCvCJDgCJDg"
            s = s & "CJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgCJDgBni+V5S7"
            s = s & "////////////////D2CWCpnuC571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571CpzyBFaO////////////////AFKKC571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571AFKK////////////////AFKKC571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571AFKK"
            s = s & "////////////////AFKKC571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571AFKK////////////////AFKKC571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571AFKK////////////////AFKKC571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571AFKK"
            s = s & "////////////////AFKKC571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571AFKK////////////////AFKKC571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571AFKK////////////////AFKKC571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571AFKK"
            s = s & "////////////////AFKKC571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571AFKK////////////////AFKKC571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571AFKK////////////////AFKKC571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571AFKK"
            s = s & "////////////////BVaOCpzyC571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571AFKK////////////////Zpy/CH7FC571C571C571C571C571C571C571C571C571C571"
            s = s & "C571C571C571C571C571C571C571C571C571C571C571C571C571C571BoXSQoWx////////////////8PX4RpS8DGWeEHKu"
            s = s & "EHKuEHKuEHKuEHKuEHKuEHKuEHKuEHKuEHKuEXGtEXCsEXCsEXCsEXCsEXCsEXCsEXCsEXCsEXCsEXCsEXCsDWScOYCt4Orw"
            s = s & "////////////////////pun9SMjxQ73mQ73mQ73mQ73mQ73mQ73mQ73mQ73mQ73mQ73mhczmwNXiwNXiwNXiwNXiwNXiwNXi"
            s = s & "wNXiwNXiwNXiwNXiwNXi4Orw////////////////////////////3vf+TdP8TdP8TdP8TdP8TdP8TdP8TdP8TdP8TdP8TdP8"
            s = s & "TdP8ve7+////////////////////////////////////////////////////////////////////////////////wvD+f+D9"
            s = s & "et39et39et39et39et39et39et39f+D9ve7++v7/////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
        Case "run.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "/////////////////////////////////////////////////////v/+9vvz8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv"
            s = s & "8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv9fry/f79/////////////////////v/+8vjt7/jq7/jq"
            s = s & "7/jq7/jq1erJ6PTh7/jq7/jq7/jq7/jqzOa+7/jq7/jq7/jq7/jq6PTh1erJ7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq/f79"
            s = s & "////////////////+Pz27/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jq7/jq7/jq7/jqx+O37/jq7/jq7/jq7/jq5vPf0ejF7/jq"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq7/jq9vvz////////////////8/vx7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jq7/jq7/jq7/jq"
            s = s & "x+O37/jq7/jq7/jq7/jq5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv7/jq7/jq7/jq"
            s = s & "7/jq7/jq0ejF5vPf7/jq3urZ7/jq7/jqx+O37/jq7/jq7/jq7/jq5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv"
            s = s & "////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jqMXoOja2F6PPjx+O37/jq7/jq7/jq7/jq5vPf0ejF7/jq"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jqMXoOTJwuaZZe"
            s = s & "ttSn7/jq7/jq7/jq7/jq5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv7/jq7/jq7/jq"
            s = s & "7/jq7/jq0ejF5vPf7/jqMXoOSqggR6IdTow5wdO77/jq7/jq7/jq5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv"
            s = s & "////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jqMXoOSqggSqggSqcfVZVDlLGM7PXn7/jq5vPf0ejF7/jq"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv1+vM0ejF0ejF0ejF0ejFyeS6z+bB0ejFMXoOSqggSqgg"
            s = s & "SqggSqggSJooYZRSw9u3z+bByeS60ejF0ejF0ejF0ejF0ejF0ejF0ejF1+vM8vrv////////////////8vrv6vXk5vPf5vPf"
            s = s & "5vPf5vPfz+fC3+/W5vPfMXoOSqggSqggSqggSqggSqggRqEdWpBMuM+vz+fC5vPf5vPf5vPf5vPf5vPf5vPf5vPf6vXk8vrv"
            s = s & "////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jqMXoOSqggSqggSqggSqggSqggSqggSaYfVJRCjbGB7PXn"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jqMXoOSqggSqgg"
            s = s & "SqggSqggSqggSqggSqggSqggS5kvdZ1r4u3d7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv7/jq7/jq7/jq"
            s = s & "7/jq7/jq0ejF5vPf7/jqMXoOSqggSqggSqggSqggSqggSqggSqggSqggSqggRp8dW49M2+fV7/jq7/jq7/jq7/jq7/jq8vrv"
            s = s & "////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jqMXoOSqggSqggSqggSqggSqggSqggSqggSqggSqggT5k3"
            s = s & "dZ1r6PPj7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv0ejFx+O3x+O3x+O3x+O3x+O3x+O3x+O3MXoOSqggSqgg"
            s = s & "SqggSqggSqggSqggSqggSaYfTpA3h693xOG0x+O3x+O3x+O3x+O3x+O30ejF8vrv////////////////8vrv7/jq7/jq7/jq"
            s = s & "7/jq7/jq0ejF5vPf7/jqMXoOSqggSqggSqggSqggSqggSqggRqEdWpBMsMyk7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv"
            s = s & "////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jqMXoOSqggSqggSqggSqggSqggUZs4bZhi1uXP0ejF7/jq"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jqMXoOSqggSqgg"
            s = s & "SqggSqcfVZVDmbSS7PXn5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////8vrv7/jq7/jq7/jq"
            s = s & "7/jq7/jq0ejF5vPf7/jqMXoOSqggSqggR6IdVo1HwdO77/jq7/jq5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv"
            s = s & "////////////////8vrv6PTh5vPf5vPf5vPf5vPfz+fC3+/W5vPfMXoOSqggS5wsWpFJ1uXP5vPf5vPf5vPf3+/Wz+bB5vPf"
            s = s & "5vPf5vPf5vPf5vPf5vPf5vPf6PTh8vrv////////////////8vrv2u3P0ejF0ejF0ejF0ejFyeS6z+bB0ejFMXoOUJU7faZv"
            s = s & "wt6y0ejF0ejF0ejF0ejFz+fCyeS60ejF0ejF0ejF0ejF0ejF0ejF0ejF2u3P8vrv////////////////8vrv7/jq7/jq7/jq"
            s = s & "7/jq7/jq0ejF5vPf7/jqVYpIu8607/jqx+O37/jq7/jq7/jq7/jq5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv"
            s = s & "////////////////8vrv7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jq7/jq7/jq7/jqx+O37/jq7/jq7/jq7/jq5vPf0ejF7/jq"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq7/jq8vrv////////////////9Pvy7/jq7/jq7/jq7/jq7/jq0ejF5vPf7/jq7/jq7/jq7/jq"
            s = s & "x+O37/jq7/jq7/jq7/jq5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq9Prw////////////////+v347/jq7/jq7/jq"
            s = s & "7/jq7/jq0ejF5vPf7/jq7/jq7/jq7/jqx+O37/jq7/jq7/jq7/jq5vPf0ejF7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq9/v1"
            s = s & "////////////////////9vz07/jq7/jq7/jq7/jq3O7S6vXk7/jq7/jq7/jq7/jq1+vM7/jq7/jq7/jq7/jq6vXk3O7S7/jq"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq8vrv/v/+////////////////////////+/369vz08vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv"
            s = s & "8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv8vrv9vz0+fz3////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
        Case "close.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "+vr/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+fn/////////////////"
            s = s & "/////////////////////////f3/6+v+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+6Oj9+/v/////////////////////////////////7e3+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+6en+/////////////////////////////v7/4+P94uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL++/v/////"
            s = s & "////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+z8/4Xl7hz8/44uL+4uL+4uL+4uL+4uL+4uL+1dX6Z2fixMT14uL+4uL+4uL+4uL+4uL+4uL++Pj/////"
            s = s & "////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+z8/4SUndJibcSUndz8/44uL+4uL+4uL+4uL+1dX6Xl7hJibcOjrc"
            s = s & "xMT14uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+j4/oJibcJibcJibcSUnd"
            s = s & "z8/44uL+4uL+1dX6Xl7hJibcJibcJibcdnbk4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+"
            s = s & "4uL+4uL+4uL+39/9g4PmJibcJibcJibcSUndz8/41dX6Xl7hJibcJibcJibcb2/j3Nz84uL+4uL+4uL+4uL+4uL++Pj/////"
            s = s & "////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+4uL+39/9g4PmJibcJibcJibcRUXdV1fgJibcJibcJibcb2/j3Nz8"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+4uL+4uL+39/9g4PmJibc"
            s = s & "JibcJibcJibbJibbJibcb2/j3Nz84uL+4uL+4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL+39/9fHzlJibcJibcJibcJibbZmbh3Nz84uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL++Pj/////"
            s = s & "////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+1dX6XV3hJSXcJibcJibcJibcR0fdz8/44uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+4uL+4uL+1dX6Xl7hJibc"
            s = s & "JibcJSXcJibcJibcJibcSUndz8/44uL+4uL+4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+1dX6Xl7hJibcJibcJibcbW3jgIDmJibcJibcJibcSUndz8/44uL+4uL+4uL+4uL+4uL+4uL++Pj/////"
            s = s & "////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+1dX6Xl7hJibcJibcJibcb2/j3Nz839/9g4PmJibcJibcJibcSUnd"
            s = s & "z8/44uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+lZXqJibcJibcJibcb2/j"
            s = s & "3Nz84uL+4uL+39/9g4PmJibcJibcJibcfX3l4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+"
            s = s & "4uL+4uL+4uL+39/9g4PmJibcb2/j3Nz84uL+4uL+4uL+4uL+39/9g4PmJibcb2/j3Nz84uL+4uL+4uL+4uL+4uL++Pj/////"
            s = s & "////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+4uL+39/9n5/s3Nz84uL+4uL+4uL+4uL+4uL+4uL+39/9n5/s3Nz8"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+Pj/4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL++Pj/////////////////////////+fn/4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL++Pj/////"
            s = s & "////////////////////////5OT94uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL+/f3/////////////////////////////8/P+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+"
            s = s & "4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+7+/+////////////////////////////////////9PT+"
            s = s & "5eX94uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+4uL+5OT98PD+/v7/////////"
            s = s & "////////////////////////////////////+/v/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/+Pj/"
            s = s & "+Pj/+Pj/+/v/////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
        Case "copies.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////1uDMe5pgWoUvWoUvWoUvWoUvWoUvWoUvWoUvaI9HrcGd"
            s = s & "////////////////////////////////////////////////////////////////////////////////8PPsd5dd4Ova7/jq"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq6fPknbOPwtGy////////////////////////////////////////////////6/DmssWdrMKX"
            s = s & "rMKXrMKXrMKXrMKXh6dqjad70t/L7/jq7/jq7/jq7/jq7/jq7/jq7/jq7/jq0t/Lj6p6////////////////////////////"
            s = s & "////////////////1uDMfpxntMaruMmuuMmuuMmuuMmuuMmukquCkquCmbCKrL+g7/jq7/jq7/jq7/jq7/jq7/jq7/jq1uLP"
            s = s & "i6d1////////////////////////////////////////////hKNu2eXS7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7PXne5lk"
            s = s & "5vDg7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1////////////////////////////////////////////WoUv7/jq7/jq7/jq"
            s = s & "7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1////////////////////////////"
            s = s & "////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLP"
            s = s & "i6d1////////////////////////////////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD"
            s = s & "1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1////////////////////////////////////////////WoUv7/jq7/jq7/jq"
            s = s & "7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1OHMttagsdOasdOasdOasdOasdOasdOasdOasdOasdOasdOa6/Tm////////////"
            s = s & "////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLP"
            s = s & "i6d1////////////////////////////////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD"
            s = s & "1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1////////////////////////////////////////////WoUv7/jq7/jq7/jq"
            s = s & "7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1////////////////////////////"
            s = s & "////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLP"
            s = s & "i6d1////////////////////////////////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD"
            s = s & "1OHMttagsdOasdOasdOasdOasdOasdOasdOasdOasdOasdOa6/Tm////////////////////////////WoUv7/jq7/jq7/jq"
            s = s & "7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1////////////////////////////"
            s = s & "////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLP"
            s = s & "i6d1////////////////////////////////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD"
            s = s & "1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1////////////////////////////////////////////WoUv7/jq7/jq7/jq"
            s = s & "7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1////////////////////////////"
            s = s & "////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD1OHMttagsdOasdOasdOasdOasdOasdOa"
            s = s & "sdOasdOasdOasdOa6/Tm////////////////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jquMmuuMmu7/jqk6yD"
            s = s & "1uLP7/jq7/jq7/jq7/jq7/jq7/jq1uLPi6d1q9aWSqggSqggYbM90+nI////////////////////////WoUv7/jq7/jq7/jq"
            s = s & "7/jq7/jq7/jq7/jqyNa/p7ub7/jqk6yD1uLP7/jq7/jq7/jq7/jq7/jq7/jqy9nDd65fSqggSqggSqggSqggUKso7/fr////"
            s = s & "////////////////WoUv7/jq7/jq7/jq7/jq7/jq7/jq7/jq7PXnh6N0y9nDkKqA1uLP7/jq7/jq7/jq7/jq7/jq1uLPf5xo"
            s = s & "ZbNCir5/Xq1GfbhvSqggSqggr9md////////////////////YYo87PXn7/jq7/jq7/jq7/jq7/jq7/jq7/jq5vDgoreVXoc4"
            s = s & "WoUvWoUvWoUvWoUvWoUvWoUvi6d10dzFYbM9qs2j2ufXfbhvSqggSqggq9aW////////////////////la6Az9zH7/jq7/jq"
            s = s & "7/jq7/jq7/jq7/jq7/jq7/jq5vDgdpZZ+/v5////////////////////////////lcx8lcOM7vTtb7NdSqggSqgg3u/W////"
            s = s & "////////////////8PPsgJ9onbOPuMmuuMmuuMmuuMmuuMmuuMmup7yafpxnxdO7////////////////////////////////"
            s = s & "+vz5ps2dXq1Gb7NdUKsotdul////////////////////////////+/v5zNm/rMKXrMKXrMKXrMKXrMKXrMKXwtGy6/Dm////"
            s = s & "////////////////////////////////////////0+nI0+nI6fTk////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
        Case "pages.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "/////Ofe9bGS9bGS/Ofe////////////////////////////////8O7ttaumqqCZqqCZqqCZqqCZqqCZqqCZy8TA////////"
            s = s & "////z8rGqqCZqqCZqqCZqqCZqqCZqqCZyZmG62Ml62Ml62Ml6200/Ozl////////////////////////////fnNsubWzw8C+"
            s = s & "w8C+w8C+w8C+w8C+w8C+l5CLvbaz////y8TAh356w8C+w8C+w8C+w8C+w8C+w8C+2o9+62Ml7X5f62Ml62Ml9bWZ////////"
            s = s & "////////////////29bTjoeD////////////////////////////4+Lii4F7////j4aA4N7d////////////////////////"
            s = s & "9Lit99PN+NzX7X5f62Ml9bWZ////////////////////////1dDNlo+K9/Xz6eHZ6eHZ6eHZ6eHZ8u3q////4+Lii4F7////"
            s = s & "i4F74+Li7ebg6eHZ6eHZ6eHZ6eHZ/fz8++zp8aiY9svD62Ml6200/fHr////////////////////////1dDNlo+K////9/Xz"
            s = s & "9/Xz9/Xz9/Xz/fz8////4+Lii4F7////i4F74+Li+/r59/Xz9/Xz9/Xz+fj2////////qImA8M/J9bKT/Ofe////////////"
            s = s & "////////////////1dDNlo+K+fj28evn8evn8evn8evn8evn8u3q4+Lii4F7////i4F74+Li8u3q8evn8evn8evn8evn8evn"
            s = s & "+fj2lo+K1dDN////////////////////////////////////1dDNlo+K/fz88evn8evn8evn8evn8evn9vPw4+Lii4F7////"
            s = s & "i4F74+Li9vPw8evn8evn8evn8evn8evn/fz8lo+K1dDN////////////////////////////////////1dDNlo+K/fz89/Xz"
            s = s & "9/Xz9/Xz9/Xz9/Xz+fj24+Lii4F7////i4F74+Li+fj29/Xz9/Xz9/Xz9/Xz9/Xz/fz8lo+K1dDN////////////////////"
            s = s & "////////////////1dDNlo+K+fj26eHZ6eHZ6eHZ6eHZ6eHZ7+nk4+Lii4F7////i4F74+Li7+nk6eHZ6eHZ6eHZ6eHZ6eHZ"
            s = s & "+fj2lo+K1dDN////////////////////////////////////29bTjoeD////////////////////////////4+Lii4F7////"
            s = s & "j4aA4N7d////////////////////////////lo+K1dDN////////////////////////////////////////fnNswb684+Li"
            s = s & "4+Li4+Li4+Li4+Li3NrZl5CLvbaz////y8TAh3563NrZ4+Li4+Li4+Li4+Li4+LixsPCe29o+/r6////////////////////"
            s = s & "////////////////////8O7tsKiki4F7i4F7i4F7i4F7i4F7lIyGy8TA////////////z8rGlIyGi4F7i4F7i4F7i4F7i4F7"
            s = s & "qqOe8O7t////////////////////////////////////////////////4Nza1dDN1dDN1dDN1dDN1dDN1dDN8O7t////////"
            s = s & "////9fTz1dDN1dDN1dDN1dDN1dDN1dDN29bT////////////////////////////////////////////////mZCLh395lo+K"
            s = s & "lo+Klo+Klo+Klo+Klo+Kd2tl1dDN////4NzabWFXlo+Klo+Klo+Klo+Klo+Klo+KjoeDj4aA////////////////////////"
            s = s & "////////////////5eLgf3Vv////////////////////////////4N7dj4aA////mZCL2NbV////////////////////////"
            s = s & "////joeD29bT////////////////////////////////////1dDNlo+K+fj28evn8evn8evn8evn9vPw////4+Lii4F7////"
            s = s & "i4F74+Li8u3q8evn8evn8evn8evn/fz8////lo+K1dDN////////////////////////////////////1dDNlo+K/fz88evn"
            s = s & "8evn8evn8evn+fj2////4+Lii4F7////i4F74+Li9vPw8evn8evn8evn8u3q////////lo+K1dDN////////////////////"
            s = s & "////////////////1dDNlo+K/fz89/Xz9/Xz9/Xz9/Xz9/Xz+fj24+Lii4F7////i4F74+Li+fj29/Xz9/Xz9/Xz9/Xz9/Xz"
            s = s & "/fz8lo+K1dDN////////////////////////////////////1dDNlo+K+fj26eHZ6eHZ6eHZ6eHZ6eHZ7+nk4+Lii4F7////"
            s = s & "i4F74+Li7+nk6eHZ6eHZ6eHZ6eHZ6eHZ+fj2lo+K1dDN////////////////////////////////////1dDNlo+K////////"
            s = s & "////////////////////4+Lii4F7////i4F74+Li////////////////////////////lo+K1dDN////////////////////"
            s = s & "////////////////1dDNlo+K9/Xz4dXL4dXL4dXL4dXL4dXL6eHZ4+Lii4F7////i4F74+Li6eHZ4dXL4dXL4dXL4dXL4dXL"
            s = s & "9/Xzlo+K1dDN////////////////////////////////////1dDNlo+K////////////////////////////4+Lii4F7////"
            s = s & "i4F74+Li////////////////////////////lo+K1dDN////////////////////////////////////+/r6dGhh4+Li////"
            s = s & "////////////////+Pj4ubWzp5+b////uLGsqaSh+Pj4////////////////////5+bld2tl8O7t////////////////////"
            s = s & "////////////////////1dDNi4F7VUEzVUEzVUEzVUEzVUEzZldMp5+b+/r6////+/r6qqOeZldMVUEzVUEzVUEzVUEzVUEz"
            s = s & "hHp11dDN////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
        Case "landscape.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////8+7Z0r5txq1G"
            s = s & "xq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gxq1Gz7tm7ufK////"
            s = s & "/////////////////fz4uJokspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEI"
            s = s & "spEIspEIspEIspEIspEIspEIs5MX+PXp////////////////6eC6spEI1see//7s//7s//7s//7s//7s//7s//7s//7s//7s"
            s = s & "//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s4de3spEI5Nmq////////////////4tWispEI5t2///7s"
            s = s & "//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s5t2/spEI4tWi"
            s = s & "////////////////4tWispEI5t2///7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s"
            s = s & "//7s//7s//7s//7s//7s5t2/spEI4tWi////////////////4tWispEI5t2///7s//7s//7s//7s//7s//7s//7s//7s//7s"
            s = s & "//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s5t2/spEI4tWi////////////////4tWispEI5t2///7s"
            s = s & "//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s5t2/spEI4tWi"
            s = s & "////////////////4tWispEI5t2///7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s"
            s = s & "//7s//7s//7s//7s//7s5t2/spEI4tWi////////////////4tWispEI5t2///7s//7s//7s//7s//7s//7s//7s//7s//7s"
            s = s & "//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s5t2/spEI4tWi////////////////4tWispEI5t2///7s"
            s = s & "//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s5t2/spEI4tWi"
            s = s & "////////////////4tWispEI5t2///7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s//7s"
            s = s & "//7s//7s//7s//7s//7s5t2/spEI4tWi////////////////8OrRspEIxrB37+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ"
            s = s & "7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ7+nQ0cCTspEI6eC6////////////////////xq1GspEIspEI"
            s = s & "spEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIspEIvKAu/fz4"
            s = s & "////////////////////////59yy2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD2MiD"
            s = s & "2MiD2MiD2MiD2MiD2MiD4tWi/fz4////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
        Case "portrait.bmp"
            s = ""
            s = s & "Qk02DAAAAAAAADYAAAAoAAAAIAAAACAAAAABABgAAAAAAAAAAABNEAAATRAAAAAAAAAAAAAA////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////9fnxtdmgpNGKpNGK"
            s = s & "pNGKpNGKpNGKpNGKpNGKpNGKr9aZ7/bq////////////////////////////////////////////////////////////////"
            s = s & "////////////7/bqUaYfSqMWSqMWSqMWSqMWSqMWSqMWSqMWSqMWSqMWSqMWSqMW3u7U////////////////////////////"
            s = s & "////////////////////////////////////////////q9ORSqMWsNGn2evT2evT2evT2evT2evT2evT2evT2evTxt6/SqMW"
            s = s & "msx9////////////////////////////////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w"
            s = s & "9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////////////////////////////////////"
            s = s & "////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////"
            s = s & "////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMW"
            s = s & "jsZv////////////////////////////////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w"
            s = s & "9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////////////////////////////////////"
            s = s & "////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////"
            s = s & "////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMW"
            s = s & "jsZv////////////////////////////////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w"
            s = s & "9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////////////////////////////////////"
            s = s & "////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////"
            s = s & "////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMW"
            s = s & "jsZv////////////////////////////////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w"
            s = s & "9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////////////////////////////////////"
            s = s & "////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////"
            s = s & "////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMW"
            s = s & "jsZv////////////////////////////////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w"
            s = s & "9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////////////////////////////////////"
            s = s & "////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////"
            s = s & "////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMW"
            s = s & "jsZv////////////////////////////////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w"
            s = s & "9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////////////////////////////////////"
            s = s & "////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////"
            s = s & "////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMW"
            s = s & "jsZv////////////////////////////////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w"
            s = s & "9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////////////////////////////////////"
            s = s & "////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////"
            s = s & "////////////////////////////////////////////jsZvSqMW5/Ti9P3w9P3w9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMW"
            s = s & "jsZv////////////////////////////////////////////////////////////////////////jsZvSqMW5PHf9P3w9P3w"
            s = s & "9P3w9P3w9P3w9P3w9P3w9P3w5/TiSqMWjsZv////////////////////////////////////////////////////////////"
            s = s & "////////////x+K3SqMWibt7udaxudaxudaxudaxudaxudaxudaxudaxnMWQSqMWq9OR////////////////////////////"
            s = s & "////////////////////////////////////////////+vz5fb1ZSqMWSqMWSqMWSqMWSqMWSqMWSqMWSqMWSqMWSqMWbLRD"
            s = s & "7/bq////////////////////////////////////////////////////////////////////////////////5PDb0+jF0+jF"
            s = s & "0+jF0+jF0+jF0+jF0+jF0+jF3u7U////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////////////////////////////////////////////////////////////"
            s = s & "////////////////////////////////////////"
            EmbeddedIconBase64 = s
    End Select
End Function

Private Sub WriteBase64BinaryFile(ByVal filePath As String, ByVal encoded As String)
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim stream As Object

    On Error GoTo EH
    If Len(encoded) = 0 Then Exit Sub

    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = encoded

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write xmlNode.nodeTypedValue
    stream.SaveToFile filePath, 2
    stream.Close
    Exit Sub

EH:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
End Sub

Private Function CombineLocalPath(ByVal folderPath As String, ByVal fileName As String) As String
    If Right$(folderPath, 1) = "\" Then
        CombineLocalPath = folderPath & fileName
    Else
        CombineLocalPath = folderPath & "\" & fileName
    End If
End Function

Private Sub EnsureLocalFolderExists(ByVal folderPath As String)
    On Error Resume Next
    If Len(folderPath) > 0 Then
        If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    End If
    On Error GoTo 0
End Sub

Private Sub LoadLayouts()
    Dim layouts As Variant
    Dim i As Long

    On Error GoTo EH

    mCachedLayoutName = ""
    mCachedCapacity = 0
    cboLayout.Clear
    layouts = GetMimakiTemplateLayoutNamesV21()
    If IsArray(layouts) Then
        For i = LBound(layouts) To UBound(layouts)
            cboLayout.AddItem CStr(layouts(i))
        Next i
    End If

    If cboLayout.ListCount = 0 Then cboLayout.AddItem "KARTA"
    cboLayout.ListIndex = 0
    Exit Sub

EH:
    cboLayout.Clear
    cboLayout.AddItem "KARTA"
    cboLayout.ListIndex = 0
    SetStatus "Sablonu sa nepodarilo nacitat, pouzity fallback KARTA."
End Sub

Private Sub UpdateLayoutInfo()
    Dim layoutName As String
    Dim capacity As Long
    Dim pagesNeeded As Long
    Dim workflowMode As Long
    Dim copiesCount As Long
    Dim startSlot As Long
    Dim firstPage As Long
    Dim lastPage As Long
    Dim slotWidth As Double
    Dim slotHeight As Double
    Dim graphicWidth As Double
    Dim graphicHeight As Double
    Dim sizeInfo As String

    On Error GoTo EH
    If mUpdating Then Exit Sub

    layoutName = Trim$(cboLayout.Value)
    copiesCount = ParsePositiveLong(txtCopies.Text, 50)
    startSlot = ParsePositiveLong(txtStartSlot.Text, 1)
    workflowMode = GetSelectedSourceWorkflow()
    ParsePageRange Trim$(txtPageRange.Text), firstPage, lastPage

    capacity = GetCachedLayoutCapacity(layoutName)
    pagesNeeded = EstimatePagesFromCapacity(capacity, workflowMode, copiesCount, startSlot, firstPage, lastPage)

    If capacity > 0 Then
        If GetCachedLayoutSlotInfo(layoutName, slotWidth, slotHeight, graphicWidth, graphicHeight) Then
            sizeInfo = vbCrLf & "Slot: " & FormatMmPair(slotWidth, slotHeight) & _
                       " | Grafika + spad: " & FormatMmPair(graphicWidth, graphicHeight)
        End If
        lblLayoutInfo.Caption = layoutName & ": " & CStr(capacity) & " pozicii na harok" & vbCrLf & _
                                "Odhad: " & CStr(pagesNeeded) & " stran" & sizeInfo
    Else
        lblLayoutInfo.Caption = layoutName & ": produkt sa v sablone nenasiel"
    End If
    Exit Sub

EH:
    lblLayoutInfo.Caption = "Informacie o produkte sa nepodarilo nacitat."
End Sub

Private Sub RefreshPanelForLayoutChange()
    On Error GoTo CleanUp
    If mUpdating Then Exit Sub

    mCachedLayoutName = ""
    mCachedCapacity = 0
    mCachedSlotInfoLayoutName = ""
    mUpdating = True
    LoadVectorSetting Trim$(cboLayout.Value)
    UpdateSourceModeButtons

CleanUp:
    mUpdating = False
    ApplyFullSheetCopiesDefault
    UpdateLayoutInfo
End Sub

Private Sub ApplyFullSheetCopiesDefault()
    Dim layoutName As String
    Dim capacity As Long

    On Error Resume Next
    If Not optCopies.Value Then Exit Sub

    layoutName = Trim$(cboLayout.Value)
    If Len(layoutName) = 0 Then Exit Sub

    capacity = GetCachedLayoutCapacity(layoutName)
    If capacity > 0 Then txtCopies.Text = CStr(capacity)
    On Error GoTo 0
End Sub

Private Function GetCachedLayoutCapacity(ByVal layoutName As String) As Long
    layoutName = UCase$(Trim$(layoutName))
    If Len(layoutName) = 0 Then Exit Function

    If layoutName <> mCachedLayoutName Or mCachedCapacity <= 0 Then
        mCachedLayoutName = layoutName
        mCachedCapacity = GetMimakiLayoutSlotCountV21(layoutName)
    End If

    GetCachedLayoutCapacity = mCachedCapacity
End Function

Private Function GetCachedLayoutSlotInfo(ByVal layoutName As String, ByRef slotWidth As Double, ByRef slotHeight As Double, ByRef graphicWidth As Double, ByRef graphicHeight As Double) As Boolean
    layoutName = UCase$(Trim$(layoutName))
    If Len(layoutName) = 0 Then Exit Function

    If layoutName <> mCachedSlotInfoLayoutName Or mCachedSlotWidth <= 0# Or mCachedSlotHeight <= 0# Then
        mCachedSlotInfoLayoutName = layoutName
        mCachedSlotWidth = 0#
        mCachedSlotHeight = 0#
        mCachedGraphicWidth = 0#
        mCachedGraphicHeight = 0#
        If Not GetMimakiLayoutFirstSlotSizeV21(layoutName, mCachedSlotWidth, mCachedSlotHeight, mCachedGraphicWidth, mCachedGraphicHeight) Then Exit Function
    End If

    slotWidth = mCachedSlotWidth
    slotHeight = mCachedSlotHeight
    graphicWidth = mCachedGraphicWidth
    graphicHeight = mCachedGraphicHeight
    GetCachedLayoutSlotInfo = True
End Function

Private Function FormatMmPair(ByVal widthMm As Double, ByVal heightMm As Double) As String
    FormatMmPair = FormatMmValue(widthMm) & " x " & FormatMmValue(heightMm) & " mm"
End Function

Private Function FormatMmValue(ByVal valueMm As Double) As String
    If Abs(valueMm - Round(valueMm, 0)) < 0.005 Then
        FormatMmValue = Format$(valueMm, "0")
    Else
        FormatMmValue = Format$(valueMm, "0.##")
    End If
End Function

Private Function EstimatePagesFromCapacity(ByVal capacity As Long, ByVal workflowMode As Long, ByVal copiesCount As Long, ByVal startSlot As Long, ByVal firstPage As Long, ByVal lastPage As Long) As Long
    Dim totalItems As Long
    Dim pageCount As Long
    Dim isCardLayout As Boolean
    Dim preserveVector As Boolean

    If capacity <= 0 Then Exit Function
    If startSlot < 1 Then startSlot = 1
    isCardLayout = IsCardLayoutName(cboLayout.Value)

    If workflowMode = mki23SwCopies Then
        totalItems = copiesCount
    ElseIf Not mSourceDoc Is Nothing Then
        NormalizePanelPageRange mSourceDoc.Pages.Count, firstPage, lastPage
        If isCardLayout Then
            totalItems = GetWorkflowCardCount(workflowMode, firstPage, lastPage)
        Else
            pageCount = (lastPage - firstPage) + 1
            If pageCount > 0 Then totalItems = pageCount
        End If
    ElseIf Documents.Count > 0 Then
        NormalizePanelPageRange ActiveDocument.Pages.Count, firstPage, lastPage
        If isCardLayout Then
            totalItems = GetWorkflowCardCount(workflowMode, firstPage, lastPage)
        Else
            pageCount = (lastPage - firstPage) + 1
            If pageCount > 0 Then totalItems = pageCount
        End If
    End If

    If totalItems <= 0 Then Exit Function
    EstimatePagesFromCapacity = ((startSlot + totalItems - 2) \ capacity) + 1
End Function

Private Function IsCardLayoutName(ByVal layoutName As String) As Boolean
    Dim s As String

    s = UCase$(Trim$(layoutName))
    s = Replace$(s, " ", "")
    s = Replace$(s, "_", "")
    s = Replace$(s, "-", "")
    If Left$(s, 8) = "TEMPLATE" Then s = Mid$(s, 9)

    IsCardLayoutName = (s = "KARTA" Or _
                        s = "KARTY" Or _
                        s = "PVCKARTA" Or _
                        s = "PVCKARTY")
End Function

Private Function GetWorkflowCardCount(ByVal workflowMode As Long, ByVal firstPage As Long, ByVal lastPage As Long) As Long
    Dim pageCount As Long

    pageCount = (lastPage - firstPage) + 1
    If pageCount <= 0 Then Exit Function

    Select Case workflowMode
        Case mki23SwCopies
            GetWorkflowCardCount = ParsePositiveLong(txtCopies.Text, 50)
        Case mki23SwCommonFrontFirst, mki23SwCommonBackLast
            If pageCount >= 2 Then GetWorkflowCardCount = pageCount - 1
        Case mki23SwAlternatingPairs
            GetWorkflowCardCount = pageCount \ 2
    End Select
End Function

Private Function GetSelectedSourceWorkflow() As Long
    If optAlternatingPairs.Value Then
        GetSelectedSourceWorkflow = mki23SwAlternatingPairs
    ElseIf optCommonBackLast.Value Then
        GetSelectedSourceWorkflow = mki23SwCommonBackLast
    ElseIf optAllPages.Value Then
        GetSelectedSourceWorkflow = mki23SwCommonFrontFirst
    Else
        GetSelectedSourceWorkflow = mki23SwCopies
    End If
End Function

Private Function GetSelectedSourceMode() As Long
    If GetSelectedSourceWorkflow() = mki23SwCopies Then
        GetSelectedSourceMode = mki21SmCopiesFromCurrentPage
    Else
        GetSelectedSourceMode = mki21SmAllPagesAsItems
    End If
End Function

Private Function GetSelectedOrientation() As Long
    If tglPortrait.Value Then
        GetSelectedOrientation = mki21IoPortrait
    Else
        GetSelectedOrientation = mki21IoLandscape
    End If
End Function

Private Function GetSelectedSideCode() As String
    If tglBack.Value Then
        GetSelectedSideCode = "Z"
    Else
        GetSelectedSideCode = "P"
    End If
End Function

Private Sub SetOrientationButtons(ByVal orientation As Long)
    mUpdating = True
    tglLandscape.Value = (orientation = mki21IoLandscape)
    tglPortrait.Value = (orientation = mki21IoPortrait)
    UpdateOrientationButtons
    mUpdating = False
End Sub

Private Sub SetSideButtons(ByVal sideCode As String)
    mUpdating = True
    tglFront.Value = (UCase$(sideCode) <> "Z")
    tglBack.Value = (UCase$(sideCode) = "Z")
    UpdateSideButtons
    mUpdating = False
End Sub

Private Sub SetSourceWorkflowButtons(ByVal workflowMode As Long)
    mUpdating = True
    optCopies.Value = (workflowMode = mki23SwCopies)
    optAllPages.Value = (workflowMode = mki23SwCommonFrontFirst)
    optCommonBackLast.Value = (workflowMode = mki23SwCommonBackLast)
    optAlternatingPairs.Value = (workflowMode = mki23SwAlternatingPairs)
    UpdateSourceModeButtons
    mUpdating = False
End Sub

Private Sub SetSourceModeButtons(ByVal sourceMode As Long)
    If sourceMode = mki21SmAllPagesAsItems Then
        If IsCardLayoutName(cboLayout.Value) Then
            SetSourceWorkflowButtons mki23SwAlternatingPairs
        Else
            SetSourceWorkflowButtons mki23SwCommonFrontFirst
        End If
    Else
        SetSourceWorkflowButtons mki23SwCopies
    End If
End Sub

Private Sub StyleToggleButton(ByVal btn As MSForms.ToggleButton, ByVal isSelected As Boolean, ByVal activeColor As Long)
    If isSelected Then
        btn.BackColor = activeColor
    Else
        btn.BackColor = COLOR_MUTED
    End If
    btn.ForeColor = COLOR_TEXT_DARK
End Sub

Private Sub UpdateSourceModeButtons()
    Dim isCardLayout As Boolean

    isCardLayout = IsCardLayoutName(cboLayout.Value)
    If isCardLayout Then
        optAllPages.Caption = "P ZZZ"
        optCommonBackLast.Visible = True
        optAlternatingPairs.Visible = True
        optCommonBackLast.Enabled = True
        optAlternatingPairs.Enabled = True
    Else
        optAllPages.Caption = "STRANY"
        optCommonBackLast.Value = False
        optAlternatingPairs.Value = False
        optCommonBackLast.Visible = False
        optAlternatingPairs.Visible = False
        optCommonBackLast.Enabled = False
        optAlternatingPairs.Enabled = False
        If Not optCopies.Value And Not optAllPages.Value Then optCopies.Value = True
    End If

    StyleToggleButton optCopies, optCopies.Value, COLOR_PRIMARY_SOFT
    StyleToggleButton optAllPages, optAllPages.Value, COLOR_SECONDARY_SOFT
    StyleToggleButton optCommonBackLast, optCommonBackLast.Value, COLOR_SECONDARY_SOFT
    StyleToggleButton optAlternatingPairs, optAlternatingPairs.Value, COLOR_SECONDARY_SOFT
    txtCopies.Enabled = optCopies.Value
    txtPageRange.Enabled = Not optCopies.Value
    tglFront.Enabled = optCopies.Value
    tglBack.Enabled = optCopies.Value
    If Not optCopies.Value Then
        tglFront.BackColor = COLOR_MUTED
        tglBack.BackColor = COLOR_MUTED
    End If
End Sub

Private Sub UpdateOrientationButtons()
    StyleToggleButton tglLandscape, tglLandscape.Value, COLOR_PRIMARY_SOFT
    StyleToggleButton tglPortrait, tglPortrait.Value, COLOR_PRIMARY_SOFT
End Sub

Private Sub UpdateSideButtons()
    StyleToggleButton tglFront, tglFront.Value, COLOR_PRIMARY_SOFT
    StyleToggleButton tglBack, tglBack.Value, COLOR_WARNING_SOFT
End Sub

Private Sub LoadLastSettings()
    Dim layoutName As String
    Dim sourceMode As Long
    Dim workflowMode As Long
    Dim inputOrientation As Long
    Dim sideCode As String

    layoutName = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Layout", "")
    If Len(layoutName) > 0 Then SelectComboValue cboLayout, layoutName
    If cboLayout.ListIndex < 0 And cboLayout.ListCount > 0 Then cboLayout.ListIndex = 0

    workflowMode = CLng(Val(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "SourceWorkflow", "-1")))
    If workflowMode < mki23SwCopies Or workflowMode > mki23SwAlternatingPairs Then
        sourceMode = CLng(Val(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "SourceMode", CStr(mki21SmCopiesFromCurrentPage))))
        If sourceMode = mki21SmAllPagesAsItems Then
            workflowMode = mki23SwAlternatingPairs
        Else
            workflowMode = mki23SwCopies
        End If
    End If
    SetSourceWorkflowButtons workflowMode

    txtCopies.Text = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Copies", "50")
    txtPageRange.Text = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "PageRange", "")

    inputOrientation = CLng(Val(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Orientation", CStr(mki21IoLandscape))))
    If inputOrientation <> mki21IoPortrait Then inputOrientation = mki21IoLandscape
    SetOrientationButtons inputOrientation

    sideCode = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Side", "P")
    SetSideButtons sideCode

    txtDocNumber.Text = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "JobId", "01")
    txtStartSlot.Text = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "StartSlot", "1")
    txtOutputFolder.Text = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "OutputFolder", GetDefaultOutputFolder())
    chkOpenExisting.Value = StringToBool(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "OpenExisting", "1"), True)
    LoadVectorSetting Trim$(cboLayout.Value)
End Sub

Private Sub ApplySourceDocumentAutoDefaults()
    Dim pageCount As Long
    Dim isCardLayout As Boolean

    If mSourceDoc Is Nothing Then Exit Sub
    isCardLayout = IsCardLayoutName(cboLayout.Value)

    On Error Resume Next
    pageCount = mSourceDoc.Pages.Count
    On Error GoTo 0

    If pageCount <= 1 Then
        SetSourceWorkflowButtons mki23SwCopies
        txtPageRange.Text = "1"
    Else
        If Not isCardLayout Then
            SetSourceWorkflowButtons mki23SwCommonFrontFirst
        ElseIf GetSelectedSourceWorkflow() = mki23SwCopies Then
            SetSourceWorkflowButtons mki23SwAlternatingPairs
        End If
        txtPageRange.Text = "1-" & CStr(pageCount)
    End If
End Sub

Private Sub SaveLastSettings()
    On Error Resume Next
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Layout", Trim$(cboLayout.Value)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "SourceMode", CStr(GetSelectedSourceMode())
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "SourceWorkflow", CStr(GetSelectedSourceWorkflow())
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Copies", Trim$(txtCopies.Text)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "PageRange", Trim$(txtPageRange.Text)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Orientation", CStr(GetSelectedOrientation())
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Side", GetSelectedSideCode()
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "JobId", Trim$(txtDocNumber.Text)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "StartSlot", Trim$(txtStartSlot.Text)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "OutputFolder", Trim$(txtOutputFolder.Text)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "OpenExisting", BoolToString(chkOpenExisting.Value)
    SaveVectorSetting Trim$(cboLayout.Value)
    On Error GoTo 0
End Sub

Private Function VectorSettingKey(ByVal layoutName As String) As String
    Dim keyName As String

    keyName = "Vector_" & Trim$(layoutName)
    keyName = Replace(keyName, "\", "_")
    keyName = Replace(keyName, "/", "_")
    keyName = Replace(keyName, ":", "_")
    If Len(keyName) > 255 Then keyName = Left$(keyName, 255)
    VectorSettingKey = keyName
End Function

Private Sub LoadVectorSetting(ByVal layoutName As String)
    If chkPreserveVector Is Nothing Then Exit Sub
    chkPreserveVector.Value = StringToBool(GetSetting(SETTINGS_APP, "LayoutOptions", VectorSettingKey(layoutName), "0"), False)
End Sub

Private Sub SaveVectorSetting(ByVal layoutName As String)
    If chkPreserveVector Is Nothing Or Len(Trim$(layoutName)) = 0 Then Exit Sub
    SaveSetting SETTINGS_APP, "LayoutOptions", VectorSettingKey(layoutName), BoolToString(chkPreserveVector.Value)
End Sub

Private Sub chkPreserveVector_Click()
    If mUpdating Then Exit Sub
    SaveVectorSetting Trim$(cboLayout.Value)
End Sub

Private Sub SelectComboValue(ByVal cbo As MSForms.ComboBox, ByVal wantedValue As String)
    Dim i As Long

    For i = 0 To cbo.ListCount - 1
        If StrComp(Trim$(CStr(cbo.List(i))), Trim$(wantedValue), vbTextCompare) = 0 Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Function BoolToString(ByVal value As Boolean) As String
    If value Then
        BoolToString = "1"
    Else
        BoolToString = "0"
    End If
End Function

Private Function StringToBool(ByVal value As String, ByVal defaultValue As Boolean) As Boolean
    value = LCase$(Trim$(value))
    If value = "1" Or value = "true" Or value = "ano" Or value = "ano" Or value = "yes" Then
        StringToBool = True
    ElseIf value = "0" Or value = "false" Or value = "nie" Or value = "no" Then
        StringToBool = False
    Else
        StringToBool = defaultValue
    End If
End Function

Private Function ParsePositiveLong(ByVal rawText As String, ByVal fallbackValue As Long) As Long
    ParsePositiveLong = CLng(Val(rawText))
    If ParsePositiveLong <= 0 Then ParsePositiveLong = fallbackValue
End Function

Private Sub ParsePageRange(ByVal rawText As String, ByRef firstPage As Long, ByRef lastPage As Long)
    Dim dashPos As Long

    rawText = Trim$(rawText)
    firstPage = 1
    lastPage = 0

    If Len(rawText) = 0 Then Exit Sub

    dashPos = InStr(1, rawText, "-", vbTextCompare)
    If dashPos > 0 Then
        firstPage = CLng(Val(Left$(rawText, dashPos - 1)))
        lastPage = CLng(Val(Mid$(rawText, dashPos + 1)))
    Else
        firstPage = CLng(Val(rawText))
        lastPage = firstPage
    End If

    If firstPage <= 0 Then firstPage = 1
End Sub

Private Sub NormalizePanelPageRange(ByVal pageCount As Long, ByRef firstPage As Long, ByRef lastPage As Long)
    If pageCount <= 0 Then
        firstPage = 1
        lastPage = 1
        Exit Sub
    End If

    If firstPage < 1 Then firstPage = 1
    If lastPage <= 0 Or lastPage > pageCount Then lastPage = pageCount
    If firstPage > pageCount Then firstPage = pageCount
    If lastPage < firstPage Then lastPage = firstPage
End Sub

Private Function PickBestSourceDocument() As Document
    Dim d As Document

    On Error Resume Next

    If Documents.Count = 0 Then Exit Function

    If Not ShouldSkipAsSourceDocument(ActiveDocument) Then
        Set PickBestSourceDocument = ActiveDocument
        Exit Function
    End If

    For Each d In Documents
        If Not ShouldSkipAsSourceDocument(d) Then
            Set PickBestSourceDocument = d
            Exit Function
        End If
    Next d
    On Error GoTo 0
End Function

Private Sub RefreshSourceDocumentFromActive()
    On Error Resume Next
    If Documents.Count = 0 Then Exit Sub
    If Not ShouldSkipAsSourceDocument(ActiveDocument) Then
        Set mSourceDoc = ActiveDocument
        CaptureCurrentSourceSelection
    End If
    On Error GoTo 0
End Sub

Private Sub CaptureCurrentSourceSelection()
    Dim sr As ShapeRange
    Dim sh As Shape
    Dim captured As Collection

    On Error Resume Next
    If mSourceDoc Is Nothing Then Exit Sub

    mSourceDoc.Activate
    Set sr = ActiveSelectionRange
    If Not sr Is Nothing Then
        If sr.Count > 0 Then
            Set captured = New Collection
            For Each sh In sr.Shapes
                If IsPanelSourceCandidateShape(sh) Then captured.Add sh
            Next sh
            If captured.Count > 0 Then Set mSourceSelectionShapes = captured
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub RestoreCachedSourceSelection()
    Dim sh As Shape
    Dim hasAny As Boolean

    On Error Resume Next
    If mSourceDoc Is Nothing Then Exit Sub
    If mSourceSelectionShapes Is Nothing Then Exit Sub

    mSourceDoc.Activate
    ActiveDocument.ClearSelection
    For Each sh In mSourceSelectionShapes
        If Not sh Is Nothing Then
            If IsPanelSourceCandidateShape(sh) Then
                If Not hasAny Then
                    sh.CreateSelection
                    hasAny = True
                Else
                    CallByName sh, "AddToSelection", VbMethod
                End If
            End If
        End If
    Next sh
    On Error GoTo 0
End Sub

Private Function IsPanelSourceCandidateShape(ByVal sh As Shape) As Boolean
    Dim layerName As String

    If sh Is Nothing Then Exit Function
    On Error Resume Next
    layerName = UCase$(Trim$(sh.Layer.Name))
    On Error GoTo 0
    If layerName = "MIMAKI_TEMP_RASTER" Then Exit Function
    If layerName = "MIMAKI_TEMP_IMPORT" Then Exit Function
    If Right$(layerName, 6) = "_SLOTS" Then Exit Function
    IsPanelSourceCandidateShape = True
End Function

Private Function ActiveDocumentHasPanelSourceSelection() As Boolean
    Dim sr As ShapeRange
    Dim sh As Shape

    On Error Resume Next
    If Documents.Count = 0 Then Exit Function
    If ShouldSkipAsSourceDocument(ActiveDocument) Then Exit Function
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then Exit Function
    For Each sh In sr.Shapes
        If IsPanelSourceCandidateShape(sh) Then
            ActiveDocumentHasPanelSourceSelection = True
            Exit Function
        End If
    Next sh
    On Error GoTo 0
End Function

Private Function NormalizePanelJobId(ByVal rawValue As String) As String
    Dim s As String
    Dim badChars As Variant
    Dim i As Long

    s = Trim$(rawValue)
    badChars = Array("", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i
    NormalizePanelJobId = s
End Function

Private Function ShouldSkipAsSourceDocument(ByVal doc As Document) As Boolean
    Dim docText As String

    On Error GoTo SafeExit

    If doc Is Nothing Then
        ShouldSkipAsSourceDocument = True
        Exit Function
    End If

    docText = UCase$(Trim$(doc.Title) & " " & Trim$(doc.FullFileName))
    If InStr(1, docText, "MIMAKI_SABLONATLAC", vbTextCompare) > 0 Then
        ShouldSkipAsSourceDocument = True
        Exit Function
    End If
    If InStr(1, docText, "_PVCKARTA_", vbTextCompare) > 0 Then
        ShouldSkipAsSourceDocument = True
        Exit Function
    End If

SafeExit:
End Function

Private Function GetDefaultOutputFolder() As String
    Dim fullName As String

    On Error Resume Next
    If Not mSourceDoc Is Nothing Then
        fullName = Trim$(mSourceDoc.FullFileName)
        If Len(fullName) > 0 Then
            GetDefaultOutputFolder = Left$(fullName, InStrRev(fullName, "\") - 1)
            Exit Function
        End If
    End If
    On Error GoTo 0

    GetDefaultOutputFolder = MimakiV21_GetDefaultOutputRoot()
End Function

Private Function BrowseForOutputFolder(ByVal initialFolder As String) As String
    Dim shellApp As Object
    Dim folderObj As Object

    On Error GoTo EH

    Set shellApp = CreateObject("Shell.Application")
    Set folderObj = shellApp.BrowseForFolder(0, "Vyber vystupny folder", 0, initialFolder)
    If Not folderObj Is Nothing Then
        BrowseForOutputFolder = folderObj.Self.Path
    End If
    Exit Function

EH:
    BrowseForOutputFolder = ""
End Function

Private Sub SetStatus(ByVal text As String)
    lblStatus.Caption = text
End Sub

Private Sub cboLayout_Change()
    RefreshPanelForLayoutChange
End Sub

Private Sub cmdRefresh_Click()
    mUpdating = True
    LoadLayouts
    mUpdating = False
    RefreshPanelForLayoutChange
End Sub

Private Sub cmdOutputFolder_Click()
    Dim selectedFolder As String

    selectedFolder = BrowseForOutputFolder(Trim$(txtOutputFolder.Text))
    If Len(Trim$(selectedFolder)) > 0 Then
        txtOutputFolder.Text = selectedFolder
        SaveLastSettings
    End If
End Sub

Private Sub optCopies_Click()
    If mUpdating Then Exit Sub
    SetSourceWorkflowButtons mki23SwCopies
    ApplyFullSheetCopiesDefault
    UpdateLayoutInfo
End Sub

Private Sub optAllPages_Click()
    If mUpdating Then Exit Sub
    SetSourceWorkflowButtons mki23SwCommonFrontFirst
    UpdateLayoutInfo
End Sub

Private Sub optCommonBackLast_Click()
    If mUpdating Then Exit Sub
    SetSourceWorkflowButtons mki23SwCommonBackLast
    UpdateLayoutInfo
End Sub

Private Sub optAlternatingPairs_Click()
    If mUpdating Then Exit Sub
    SetSourceWorkflowButtons mki23SwAlternatingPairs
    UpdateLayoutInfo
End Sub

Private Sub txtCopies_Change()
    UpdateLayoutInfo
End Sub

Private Sub txtStartSlot_Change()
    UpdateLayoutInfo
End Sub

Private Sub txtPageRange_Change()
    UpdateLayoutInfo
End Sub

Private Sub tglLandscape_Click()
    If mUpdating Then Exit Sub
    SetOrientationButtons mki21IoLandscape
End Sub

Private Sub tglPortrait_Click()
    If mUpdating Then Exit Sub
    SetOrientationButtons mki21IoPortrait
End Sub

Private Sub tglFront_Click()
    If mUpdating Then Exit Sub
    SetSideButtons "P"
End Sub

Private Sub tglBack_Click()
    If mUpdating Then Exit Sub
    SetSideButtons "Z"
End Sub

Private Sub cmdRun_Click()
    Dim layoutName As String
    Dim sideCode As String
    Dim jobId As String
    Dim startSlot As Long
    Dim sourceMode As Long
    Dim workflowMode As Long
    Dim copiesCount As Long
    Dim inputOrientation As Long
    Dim outDoc As Document
    Dim firstPage As Long
    Dim lastPage As Long
    Dim isCardLayout As Boolean

    On Error GoTo EH

    If Documents.Count = 0 Then
        MsgBox "Nie je otvoreny zdrojovy dokument.", vbExclamation, "Mimaki Imposer"
        Exit Sub
    End If

    layoutName = Trim$(cboLayout.Value)
    If Len(layoutName) = 0 Then
        MsgBox "Vyber produkt / layout.", vbExclamation, "Mimaki Imposer"
        Exit Sub
    End If
    isCardLayout = IsCardLayoutName(layoutName)

    workflowMode = GetSelectedSourceWorkflow()
    If Not isCardLayout Then
        If workflowMode <> mki23SwCopies And workflowMode <> mki23SwCommonFrontFirst Then workflowMode = mki23SwCopies
    End If
    sourceMode = GetSelectedSourceMode()
    If Not isCardLayout And workflowMode = mki23SwCommonFrontFirst Then sourceMode = mki21SmAllPagesAsItems
    copiesCount = ParsePositiveLong(txtCopies.Text, 50)
    jobId = NormalizePanelJobId(txtDocNumber.Text)
    If Len(jobId) = 0 Then
        MsgBox "Dopln cislo VP.", vbExclamation, "Mimaki Imposer"
        Exit Sub
    End If
    txtDocNumber.Text = jobId
    startSlot = ParsePositiveLong(txtStartSlot.Text, 1)
    sideCode = GetSelectedSideCode()
    inputOrientation = GetSelectedOrientation()
    preserveVector = chkPreserveVector.Value
    ParsePageRange Trim$(txtPageRange.Text), firstPage, lastPage

    If sourceMode = mki21SmCopiesFromCurrentPage Then
        If mSourceDoc Is Nothing Then
            If ActiveDocumentHasPanelSourceSelection() Then RefreshSourceDocumentFromActive
        ElseIf ActiveDocument Is mSourceDoc Then
            If ActiveDocumentHasPanelSourceSelection() Then CaptureCurrentSourceSelection
        End If
    ElseIf mSourceDoc Is Nothing Or ShouldSkipAsSourceDocument(mSourceDoc) Then
        RefreshSourceDocumentFromActive
    End If
    If mSourceDoc Is Nothing Or ShouldSkipAsSourceDocument(mSourceDoc) Then
        Set mSourceDoc = PickBestSourceDocument()
    End If
    If mSourceDoc Is Nothing Then
        MsgBox "Nepodarilo sa urcit zdrojovy dokument. Aktivuj povodny CDR s grafikou a otvor panel znova.", vbExclamation, "Mimaki Imposer"
        Exit Sub
    End If
    If sourceMode = mki21SmCopiesFromCurrentPage Then RestoreCachedSourceSelection

    SaveLastSettings
    SetStatus "Rozkladam..."
    If workflowMode = mki23SwCopies Or Not isCardLayout Then
        Set outDoc = CreateMimakiImpositionFromSourceDocByJobIdV21(mSourceDoc, layoutName, sideCode, jobId, startSlot, sourceMode, copiesCount, chkOpenExisting.Value, inputOrientation, Trim$(txtOutputFolder.Text), firstPage, lastPage, True, preserveVector)
    Else
        SetStatus "Rozkladam P/Z workflow..."
        Set outDoc = CreateMimakiPzWorkflowFromSourceDocByJobIdV23(mSourceDoc, layoutName, jobId, startSlot, workflowMode, chkOpenExisting.Value, inputOrientation, Trim$(txtOutputFolder.Text), firstPage, lastPage, True, preserveVector)
    End If
    If Not outDoc Is Nothing Then
        outDoc.Activate
        SetStatus "Hotovo: " & outDoc.Title
        Unload Me
    End If
    Exit Sub

EH:
    ShowFriendlyError "Rozlozenie", Err.Number, Err.Description
End Sub

Private Sub ShowFriendlyError(ByVal actionName As String, ByVal errNumber As Long, ByVal errDescription As String)
    Dim userMessage As String

    userMessage = BuildFriendlyErrorMessage(errNumber, errDescription)
    SetStatus FirstLine(userMessage)
    MsgBox userMessage & vbCrLf & vbCrLf & _
           "Detail pre servis: chyba " & CStr(errNumber) & vbCrLf & _
           "Popis: " & SafeErrorText(errDescription), _
           vbExclamation, "Mimaki Imposer - " & actionName
End Sub

Private Function SafeErrorText(ByVal errDescription As String) As String
    If Len(Trim$(errDescription)) = 0 Then
        SafeErrorText = "(bez popisu)"
    Else
        SafeErrorText = errDescription
    End If
End Function

Private Function BuildFriendlyErrorMessage(ByVal errNumber As Long, ByVal errDescription As String) As String
    Dim d As String

    Select Case errNumber
        Case vbObjectError + 400
            BuildFriendlyErrorMessage = "Nie je otvoreny zdrojovy dokument." & vbCrLf & vbCrLf & _
                                        "Otvor povodny CDR s grafikou a spusti Rozlozit znova."
            Exit Function
        Case vbObjectError + 401
            BuildFriendlyErrorMessage = "Vybrany produkt sa nenasiel v sablone." & vbCrLf & vbCrLf & _
                                        "V sablonovom CDR musi existovat vrstva s nazvom produktu a priponou _SLOTS, napriklad KARTA_SLOTS alebo PERO_SLOTS."
            Exit Function
        Case vbObjectError + 402
            BuildFriendlyErrorMessage = "Sablonova vrstva sa nasla, ale nepodarilo sa ju znovu otvorit." & vbCrLf & vbCrLf & _
                                        "Skus zavriet a znovu otvorit sablonovy CDR, potom obnov zoznam produktov."
            Exit Function
        Case vbObjectError + 500
            BuildFriendlyErrorMessage = "Nenasiel som ziadnu grafiku na rozlozenie." & vbCrLf & vbCrLf & _
                                        "Skontroluj, ci je otvoreny spravny zdrojovy CDR a ci zvoleny rezim zodpoveda tomu, co chces rozkladat."
            Exit Function
        Case vbObjectError + 501
            BuildFriendlyErrorMessage = "Nie je vybrata ziadna grafika na zdrojovej strane." & vbCrLf & vbCrLf & _
                                        "Rezim Kopie strany potrebuje oznaceny objekt alebo skupinu v povodnom CDR." & vbCrLf & _
                                        "Klikni do zdrojoveho dokumentu, oznac grafiku a spusti Rozlozit znova."
            Exit Function
        Case vbObjectError + 502
            BuildFriendlyErrorMessage = "Grafika sa skopirovala v tvare, ktory tento rozklad nevie spracovat." & vbCrLf & vbCrLf & _
                                        "Skus zdrojovu grafiku najprv zoskupit alebo zjednodusit a spusti rozklad znova."
            Exit Function
        Case vbObjectError + 503
            BuildFriendlyErrorMessage = "Zdrojovy vyber je prazdny." & vbCrLf & vbCrLf & _
                                        "Oznac grafiku v povodnom CDR a spusti Rozlozit znova."
            Exit Function
        Case vbObjectError + 504, vbObjectError + 505
            BuildFriendlyErrorMessage = "Docasna TIFF priprava Print Data zlyhala." & vbCrLf & vbCrLf & _
                                        "Skontroluj zdrojovu grafiku a white objekty pripravene ako samostatne krivky so spotom RDG_WHITE."
            Exit Function
        Case vbObjectError + 508, vbObjectError + 509, vbObjectError + 510, vbObjectError + 511, vbObjectError + 512, vbObjectError + 513, vbObjectError + 514, vbObjectError + 515
            BuildFriendlyErrorMessage = "Docasny temp export/import zlyhal." & vbCrLf & vbCrLf & _
                                        "Skontroluj zapis do W:\Temp\MimakiImposer a subor mimaki_debug.log. Detail chyby je zobrazeny nizsie."
            Exit Function
        Case vbObjectError + 520
            BuildFriendlyErrorMessage = "Vystupny CDR este nema nazov suboru." & vbCrLf & vbCrLf & _
                                        "Najprv spusti rozlozenie alebo uloz vystupny dokument."
            Exit Function
        Case vbObjectError + 530
            BuildFriendlyErrorMessage = "Vystupny priecinok je prazdny." & vbCrLf & vbCrLf & _
                                        "Vyber priecinok, kam sa ma ulozit produkcne PDF."
            Exit Function
        Case vbObjectError + 531, vbObjectError + 532
            BuildFriendlyErrorMessage = "RasterLink hotfolder nie je dostupny." & vbCrLf & vbCrLf & _
                                        "Skontroluj disk X:\ alebo sietove pripojenie."
            Exit Function
        Case vbObjectError + 533
            BuildFriendlyErrorMessage = "Produkcne PDF sa nevytvorilo alebo ho neviem najst." & vbCrLf & vbCrLf & _
                                        "Skontroluj export PDF a cielovy priecinok."
            Exit Function
        Case vbObjectError + 560, vbObjectError + 561
            BuildFriendlyErrorMessage = "Na zdrojovej strane sa nepodarilo vybrat objekty." & vbCrLf & vbCrLf & _
                                        "Skontroluj, ci stranka obsahuje odomknutu a viditelnu grafiku."
            Exit Function
        Case vbObjectError + 562, vbObjectError + 563
            BuildFriendlyErrorMessage = "Obsah zdrojovej strany sa nepodarilo pripravit ako jeden kus." & vbCrLf & vbCrLf & _
                                        "Skontroluj zamknute objekty, zamknute vrstvy alebo problemove objekty na danej strane."
            Exit Function
    End Select

    d = LCase$(Trim$(errDescription))

    If InStr(1, d, "no selected source shapes", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Nie je vybrata ziadna grafika na zdrojovej strane." & vbCrLf & vbCrLf & _
                                    "Rezim Kopie strany potrebuje oznaceny objekt alebo skupinu v povodnom CDR." & vbCrLf & _
                                    "Klikni do zdrojoveho dokumentu, oznac grafiku a spusti Rozlozit znova."
        Exit Function
    End If

    If InStr(1, d, "no source items found", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Nenasiel som ziadnu grafiku na rozlozenie." & vbCrLf & vbCrLf & _
                                    "Skontroluj, ci je otvoreny spravny zdrojovy CDR a ci zvoleny rezim zodpoveda tomu, co chces rozkladat."
        Exit Function
    End If

    If InStr(1, d, "layout", vbTextCompare) > 0 And InStr(1, d, "not found", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Vybrany produkt sa nenasiel v sablone." & vbCrLf & vbCrLf & _
                                    "V sablonovom CDR musi existovat vrstva s nazvom produktu a priponou _SLOTS, napriklad KARTA_SLOTS alebo PERO_SLOTS."
        Exit Function
    End If

    If InStr(1, d, "could not be reopened", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Sablonova vrstva sa nasla, ale nepodarilo sa ju znovu otvorit." & vbCrLf & vbCrLf & _
                                    "Skus zavriet a znovu otvorit sablonovy CDR, potom obnov zoznam produktov."
        Exit Function
    End If

    If InStr(1, d, "unsupported paste type", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Grafika sa skopirovala v tvare, ktory tento rozklad nevie spracovat." & vbCrLf & vbCrLf & _
                                    "Skus zdrojovu grafiku najprv zoskupit alebo zjednodusit a spusti rozklad znova."
        Exit Function
    End If

    If InStr(1, d, "tiff", vbTextCompare) > 0 Or InStr(1, d, "raster 300 dpi", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Docasna TIFF priprava Print Data zlyhala." & vbCrLf & vbCrLf & _
                                    "Skontroluj zdrojovu grafiku a white objekty pripravene ako samostatne krivky so spotom RDG_WHITE."
        Exit Function
    End If

    If InStr(1, d, "could not select source objects", vbTextCompare) > 0 Or _
       InStr(1, d, "source selection is empty", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Na zdrojovej strane sa nepodarilo vybrat objekty." & vbCrLf & vbCrLf & _
                                    "Skontroluj, ci stranka obsahuje odomknutu a viditelnu grafiku."
        Exit Function
    End If

    If InStr(1, d, "could not group source objects", vbTextCompare) > 0 Or _
       InStr(1, d, "could not be prepared as one group", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Obsah zdrojovej strany sa nepodarilo pripravit ako jeden kus." & vbCrLf & vbCrLf & _
                                    "Skontroluj zamknute objekty, zamknute vrstvy alebo problemove objekty na danej strane."
        Exit Function
    End If

    If InStr(1, d, "hotfolder is not available", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "RasterLink hotfolder nie je dostupny." & vbCrLf & vbCrLf & _
                                    "Skontroluj disk X:\ alebo sietove pripojenie. Lokalne subory ostanu ulozene."
        Exit Function
    End If

    If InStr(1, d, "output document has no cdr file name", vbTextCompare) > 0 Then
        BuildFriendlyErrorMessage = "Vystupny CDR este nema nazov suboru." & vbCrLf & vbCrLf & _
                                    "Najprv spusti rozlozenie alebo uloz vystupny dokument."
        Exit Function
    End If

    BuildFriendlyErrorMessage = "Nastala chyba v Mimaki Imposeri." & vbCrLf & vbCrLf & _
                                "Skontroluj aktivny dokument, vyber grafiky, zvoleny produkt a sablonu. Ak sa chyba zopakuje, posli cislo chyby servisne."
End Function

Private Function FirstLine(ByVal textValue As String) As String
    Dim p As Long

    p = InStr(1, textValue, vbCrLf, vbBinaryCompare)
    If p > 1 Then
        FirstLine = Left$(textValue, p - 1)
    Else
        FirstLine = textValue
    End If
End Function

Private Sub cmdClose_Click()
    SaveLastSettings
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveLastSettings
End Sub
