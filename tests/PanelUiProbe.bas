Option Explicit

Public Function ValidatePanel() As String
    Dim f As Object
    Dim designer As Object
    Dim originalLeft As Single
    Dim expectedLeft As Single
    Dim vectorSetting As String
    Dim hasVectorSnapshot As Boolean
    Dim result As String
    Dim picture As Object

    On Error GoTo Failed
    Set designer = Application.VBE.VBProjects("MimakiPanelTest").VBComponents("frmMimakiImposerPanel").designer
    originalLeft = designer.Controls("txtDocNumber").Left
    expectedLeft = originalLeft + 2
    designer.Controls("txtDocNumber").Left = expectedLeft
    Set f = New frmMimakiImposerPanel
    Require f.Controls.Count = 50, "50 native controls"
    Require Abs(f.txtDocNumber.Left - expectedLeft) < 0.1, "designer movement survives initialization"
    Set picture = f.cmdRun.picture
    Require Not picture Is Nothing, "embedded run icon"
    Set picture = f.cmdOutputFolder.picture
    Require Not picture Is Nothing, "embedded folder icon"
    Set picture = f.tglLandscape.picture
    Require Not picture Is Nothing, "embedded orientation icon"

    f.cboLayout.AddItem "KARTA"
    f.cboLayout.AddItem "PANEL UI TEST"
    f.cboLayout.Value = "KARTA"
    Require f.optCommonBackLast.Visible And f.optAlternatingPairs.Visible, "card workflows visible"
    f.optCopies.Value = True
    Require f.txtCopies.Visible And Not f.txtPageRange.Visible, "copies field"
    f.optAllPages.Value = True
    Require f.txtPageRange.Visible And Not f.txtCopies.Visible, "page range field"
    Require Not f.tglFront.Enabled And Not f.tglBack.Enabled, "automatic duplex sides"
    f.cboLayout.Value = "PANEL UI TEST"
    Require Not f.optCommonBackLast.Visible And Not f.optAlternatingPairs.Visible, "single-sided product workflows"
    Require f.optAllPages.Caption = "Strany", "single-sided pages caption"
    f.optCopies.Value = True
    Require f.tglFront.Enabled And f.tglBack.Enabled, "copy side controls"
    f.tglPortrait.Value = True
    Require f.tglPortrait.Value And Not f.tglLandscape.Value, "orientation event"
    f.tglBack.Value = True
    Require f.tglBack.Value And Not f.tglFront.Value, "side event"

    vectorSetting = GetSetting("MimakiImposer", "LayoutOptions", "Vector_PANEL UI TEST", "{missing}")
    hasVectorSnapshot = True
    f.chkPreserveVector.Value = False
    f.chkPreserveVector.Value = True
    Require Not f.chkOpenExisting.Enabled, "vector existing-output restriction"
    Require InStr(f.lblProcessing.Caption, "bez TIFF") > 0, "vector processing label"
    f.chkPreserveVector.Value = False
    Require f.chkOpenExisting.Enabled, "TIFF existing-output option restored"
    Require InStr(f.lblProcessing.Caption, "600 dpi") > 0, "TIFF processing label"
    Require Abs(f.txtDocNumber.Left - expectedLeft) < 0.1, "designer movement survives mode changes"
    result = "PASS: native controls; embedded icons; editable geometry; card/single-sided workflows; input fields; orientation; side; vector/TIFF state."
    GoTo CleanUp

Failed:
    result = "FAIL: " & Err.Number & " / " & Err.Description
CleanUp:
    On Error Resume Next
    If Not f Is Nothing Then Unload f
    designer.Controls("txtDocNumber").Left = originalLeft
    If hasVectorSnapshot Then
        If vectorSetting = "{missing}" Then
            DeleteSetting "MimakiImposer", "LayoutOptions", "Vector_PANEL UI TEST"
        Else
            SaveSetting "MimakiImposer", "LayoutOptions", "Vector_PANEL UI TEST", vectorSetting
        End If
    End If
    ValidatePanel = result
End Function

Private Sub Require(ByVal Condition As Boolean, ByVal Message As String)
    If Not Condition Then Err.Raise vbObjectError + 2500, "PanelUiProbe", Message
End Sub

Public Sub ShowPreview()
    Load frmMimakiImposerPanel
    With frmMimakiImposerPanel
        .cboLayout.AddItem "PVC KARTY"
        .cboLayout.Value = "PVC KARTY"
        .optCopies.Value = True
        .txtCopies.Value = "20"
        .lblCapacity.Caption = "20"
        .lblSlotSize.Caption = "54 x 85 mm"
        .lblGraphicSize.Caption = "60 x 91 mm"
        .lblLayoutInfo.Caption = "20 kusov | 1 harok | Start: 1"
        .lblStatus.Caption = "Nahlad nativneho panela. Rozkladanie je tu vypnute."
        .cmdRun.Enabled = False
        .Show vbModeless
    End With
End Sub

Public Sub ClosePreview()
    Unload frmMimakiImposerPanel
End Sub
