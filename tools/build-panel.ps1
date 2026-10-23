param(
    [string]$RepositoryRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Drawing
$RepositoryRoot = [IO.Path]::GetFullPath($RepositoryRoot)
$buildPath = Join-Path $RepositoryRoot ('local-only\build-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
[IO.Directory]::CreateDirectory($buildPath) | Out-Null
$corel = New-Object -ComObject CorelDRAW.Application.27
if (-not $corel.InitializeVBA()) { throw 'Corel could not initialize its installed VBA engine.' }
$projectName = 'MimakiPanelBuild'
foreach ($existing in $corel.GMSManager.Projects) {
    if ($existing.Name -eq $projectName) { throw 'Close the previous MimakiPanelBuild project before rebuilding.' }
}
$gms = $corel.GMSManager.Projects.Create($projectName, $false, (Join-Path $buildPath 'MimakiPanelBuild.gms'))
$project = $corel.VBE.VBProjects.Item($projectName)
$form = $null
$manifest = [Collections.Generic.List[object]]::new()

function Get-OleColor([string]$Hex) {
    return [Drawing.ColorTranslator]::ToOle([Drawing.ColorTranslator]::FromHtml($Hex))
}
$colors = @{
    Panel = Get-OleColor '#f3f4f1'
    Text = Get-OleColor '#222b24'
    Muted = Get-OleColor '#687167'
    Green = Get-OleColor '#276242'
    Info = Get-OleColor '#e9ede6'
    Line = Get-OleColor '#d6dad3'
    Border = Get-OleColor '#a4ada2'
    Button = Get-OleColor '#fafbf8'
    Selected = Get-OleColor '#e5eee5'
    White = Get-OleColor '#ffffff'
}

function Add-Control([string]$Type, [string]$Name, [single]$X, [single]$Y, [single]$W, [single]$H, [string]$Caption = '') {
    $control = $script:designer.Controls.Add("Forms.$Type.1", $Name, $true)
    $control.Left = $X
    $control.Top = $Y
    $control.Width = $W
    $control.Height = $H
    $control.FontName = 'Segoe UI'
    $control.FontSize = 9
    $control.ForeColor = $colors.Text
    if ($Type -ne 'TextBox' -and $Type -ne 'ComboBox') { $control.Caption = $Caption }
    switch ($Type) {
        'Label' {
            $control.BackStyle = 0
            $control.WordWrap = $false
        }
        { $_ -in 'TextBox', 'ComboBox' } {
            $control.BackColor = $colors.White
            $control.SpecialEffect = 0
            $control.BorderStyle = 1
            $control.BorderColor = $colors.Border
            if ($Type -eq 'ComboBox') { $control.Style = 2 }
        }
        { $_ -in 'CommandButton', 'ToggleButton' } {
            $control.BackColor = $colors.Button
            $control.TabStop = $true
            $control.PicturePosition = 1
        }
        'CheckBox' {
            $control.BackStyle = 0
            $control.TabStop = $true
        }
    }
    $manifest.Add([pscustomobject]@{name=$Name;type=$Type;left=$X;top=$Y;width=$W;height=$H})
    return ,$control
}

function Add-Label([string]$Name, [single]$X, [single]$Y, [single]$W, [single]$H, [string]$Caption, [bool]$Bold = $false) {
    $label = Add-Control 'Label' $Name $X $Y $W $H $Caption
    $label.FontBold = $Bold
    return ,$label
}

function Add-Rule([string]$Name, [single]$Y) {
    $rule = Add-Label $Name 12 $Y 384 0.75 ''
    $rule.BackStyle = 1
    $rule.BackColor = $colors.Line
}

function New-PanelIcon([string]$Name, [string]$Ink, [scriptblock]$Draw) {
    $bitmap = [Drawing.Bitmap]::new(16, 16)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([Drawing.Color]::Transparent)
    $graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $pen = [Drawing.Pen]::new([Drawing.ColorTranslator]::FromHtml($Ink), 1.25)
    $brush = [Drawing.SolidBrush]::new($pen.Color)
    & $Draw $graphics $pen $brush
    $icon = [Drawing.Icon]::FromHandle($bitmap.GetHicon())
    $iconPath = Join-Path $buildPath ($Name + '.ico')
    $stream = [IO.File]::Create($iconPath)
    try { $icon.Save($stream) } finally { $stream.Dispose(); $icon.Dispose(); $pen.Dispose(); $brush.Dispose(); $graphics.Dispose(); $bitmap.Dispose() }
    return $iconPath
}

try {
    foreach ($modulePath in @('Mimaki_Template_Imposer_v2_4.bas', 'modMimakiPlacementVector.bas')) {
        $module = $project.VBComponents.Add(1)
        $module.Name = [IO.Path]::GetFileNameWithoutExtension($modulePath)
        $module.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $RepositoryRoot "src\vba\$modulePath")))
    }
    $helper = $project.VBComponents.Add(1)
    $helper.Name = 'PanelBuildHelper'
    $helper.CodeModule.AddFromString(@'
Option Explicit
Public Sub SetDesignerPicture(ByVal FormName As String, ByVal ControlName As String, ByVal FilePath As String)
    Dim Designer As Object
    Set Designer = Application.VBE.VBProjects("MimakiPanelBuild").VBComponents(FormName).Designer
    Set Designer.Controls(ControlName).Picture = LoadPicture(FilePath)
End Sub
'@)

    $form = $project.VBComponents.Add(3)
    $form.Name = 'frmMimakiImposerPanel'
    $script:designer = $form.Designer
    $form.Properties.Item('Caption').Value = 'Mimaki Imposer v2.4.0 / 2026-10-02'
    $form.Properties.Item('Width').Value = 420
    $form.Properties.Item('Height').Value = 589
    $designer.BackColor = $colors.Panel
    $designer.Font.Name = 'Segoe UI'
    $designer.Font.Size = 9
    $form.Properties.Item('StartUpPosition').Value = 1

    $step = Add-Label 'lblStepTemplate' 12 10 15 14 '01' $true
    $step.ForeColor = $colors.Green
    $null = Add-Label 'lblTemplateHeading' 29 10 230 14 'Šablóna' $true
    $combo = Add-Control 'ComboBox' 'cboLayout' 12 31 316 22
    $combo.ControlTipText = 'Produkt podľa predlohovej vrstvy, ktorej názov končí na SLOT.'
    $null = Add-Control 'CommandButton' 'cmdRefresh' 334 31 62 22 'Refresh'

    $metrics = Add-Label 'lblMetricsBackground' 12 61 384 42 ''
    $metrics.BackStyle = 1
    $metrics.BackColor = $colors.Info
    $null = Add-Label 'lblCapacityHeading' 20 67 91 12 'Pozície / hárok'
    $null = Add-Label 'lblCapacity' 20 81 91 15 '-' $true
    $null = Add-Label 'lblSlotHeading' 127 67 104 12 'Rozmer slotu'
    $null = Add-Label 'lblSlotSize' 127 81 104 15 '-' $true
    $null = Add-Label 'lblGraphicHeading' 250 67 138 12 'Grafika + 3 mm spád'
    $null = Add-Label 'lblGraphicSize' 250 81 138 15 '-' $true
    foreach ($x in @(118, 241)) {
        $rule = Add-Label ('lblMetricRule' + $x) $x 68 0.75 28 ''
        $rule.BackStyle = 1
        $rule.BackColor = $colors.Line
    }
    Add-Rule 'lblTemplateRule' 115

    $step = Add-Label 'lblStepInput' 12 128 15 14 '02' $true
    $step.ForeColor = $colors.Green
    $null = Add-Label 'lblInputHeading' 29 128 230 14 'Vstupná grafika' $true
    $null = Add-Control 'ToggleButton' 'optCopies' 12 150 91.5 25 'Kópie strany'
    $null = Add-Control 'ToggleButton' 'optAllPages' 109.5 150 91.5 25 'P ZZZ'
    $null = Add-Control 'ToggleButton' 'optCommonBackLast' 207 150 91.5 25 'PPP Z'
    $null = Add-Control 'ToggleButton' 'optAlternatingPairs' 304.5 150 91.5 25 'PZ PZ'
    $helpLabel = Add-Label 'lblModeHelp' 12 183 384 20 'Opakovať vybranú grafiku z aktuálnej strany.'
    $helpLabel.ForeColor = $colors.Muted
    $helpLabel.FontSize = 8
    $helpLabel.WordWrap = $true

    $null = Add-Label 'lblCopies' 12 211 184 13 'Počet kusov'
    $rangeLabel = Add-Label 'lblPageRange' 12 211 184 13 'Rozsah strán'
    $rangeLabel.Visible = $false
    $copies = Add-Control 'TextBox' 'txtCopies' 12 228 184 22
    $copies.Value = '20'
    $range = Add-Control 'TextBox' 'txtPageRange' 12 228 184 22
    $range.Visible = $false
    $range.ControlTipText = 'Rozsah strán, napríklad 1-158 alebo 3-20.'
    $null = Add-Label 'lblOrientation' 208 211 188 13 'Orientácia vstupu'
    $null = Add-Control 'ToggleButton' 'tglLandscape' 208 228 91 22 'Na šírku'
    $null = Add-Control 'ToggleButton' 'tglPortrait' 305 228 91 22 'Na výšku'
    $null = Add-Control 'CheckBox' 'chkPreserveVector' 12 263 157 16 'Zachovať vektory'
    $processing = Add-Label 'lblProcessing' 175 265 221 13 'TIFF 600 dpi | WHITE vektorovo'
    $processing.ForeColor = $colors.Muted
    $processing.FontSize = 8
    $processing.TextAlign = 3
    Add-Rule 'lblInputRule' 291

    $step = Add-Label 'lblStepOutput' 12 304 15 14 '03' $true
    $step.ForeColor = $colors.Green
    $null = Add-Label 'lblOutputHeading' 29 304 230 14 'Výstup' $true
    $null = Add-Label 'lblDocNumber' 12 327 212 13 'Výrobný príkaz / VP'
    $null = Add-Label 'lblStartSlot' 236 327 70 13 'Začať od slotu'
    $null = Add-Label 'lblSide' 310 327 86 13 'Strana'
    $job = Add-Control 'TextBox' 'txtDocNumber' 12 344 212 22
    $job.Value = '01'
    $start = Add-Control 'TextBox' 'txtStartSlot' 236 344 62 22
    $start.Value = '1'
    $null = Add-Control 'ToggleButton' 'tglFront' 310 344 40 22 'P'
    $null = Add-Control 'ToggleButton' 'tglBack' 356 344 40 22 'Z'
    $existing = Add-Control 'CheckBox' 'chkOpenExisting' 12 378 225 16 'Použiť existujúci súbor'
    $existing.ControlTipText = 'Pokračovať v existujúcom výstupe. Vektorový režim vytvorí nový súbor.'
    $null = Add-Label 'lblOutputFolder' 12 404 330 13 'Predvolený priečinok'
    $null = Add-Control 'TextBox' 'txtOutputFolder' 12 421 351 22
    $null = Add-Control 'CommandButton' 'cmdOutputFolder' 369 421 27 22 '...'
    $folderHelp = Add-Label 'lblOutputHelp' 12 449 384 14 'Uložiť vedľa zdroja. Tento priečinok použiť ako náhradný.'
    $folderHelp.ForeColor = $colors.Muted
    $folderHelp.FontSize = 8
    Add-Rule 'lblOutputRule' 471

    $estimate = Add-Label 'lblLayoutInfo' 12 482 384 13 'Načítavam produkt...'
    $estimate.ForeColor = $colors.Green
    $estimate.FontSize = 8
    $status = Add-Label 'lblStatus' 12 497 384 14 'Pripravené na rozloženie.'
    $status.FontSize = 8
    $status.WordWrap = $true
    $run = Add-Control 'CommandButton' 'cmdRun' 12 520 298 28 'Rozložiť'
    $run.BackColor = $colors.Green
    $run.ForeColor = $colors.White
    $run.FontBold = $true
    $run.Default = $true
    $close = Add-Control 'CommandButton' 'cmdClose' 319 520 77 28 'Zavrieť'
    $close.Cancel = $true
    # Set the client size after measuring the host's non-client title and borders.
    $form.Properties.Item('Height').Value = $form.Properties.Item('Height').Value + (560 - $designer.InsideHeight)
    $form.Properties.Item('Width').Value = $form.Properties.Item('Width').Value + (408 - $designer.InsideWidth)

    $tabOrder = @('cboLayout','cmdRefresh','optCopies','optAllPages','optCommonBackLast','optAlternatingPairs','txtCopies','txtPageRange','tglLandscape','tglPortrait','chkPreserveVector','txtDocNumber','txtStartSlot','tglFront','tglBack','chkOpenExisting','txtOutputFolder','cmdOutputFolder','cmdRun','cmdClose')
    for ($i = 0; $i -lt $tabOrder.Count; $i++) { $designer.Controls.Item($tabOrder[$i]).TabIndex = $i }
    $icons = @{
        cmdRun = New-PanelIcon 'run' '#ffffff' { param($g,$p,$b) $g.FillPolygon($b, [Drawing.Point[]]@([Drawing.Point]::new(5,3),[Drawing.Point]::new(12,8),[Drawing.Point]::new(5,13))) }
        cmdOutputFolder = New-PanelIcon 'folder' '#687167' { param($g,$p,$b) $g.DrawLines($p,[Drawing.Point[]]@([Drawing.Point]::new(2,12),[Drawing.Point]::new(2,4),[Drawing.Point]::new(7,4),[Drawing.Point]::new(9,6),[Drawing.Point]::new(14,6),[Drawing.Point]::new(14,12),[Drawing.Point]::new(2,12))) }
        tglLandscape = New-PanelIcon 'landscape' '#276242' { param($g,$p,$b) $g.DrawRectangle($p,2,4,12,8) }
        tglPortrait = New-PanelIcon 'portrait' '#276242' { param($g,$p,$b) $g.DrawRectangle($p,4,2,8,12) }
    }
    $pictureCode = "Public Sub SetDesignerPictures()`r`nDim d As Object`r`nSet d = Application.VBE.VBProjects(""MimakiPanelBuild"").VBComponents(""frmMimakiImposerPanel"").Designer`r`n"
    foreach ($name in $icons.Keys) {
        $quotedPath = ([string]$icons[$name]).Replace('"', '""')
        $pictureCode += "Set d.Controls(""$name"").Picture = LoadPicture(""$quotedPath"")`r`n"
    }
    $pictureCode += 'End Sub'
    $helper.CodeModule.AddFromString($pictureCode)
    $null = $corel.GMSManager.RunMacro($projectName, 'PanelBuildHelper.SetDesignerPictures')
    $designer.Controls.Item('cmdOutputFolder').Caption = ''
    $designer.Controls.Item('cmdOutputFolder').PicturePosition = 12
    $designer.Controls.Item('cmdOutputFolder').ControlTipText = 'Vybrať predvolený priečinok.'

    $codePath = Join-Path $RepositoryRoot 'src\vba\frmMimakiImposerPanel_CODE_ONLY.txt'
    $form.CodeModule.AddFromString([IO.File]::ReadAllText($codePath))
    $exportPath = Join-Path $buildPath 'frmMimakiImposerPanel.frm'
    $form.Export($exportPath)
    # VBE recases shared COM property identifiers; keep the repository spelling.
    $formText = [IO.File]::ReadAllText($exportPath, [Text.Encoding]::Default)
    $formText = $formText -creplace '\.value\b', '.Value' -creplace '\.text\b', '.Text'
    $formText = ($formText -replace '(?m)[ \t]+\r?$', '').TrimEnd() + "`r`n"
    [IO.File]::WriteAllText($exportPath, $formText, [Text.UTF8Encoding]::new($false))
    $exportCode = $formText -replace '(?s)^.*?Attribute VB_Exposed = False\r?\n', ''
    [IO.File]::WriteAllText($codePath, $exportCode, [Text.UTF8Encoding]::new($false))
    foreach ($extension in @('frm', 'frx')) {
        Copy-Item -LiteralPath ([IO.Path]::ChangeExtension($exportPath, $extension)) -Destination (Join-Path $RepositoryRoot "src\vba\frmMimakiImposerPanel.$extension") -Force
    }
    $manifestPath = Join-Path $RepositoryRoot 'src\ui\designer-manifest.json'
    [IO.Directory]::CreateDirectory((Split-Path $manifestPath -Parent)) | Out-Null
    [IO.File]::WriteAllText($manifestPath, (ConvertTo-Json -InputObject @($manifest.ToArray()) -Depth 5), [Text.UTF8Encoding]::new($false))
    Write-Output ('Exported native designer form: ' + $manifest.Count + ' controls. Corel ' + $corel.VersionMajor + '.' + $corel.VersionMinor)
    Write-Output ('Client area: ' + $designer.InsideWidth + ' x ' + $designer.InsideHeight + ' pt')
    Write-Output ('Build project: ' + $gms.FullFileName)
} finally {
    # The scratch project is separate from the user's installed GlobalMacros project.
    $gms.Unload()
}
