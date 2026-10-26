param([string]$RepositoryRoot = (Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference = 'Stop'
$corel = New-Object -ComObject CorelDRAW.Application.27
$null = $corel.InitializeVBA()
$name = 'MimakiPanelTest'
$testPath = Join-Path $RepositoryRoot ('local-only\test-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
$settingsApp = 'MimakiPanelUITest_' + (Get-Date -Format 'yyyyMMddHHmmss')
[IO.Directory]::CreateDirectory($testPath) | Out-Null
$gms = $corel.GMSManager.Projects.Create($name, $false, (Join-Path $testPath ($name + '.gms')))
try {
    $project = $corel.VBE.VBProjects.Item($name)
    foreach ($file in @('Mimaki_Template_Imposer_v2_4.bas','modMimakiPlacementVector.bas')) {
        $module = $project.VBComponents.Add(1)
        $module.Name = [IO.Path]::GetFileNameWithoutExtension($file)
        $module.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $RepositoryRoot "src\vba\$file")))
    }
    $form = $project.VBComponents.Import((Join-Path $RepositoryRoot 'src\vba\frmMimakiImposerPanel.frm'))
    # QueryClose persists preferences even for an unloaded test instance.
    # Bind the scratch form and probe to a disposable settings namespace first.
    $testCode = $form.CodeModule.Lines(1, $form.CodeModule.CountOfLines)
    $testCode = $testCode.Replace('Private Const SETTINGS_APP As String = "MimakiImposer"', ('Private Const SETTINGS_APP As String = "' + $settingsApp + '"'))
    $form.CodeModule.DeleteLines(1, $form.CodeModule.CountOfLines)
    $form.CodeModule.AddFromString($testCode)
    $manifest = Get-Content -Raw (Join-Path $RepositoryRoot 'src\ui\designer-manifest.json') | ConvertFrom-Json
    if ($form.Designer.Controls.Count -ne $manifest.Count) { throw 'Native control count differs after import.' }
    foreach ($item in $manifest) {
        $control = $form.Designer.Controls.Item($item.name)
        foreach ($property in @('Left','Top','Width','Height')) {
            if ([Math]::Abs($control.$property - $item.$property) -gt 0.25) { throw "Designer geometry differs: $($item.name).$property" }
        }
    }
    $helper = $project.VBComponents.Add(1)
    $helper.Name = 'PanelUiProbe'
    $probeCode = [IO.File]::ReadAllText((Join-Path $RepositoryRoot 'tests\PanelUiProbe.bas'))
    $helper.CodeModule.AddFromString($probeCode.Replace('"MimakiImposer"', ('"' + $settingsApp + '"')))
    $result = $corel.GMSManager.RunMacro($name, 'PanelUiProbe.ValidatePanel')
    Write-Output $result
    if ([string]$result -notlike 'PASS:*') { throw 'The native VBA UI probe failed.' }
    $form.Export((Join-Path $testPath 'roundtrip.frm'))
    Write-Output 'PASS: imported FRM/FRX, 50 controls, all coordinates and dimensions round-trip.'
} finally {
    $gms.Unload()
    $settingsPath = 'HKCU:\Software\VB and VBA Program Settings\' + $settingsApp
    if (Test-Path -LiteralPath $settingsPath) { Remove-Item -LiteralPath $settingsPath -Recurse }
}
