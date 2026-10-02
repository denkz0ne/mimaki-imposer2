Option Explicit

' Mimaki template imposition macro for CorelDRAW 2026
' Version: v2.3.16
' Generated: 2026-08-12
' New model:
' - template slot layer name ends with SLOT, e.g. pvc SLOT, _pera cervene SLOT, AKRYL SLOT
' - only shapes with outline spot color CutContour1 on that layer are treated as slots
' - output document name: [ID]_[template]_[P/Z].cdr
' - production PDFs are always single-page: [ID]_[template]_[P/Z]_p01.pdf
' - output keeps template layer visible, non-printable, and locked
' - toolbar macros can run imposition and save production PDF/CDR
' - source artwork is copied through a slot-size crop canvas with bleed

Private Const DEFAULT_TEMPLATE_DOC_PATH As String = "S:\printstudio\zakazky\.VelFo\mimaki_sablonatlac.cdr"
Private Const MIMAKI_SCRIPT_VERSION As String = "v2.4.0"
Private Const MIMAKI_GENERATED_DATE As String = "2026-10-02"
Private Const PRINT_DATA_LAYER_NAME As String = "Print Data"
Private Const WHITE_LAYER_NAME As String = "WHITE"
Private Const WHITE_SPOT_NAME As String = "RDG_WHITE"
Private Const SLOT_OUTLINE_SPOT_NAME As String = "CutContour1"
Private Const CARD_GUIDE_SHAPE_NAME As String = "MIMAKI_CARD_BOUNDARY"
Private Const TEMP_RASTER_LAYER_NAME As String = "MIMAKI_TEMP_RASTER"
Private Const TEMP_IMPORT_LAYER_NAME As String = "MIMAKI_TEMP_IMPORT"
Private Const TEMP_WORK_ROOT_PATH As String = "W:\Temp"
Private Const TARGET_DOC_PREFIX As String = "pvckarta"
Private Const DEFAULT_LAYOUT_NAME As String = "KARTA"
Private Const SLOT_LAYER_SUFFIX As String = "SLOT"
Private Const DEFAULT_SIDE_CODE As String = "P"
Private Const DEFAULT_DOC_NUMBER As Long = 1
Private Const DEFAULT_JOB_ID As String = "01"
Private Const DEFAULT_START_SLOT As Long = 1
Private Const AUTO_ROTATE_TO_SLOT As Boolean = True
Private Const PRESERVE_SOURCE_PAGE_OFFSET As Boolean = True
Private Const CROP_SOURCE_TO_FIXED_CANVAS As Boolean = True
Private Const USE_TEMP_FILE_ENGINE As Boolean = True
Private Const FORCE_TEMP_FILE_ENGINE_BITMAP As Boolean = True
Private Const RASTERIZE_SOURCE_DPI As Long = 600
Private Const SOURCE_CARD_CANVAS_LONG_MM As Double = 90#
Private Const SOURCE_CARD_CANVAS_SHORT_MM As Double = 60#
Private Const SLOT_BLEED_MM As Double = 3#
Private Const CARD_SLOT_MATCH_TOLERANCE_MM As Double = 8#
Private Const CARD_PAGE_WIDTH_MM As Double = 85#
Private Const CARD_PAGE_HEIGHT_MM As Double = 54#
Private Const DEFAULT_RASTERLINK_HOTFOLDER_PATH As String = "X:\"
Private Const DEFAULT_OUTPUT_ROOT As String = "S:\printstudio\output"
Private Const DEFAULT_ICON_FOLDER_PATH As String = "W:\Dokumenty\New project\Mimaki_Imposer_v2_3\mimaki_icons\"
Private Const CONFIG_APP As String = "MimakiImposer"
Private Const CONFIG_SECTION As String = "ConfigV21"

Public Enum TSourceModeV21
    mki21SmCopiesFromCurrentPage = 1
    mki21SmAllPagesAsItems = 2
End Enum

Public Enum TMimakiSourceWorkflowV23
    mki23SwCopies = 0
    mki23SwCommonFrontFirst = 1
    mki23SwCommonBackLast = 2
    mki23SwAlternatingPairs = 3
End Enum

Public Enum TSideModeV21
    mki21SmSideFront = 1
    mki21SmSideBack = 2
End Enum

Public Enum TInputOrientationV21
    mki21IoAuto = 0
    mki21IoLandscape = 1
    mki21IoPortrait = 2
End Enum

Private Type TSlotInfo
    Name As String
    LayerName As String
    LeftX As Double
    RightX As Double
    TopY As Double
    BottomY As Double
    CenterX As Double
    CenterY As Double
End Type

Private Type TLayoutInfo
    PageWidth As Double
    PageHeight As Double
    SlotCount As Long
    SourcePageIndex As Long
    LayoutName As String
    SlotLayerName As String
End Type

Public gMimakiV21LastOutputDoc As Document

' ============================================================
' Public entry points
' ============================================================

Public Sub MimakiV21_OpenPanel()
    frmMimakiImposerPanel.Show
End Sub

Public Function MimakiV21_GetScriptVersionLabel() As String
    MimakiV21_GetScriptVersionLabel = "Mimaki Imposer " & MIMAKI_SCRIPT_VERSION & " / " & MIMAKI_GENERATED_DATE
End Function

Public Sub MimakiV21_RunImposer()
    MimakiV21_OpenPanel
End Sub

Public Sub MimakiV21_RunImposerPrompt()
    MimakiV21_RunPromptImposition
End Sub

Public Sub MimakiV21_CycleCardPageSetup()
    Dim srcDoc As Document
    Dim pg As Page
    Dim nextOrientation As Long
    Dim pageWidth As Double
    Dim pageHeight As Double

    On Error GoTo EH

    If Documents.Count = 0 Then
        MsgBox "Nie je otvoreny ziadny dokument.", vbExclamation, "Mimaki karta"
        Exit Sub
    End If

    Set srcDoc = ActiveDocument
    Set pg = srcDoc.ActivePage
    If pg Is Nothing Then Exit Sub

    NormalizeDocumentUnits srcDoc
    nextOrientation = GetNextCardSetupOrientation()
    If nextOrientation = mki21IoPortrait Then
        pageWidth = CARD_PAGE_HEIGHT_MM
        pageHeight = CARD_PAGE_WIDTH_MM
    Else
        pageWidth = CARD_PAGE_WIDTH_MM
        pageHeight = CARD_PAGE_HEIGHT_MM
    End If

    pg.SetSize pageWidth, pageHeight
    EnsureCardBoundaryGuide pg
    SaveSetting CONFIG_APP, CONFIG_SECTION, "Orientation", CStr(nextOrientation)
    SaveSetting CONFIG_APP, CONFIG_SECTION, "CardSetupLastApplied", CStr(nextOrientation)

    Exit Sub

EH:
    MsgBox FormatMimakiV21ErrorForUser(Err.Number, Err.Description), vbCritical, "Mimaki karta"
End Sub

Public Sub MimakiV21_OpenTemplateDocument()
    Dim tplDoc As Document
    Dim templatePath As String

    On Error GoTo EH

    templatePath = MimakiV21_GetTemplateDocumentPath()
    Set tplDoc = GetOpenDocumentByFullName(templatePath)
    If tplDoc Is Nothing Then
        Set tplDoc = Application.OpenDocument(templatePath)
    End If

    If Not tplDoc Is Nothing Then tplDoc.Activate
    Exit Sub

EH:
    MsgBox "Sablonu sa nepodarilo otvorit:" & vbCrLf & templatePath & vbCrLf & vbCrLf & _
           FormatMimakiV21ErrorForUser(Err.Number, Err.Description), vbCritical, "Mimaki sablona"
End Sub

Public Sub MimakiV21_SaveProductionFiles()
    MimakiV21_SaveProductionFilesEx False
End Sub

Public Sub MimakiV21_SaveProductionFilesEx(Optional ByVal exportToHotfolder As Boolean = False)
    Dim outDoc As Document
    Dim pdfPath As String
    Dim cdrPath As String
    Dim hotfolderPdfPath As String
    Dim hotfolderPath As String

    On Error GoTo EH

    Set outDoc = ResolveMimakiOutputDocument()
    If outDoc Is Nothing Then
        MsgBox "Najprv otvor alebo vytvor vystupny Mimaki dokument.", vbExclamation, "Mimaki ulozenie"
        Exit Sub
    End If

    outDoc.Activate
    SaveMimakiProductionFiles outDoc, cdrPath, pdfPath
    If exportToHotfolder Then
        hotfolderPath = MimakiV21_GetRasterLinkHotfolderPath()
        CopyProductionPDFSetToFolder RemoveFileExtension(cdrPath) & ".pdf", outDoc.Pages.Count, hotfolderPath, hotfolderPdfPath
    End If
    Exit Sub

EH:
    MsgBox FormatMimakiV21ErrorForUser(Err.Number, Err.Description), vbCritical, "Mimaki ulozenie"
End Sub

Public Sub MimakiV21_ExportToRasterLinkHotfolder()
    Dim outDoc As Document
    Dim pdfPath As String
    Dim hotfolderPath As String

    On Error GoTo EH

    Set outDoc = ResolveMimakiOutputDocument()
    If outDoc Is Nothing Then
        MsgBox "Najprv otvor alebo vytvor vystupny Mimaki dokument.", vbExclamation, "Mimaki RasterLink"
        Exit Sub
    End If

    outDoc.Activate
    hotfolderPath = MimakiV21_GetRasterLinkHotfolderPath()
    ExportMimakiProductionPDFOnly outDoc, hotfolderPath, pdfPath
    Exit Sub

EH:
    MsgBox FormatMimakiV21ErrorForUser(Err.Number, Err.Description), vbCritical, "Mimaki RasterLink"
End Sub

Public Function MimakiV21_GetTemplateDocumentPath() As String
    MimakiV21_GetTemplateDocumentPath = DEFAULT_TEMPLATE_DOC_PATH
End Function

Public Function MimakiV21_GetRasterLinkHotfolderPath() As String
    MimakiV21_GetRasterLinkHotfolderPath = GetSetting(CONFIG_APP, CONFIG_SECTION, "HotfolderPath", DEFAULT_RASTERLINK_HOTFOLDER_PATH)
End Function

Public Function MimakiV21_GetDefaultOutputRoot() As String
    MimakiV21_GetDefaultOutputRoot = GetSetting(CONFIG_APP, CONFIG_SECTION, "DefaultOutputFolder", DEFAULT_OUTPUT_ROOT)
End Function

Public Function MimakiV21_GetIconFolderPath() As String
    MimakiV21_GetIconFolderPath = GetSetting(CONFIG_APP, CONFIG_SECTION, "IconFolderPath", DEFAULT_ICON_FOLDER_PATH)
End Function

Public Sub MimakiV21_RunPromptImposition()
    Dim srcDoc As Document
    Dim layoutName As String
    Dim sideCode As String
    Dim docNumber As Long
    Dim startSlot As Long
    Dim sourceMode As Long
    Dim copiesCount As Long
    Dim openExisting As Boolean
    Dim promptResult As Long
    Dim outDoc As Document

    On Error GoTo EH

    If Documents.Count = 0 Then
        MsgBox "Nie je otvoreny zdrojovy dokument.", vbExclamation, "Mimaki rozklad"
        Exit Sub
    End If

    Set srcDoc = ActiveDocument

    layoutName = PromptText("Enter layout name (layer name), e.g. KARTA.", DEFAULT_LAYOUT_NAME, "Layout")
    If Len(layoutName) = 0 Then Exit Sub
    layoutName = NormalizeLayoutName(layoutName)

    sideCode = PromptText("Enter side code: P = front, Z = back.", DEFAULT_SIDE_CODE, "Side")
    If Len(sideCode) = 0 Then Exit Sub
    sideCode = NormalizeSideCode(sideCode)

    startSlot = PromptLong("Start slot number (1 = first position).", DEFAULT_START_SLOT, "Start position")
    If startSlot <= 0 Then Exit Sub

    sourceMode = AskSourceMode()
    If sourceMode = 0 Then Exit Sub

    If sourceMode = mki21SmCopiesFromCurrentPage Then
        copiesCount = PromptLong("Number of copies.", 50, "Copies")
        If copiesCount <= 0 Then Exit Sub
    Else
        copiesCount = 0
    End If

    docNumber = PromptTargetNumber(DEFAULT_DOC_NUMBER)
    If docNumber <= 0 Then Exit Sub

    promptResult = PromptOpenExisting(docNumber, sideCode, srcDoc)
    If promptResult = -1 Then Exit Sub
    openExisting = (promptResult = 1)

    Set outDoc = CreateMimakiImpositionExV21(layoutName, sideCode, docNumber, startSlot, sourceMode, copiesCount, openExisting)
    If outDoc Is Nothing Then Exit Sub

    outDoc.Activate
    MsgBox "Vystupny CDR je pripraveny:" & vbCrLf & outDoc.FullFileName, vbInformation, "Mimaki rozklad"
    Exit Sub

EH:
    MsgBox FormatMimakiV21ErrorForUser(Err.Number, Err.Description), vbCritical, "Mimaki rozklad"
End Sub

Private Function FormatMimakiV21ErrorForUser(ByVal errNumber As Long, ByVal errDescription As String) As String
    FormatMimakiV21ErrorForUser = BuildMimakiV21FriendlyError(errNumber, errDescription) & vbCrLf & vbCrLf & _
                                  "Detail pre servis: chyba " & CStr(errNumber)
End Function

Private Function BuildMimakiV21FriendlyError(ByVal errNumber As Long, ByVal errDescription As String) As String
    Dim d As String
    Dim hotfolderPath As String

    hotfolderPath = MimakiV21_GetRasterLinkHotfolderPath()

    Select Case errNumber
        Case vbObjectError + 400
            BuildMimakiV21FriendlyError = "Nie je otvoreny zdrojovy dokument." & vbCrLf & vbCrLf & _
                                          "Otvor povodny CDR s grafikou a spusti akciu znova."
            Exit Function
        Case vbObjectError + 401
            BuildMimakiV21FriendlyError = "Vybrany produkt sa nenasiel v sablone." & vbCrLf & vbCrLf & _
                                          "V sablonovom CDR musi existovat predlohova vrstva, ktorej nazov konci na SLOT, napriklad pvc SLOT alebo _pera cervene SLOT."
            Exit Function
        Case vbObjectError + 402
            BuildMimakiV21FriendlyError = "Sablonova vrstva sa nasla, ale nepodarilo sa ju znovu otvorit." & vbCrLf & vbCrLf & _
                                          "Skus zavriet a znovu otvorit sablonovy CDR, potom obnov zoznam produktov."
            Exit Function
        Case vbObjectError + 500
            BuildMimakiV21FriendlyError = "Nenasiel som ziadnu grafiku na rozlozenie." & vbCrLf & vbCrLf & _
                                          "Skontroluj zdrojovy CDR a zvoleny rezim vstupu."
            Exit Function
        Case vbObjectError + 501
            BuildMimakiV21FriendlyError = "Nie je vybrata ziadna grafika na zdrojovej strane." & vbCrLf & vbCrLf & _
                                          "Rezim Kopie strany potrebuje oznaceny objekt alebo skupinu v povodnom CDR. Klikni do zdrojoveho dokumentu, oznac grafiku a spusti akciu znova."
            Exit Function
        Case vbObjectError + 502
            BuildMimakiV21FriendlyError = "Grafika sa skopirovala v tvare, ktory tento rozklad nevie spracovat." & vbCrLf & vbCrLf & _
                                          "Skus zdrojovu grafiku najprv zoskupit alebo zjednodusit."
            Exit Function
        Case vbObjectError + 503
            BuildMimakiV21FriendlyError = "Zdrojovy vyber je prazdny." & vbCrLf & vbCrLf & _
                                          "Oznac grafiku v povodnom CDR a spusti akciu znova."
            Exit Function
        Case vbObjectError + 504, vbObjectError + 505
            BuildMimakiV21FriendlyError = "Docasna TIFF priprava Print Data zlyhala." & vbCrLf & vbCrLf & _
                                          "Skontroluj zdrojovu grafiku a temp TIFF workflow. WHITE objekty musia ostat ako samostatne krivky so spot farbou RDG_WHITE."
            Exit Function
        Case vbObjectError + 508, vbObjectError + 509, vbObjectError + 510, vbObjectError + 511, vbObjectError + 512, vbObjectError + 513, vbObjectError + 514, vbObjectError + 515
            BuildMimakiV21FriendlyError = "Docasny temp export/import zlyhal." & vbCrLf & vbCrLf & _
                                          "Skontroluj, ci Corel vie exportovat/importovat TIF a ci je zapisovatelny priecinok W:\Temp\MimakiImposer. Toto je stabilizacny rezim bez clipboardu."
            Exit Function
        Case vbObjectError + 570, vbObjectError + 571, vbObjectError + 572, vbObjectError + 573
            BuildMimakiV21FriendlyError = "P/Z workflow potrebuje iny rozsah stran." & vbCrLf & vbCrLf & _
                                          "Skontroluj zvoleny rezim P ZZZ / PPP Z / PZ PZ a rozsah stran. Detail: " & errDescription
            Exit Function
        Case vbObjectError + 520
            BuildMimakiV21FriendlyError = "Vystupny CDR este nema nazov suboru." & vbCrLf & vbCrLf & _
                                          "Najprv spusti rozlozenie alebo uloz vystupny dokument."
            Exit Function
        Case vbObjectError + 530
            BuildMimakiV21FriendlyError = "Vystupny priecinok je prazdny." & vbCrLf & vbCrLf & _
                                          "Vyber priecinok, kam sa ma ulozit produkcne PDF."
            Exit Function
        Case vbObjectError + 531, vbObjectError + 532
            BuildMimakiV21FriendlyError = "RasterLink hotfolder nie je dostupny." & vbCrLf & vbCrLf & _
                                          "Skontroluj cestu " & hotfolderPath & " alebo sietove pripojenie."
            Exit Function
        Case vbObjectError + 533
            BuildMimakiV21FriendlyError = "Produkcne PDF sa nevytvorilo alebo ho neviem najst." & vbCrLf & vbCrLf & _
                                          "Skontroluj export PDF a cielovy priecinok."
            Exit Function
        Case vbObjectError + 560, vbObjectError + 561
            BuildMimakiV21FriendlyError = "Na zdrojovej strane sa nepodarilo vybrat objekty." & vbCrLf & vbCrLf & _
                                          "Skontroluj, ci stranka obsahuje odomknutu a viditelnu grafiku."
            Exit Function
        Case vbObjectError + 562, vbObjectError + 563
            BuildMimakiV21FriendlyError = "Obsah zdrojovej strany sa nepodarilo pripravit ako jeden kus." & vbCrLf & vbCrLf & _
                                          "Skontroluj zamknute objekty, zamknute vrstvy alebo problemove objekty na danej strane."
            Exit Function
    End Select

    d = LCase$(Trim$(errDescription))
    If InStr(1, d, "tiff", vbTextCompare) > 0 Or InStr(1, d, "raster 300 dpi", vbTextCompare) > 0 Then
        BuildMimakiV21FriendlyError = "Docasna TIFF priprava Print Data zlyhala." & vbCrLf & vbCrLf & _
                                      "Skontroluj zdrojovu grafiku a temp TIFF workflow."
        Exit Function
    End If

    If InStr(1, d, "hotfolder", vbTextCompare) > 0 Then
        BuildMimakiV21FriendlyError = "RasterLink hotfolder nie je dostupny." & vbCrLf & vbCrLf & _
                                      "Skontroluj cestu " & hotfolderPath & " alebo sietove pripojenie."
        Exit Function
    End If

    BuildMimakiV21FriendlyError = "Nastala chyba v Mimaki Imposeri." & vbCrLf & vbCrLf & _
                                  "Skontroluj aktivny dokument, vyber grafiky, zvoleny produkt a sablonu. Ak sa chyba zopakuje, posli cislo chyby servisne."
End Function

Public Function CreateMimakiImpositionV21(ByVal layoutName As String, ByVal sourceMode As Long, ByVal copiesCount As Long) As Document
    Set CreateMimakiImpositionV21 = CreateMimakiImpositionExV21(layoutName, DEFAULT_SIDE_CODE, DEFAULT_DOC_NUMBER, DEFAULT_START_SLOT, sourceMode, copiesCount, False)
End Function

Public Function CreateMimakiImpositionExV21(ByVal layoutName As String, ByVal sideCode As String, ByVal docNumber As Long, ByVal startSlot As Long, ByVal sourceMode As Long, ByVal copiesCount As Long, ByVal openExisting As Boolean, Optional ByVal inputOrientation As Long = mki21IoAuto, Optional ByVal outputFolder As String = "", Optional ByVal firstSourcePage As Long = 1, Optional ByVal lastSourcePage As Long = 0, Optional ByVal rasterizeSource300 As Boolean = False, Optional ByVal preserveVector As Boolean = False) As Document
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 400, "CreateMimakiImpositionExV21", "No source document is open."
    End If

    Set CreateMimakiImpositionExV21 = CreateMimakiImpositionFromSourceDocByJobIdV21(ActiveDocument, layoutName, sideCode, Format$(docNumber, "00"), startSlot, sourceMode, copiesCount, openExisting, inputOrientation, outputFolder, firstSourcePage, lastSourcePage, rasterizeSource300, preserveVector)
End Function

Public Function CreateMimakiImpositionFromSourceDocV21(ByVal sourceDoc As Document, ByVal layoutName As String, ByVal sideCode As String, ByVal docNumber As Long, ByVal startSlot As Long, ByVal sourceMode As Long, ByVal copiesCount As Long, ByVal openExisting As Boolean, Optional ByVal inputOrientation As Long = mki21IoAuto, Optional ByVal outputFolder As String = "", Optional ByVal firstSourcePage As Long = 1, Optional ByVal lastSourcePage As Long = 0, Optional ByVal rasterizeSource300 As Boolean = False, Optional ByVal preserveVector As Boolean = False) As Document
    Set CreateMimakiImpositionFromSourceDocV21 = CreateMimakiImpositionFromSourceDocByJobIdV21(sourceDoc, layoutName, sideCode, Format$(docNumber, "00"), startSlot, sourceMode, copiesCount, openExisting, inputOrientation, outputFolder, firstSourcePage, lastSourcePage, rasterizeSource300, preserveVector)
End Function

Public Function CreateMimakiImpositionFromSourceDocByJobIdV21(ByVal sourceDoc As Document, ByVal layoutName As String, ByVal sideCode As String, ByVal jobId As String, ByVal startSlot As Long, ByVal sourceMode As Long, ByVal copiesCount As Long, ByVal openExisting As Boolean, Optional ByVal inputOrientation As Long = mki21IoAuto, Optional ByVal outputFolder As String = "", Optional ByVal firstSourcePage As Long = 1, Optional ByVal lastSourcePage As Long = 0, Optional ByVal rasterizeSource300 As Boolean = False, Optional ByVal preserveVector As Boolean = False) As Document
    Dim srcDoc As Document
    Dim tplDoc As Document
    Dim outDoc As Document
    Dim layout As TLayoutInfo
    Dim slots() As TSlotInfo
    Dim tplPage As Page
    Dim tplLayer As Layer
    Dim targetPath As String
    Dim openedTemplateHere As Boolean
    Dim outExists As Boolean
    Dim errNum As Long
    Dim errDesc As String
    Dim commandGroupStarted As Boolean
    Dim optimizationEnabled As Boolean

    On Error GoTo EH

    If sourceDoc Is Nothing Then
        Err.Raise vbObjectError + 400, "CreateMimakiImpositionFromSourceDocByJobIdV21", "No source document is open."
    End If

    Set srcDoc = sourceDoc
    NormalizeDocumentUnits srcDoc

    Set tplDoc = GetOpenDocumentByFullName(MimakiV21_GetTemplateDocumentPath())
    If tplDoc Is Nothing Then
        Set tplDoc = Application.OpenDocument(MimakiV21_GetTemplateDocumentPath())
        openedTemplateHere = True
    End If
    NormalizeDocumentUnits tplDoc

    layout = ReadTemplateInfo(tplDoc, layoutName, slots)
    If layout.SlotCount = 0 Then
        Err.Raise vbObjectError + 401, "CreateMimakiImpositionFromSourceDocByJobIdV21", "Layout '" & layoutName & "' was not found in template."
    End If

    Set tplPage = tplDoc.Pages(layout.SourcePageIndex)
    Set tplLayer = FindTemplateLayerByName(tplDoc, layout.LayoutName)
    If tplLayer Is Nothing Then
        Err.Raise vbObjectError + 402, "CreateMimakiImpositionFromSourceDocByJobIdV21", "Layout layer '" & layout.LayoutName & "' was found during scan, but could not be reopened."
    End If

    targetPath = BuildTargetDocPathByJobId(srcDoc, layout.LayoutName, jobId, sideCode, outputFolder)
    outExists = FileExists(targetPath)

    If openExisting And outExists Then
        Set outDoc = GetOpenDocumentByFullName(targetPath)
        If outDoc Is Nothing Then
            Set outDoc = Application.OpenDocument(targetPath)
        End If
    Else
        If outExists And Not openExisting Then
            targetPath = BuildAvailableFilePath(targetPath)
        End If
        Set outDoc = CreateDocument
    End If

    NormalizeDocumentUnits outDoc
    PrepareOutputDocument outDoc, layout.PageWidth, layout.PageHeight

    outDoc.Activate
    SetCorelOptimization True
    optimizationEnabled = True
    If Not USE_TEMP_FILE_ENGINE Then
        ActiveDocument.BeginCommandGroup "Mimaki imposition v2.3"
        commandGroupStarted = True
    End If

    EnsureImpositionStructure tplPage, tplLayer, outDoc, layout
    If startSlot <= 1 Then ClearArtworkLayers outDoc
    BuildImposedDocument srcDoc, outDoc, tplPage, tplLayer, layout, slots, sourceMode, copiesCount, startSlot, inputOrientation, firstSourcePage, lastSourcePage, rasterizeSource300, preserveVector

    outDoc.Activate
    If commandGroupStarted Then
        ActiveDocument.EndCommandGroup
        commandGroupStarted = False
    End If
    SetCorelOptimization False
    optimizationEnabled = False

    SaveOutputDocumentNoDialog outDoc, targetPath

    Set gMimakiV21LastOutputDoc = outDoc
    Set CreateMimakiImpositionFromSourceDocByJobIdV21 = outDoc

CleanUp:
    On Error Resume Next
    If optimizationEnabled Then SetCorelOptimization False
    If commandGroupStarted Then
        If Not outDoc Is Nothing Then outDoc.Activate
        ActiveDocument.EndCommandGroup
    End If
    If openedTemplateHere Then CloseDocumentWithoutSaving tplDoc
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "CreateMimakiImpositionFromSourceDocByJobIdV21", errDesc
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Resume CleanUp
End Function

Public Function CreateMimakiPzWorkflowFromSourceDocByJobIdV23(ByVal sourceDoc As Document, ByVal layoutName As String, ByVal jobId As String, ByVal startSlot As Long, ByVal workflowMode As Long, ByVal openExisting As Boolean, Optional ByVal inputOrientation As Long = mki21IoAuto, Optional ByVal outputFolder As String = "", Optional ByVal firstSourcePage As Long = 1, Optional ByVal lastSourcePage As Long = 0, Optional ByVal rasterizeSource300 As Boolean = False, Optional ByVal preserveVector As Boolean = False) As Document
    Dim firstPage As Long
    Dim lastPage As Long
    Dim cardCount As Long
    Dim frontPages() As Long
    Dim backPages() As Long
    Dim frontDoc As Document
    Dim backDoc As Document

    If sourceDoc Is Nothing Then
        Err.Raise vbObjectError + 400, "CreateMimakiPzWorkflowFromSourceDocByJobIdV23", "No source document is open."
    End If

    firstPage = firstSourcePage
    lastPage = lastSourcePage
    NormalizeSourcePageRange sourceDoc, firstPage, lastPage
    cardCount = (lastPage - firstPage) + 1

    Select Case workflowMode
        Case mki23SwCommonFrontFirst
            If cardCount < 2 Then Err.Raise vbObjectError + 570, "CreateMimakiPzWorkflowFromSourceDocByJobIdV23", "P ZZZ workflow needs at least 2 source pages."
            frontPages = BuildRepeatedPageList(firstPage, cardCount - 1)
            backPages = BuildSequentialPageList(firstPage + 1, lastPage)
        Case mki23SwCommonBackLast
            If cardCount < 2 Then Err.Raise vbObjectError + 570, "CreateMimakiPzWorkflowFromSourceDocByJobIdV23", "PPP Z workflow needs at least 2 source pages."
            frontPages = BuildSequentialPageList(firstPage, lastPage - 1)
            backPages = BuildRepeatedPageList(lastPage, cardCount - 1)
        Case mki23SwAlternatingPairs
            If cardCount < 2 Then Err.Raise vbObjectError + 570, "CreateMimakiPzWorkflowFromSourceDocByJobIdV23", "PZ PZ workflow needs at least 2 source pages."
            If (cardCount Mod 2) <> 0 Then Err.Raise vbObjectError + 571, "CreateMimakiPzWorkflowFromSourceDocByJobIdV23", "PZ PZ workflow needs an even number of source pages."
            frontPages = BuildAlternatingPageList(firstPage, lastPage, True)
            backPages = BuildAlternatingPageList(firstPage, lastPage, False)
        Case Else
            Err.Raise vbObjectError + 572, "CreateMimakiPzWorkflowFromSourceDocByJobIdV23", "Unknown P/Z workflow mode."
    End Select

    Set frontDoc = CreateMimakiImpositionFromPageListByJobIdV23(sourceDoc, layoutName, "P", jobId, startSlot, frontPages, openExisting, inputOrientation, outputFolder, rasterizeSource300, preserveVector)
    Set backDoc = CreateMimakiImpositionFromPageListByJobIdV23(sourceDoc, layoutName, "Z", jobId, startSlot, backPages, openExisting, inputOrientation, outputFolder, rasterizeSource300, preserveVector)

    If Not frontDoc Is Nothing Then Set gMimakiV21LastOutputDoc = frontDoc
    If Not backDoc Is Nothing Then Set gMimakiV21LastOutputDoc = backDoc
    Set CreateMimakiPzWorkflowFromSourceDocByJobIdV23 = backDoc
End Function

Private Function CreateMimakiImpositionFromPageListByJobIdV23(ByVal sourceDoc As Document, ByVal layoutName As String, ByVal sideCode As String, ByVal jobId As String, ByVal startSlot As Long, ByRef sourcePages() As Long, ByVal openExisting As Boolean, Optional ByVal inputOrientation As Long = mki21IoAuto, Optional ByVal outputFolder As String = "", Optional ByVal rasterizeSource300 As Boolean = False, Optional ByVal preserveVector As Boolean = False) As Document
    Dim srcDoc As Document
    Dim tplDoc As Document
    Dim outDoc As Document
    Dim layout As TLayoutInfo
    Dim slots() As TSlotInfo
    Dim tplPage As Page
    Dim tplLayer As Layer
    Dim targetPath As String
    Dim openedTemplateHere As Boolean
    Dim outExists As Boolean
    Dim errNum As Long
    Dim errDesc As String
    Dim commandGroupStarted As Boolean
    Dim optimizationEnabled As Boolean

    On Error GoTo EH

    If sourceDoc Is Nothing Then Err.Raise vbObjectError + 400, "CreateMimakiImpositionFromPageListByJobIdV23", "No source document is open."
    If Not LongArrayHasItems(sourcePages) Then Err.Raise vbObjectError + 500, "CreateMimakiImpositionFromPageListByJobIdV23", "No source page list."

    Set srcDoc = sourceDoc
    NormalizeDocumentUnits srcDoc

    Set tplDoc = GetOpenDocumentByFullName(MimakiV21_GetTemplateDocumentPath())
    If tplDoc Is Nothing Then
        Set tplDoc = Application.OpenDocument(MimakiV21_GetTemplateDocumentPath())
        openedTemplateHere = True
    End If
    NormalizeDocumentUnits tplDoc

    layout = ReadTemplateInfo(tplDoc, layoutName, slots)
    If layout.SlotCount = 0 Then Err.Raise vbObjectError + 401, "CreateMimakiImpositionFromPageListByJobIdV23", "Layout '" & layoutName & "' was not found in template."

    Set tplPage = tplDoc.Pages(layout.SourcePageIndex)
    Set tplLayer = FindTemplateLayerByName(tplDoc, layout.LayoutName)
    If tplLayer Is Nothing Then Err.Raise vbObjectError + 402, "CreateMimakiImpositionFromPageListByJobIdV23", "Layout layer '" & layout.LayoutName & "' was found during scan, but could not be reopened."

    targetPath = BuildTargetDocPathByJobId(srcDoc, layout.LayoutName, jobId, sideCode, outputFolder)
    outExists = FileExists(targetPath)

    If openExisting And outExists Then
        Set outDoc = GetOpenDocumentByFullName(targetPath)
        If outDoc Is Nothing Then Set outDoc = Application.OpenDocument(targetPath)
    Else
        If outExists And Not openExisting Then targetPath = BuildAvailableFilePath(targetPath)
        Set outDoc = CreateDocument
    End If

    NormalizeDocumentUnits outDoc
    PrepareOutputDocument outDoc, layout.PageWidth, layout.PageHeight

    outDoc.Activate
    SetCorelOptimization True
    optimizationEnabled = True
    If Not USE_TEMP_FILE_ENGINE Then
        ActiveDocument.BeginCommandGroup "Mimaki imposition v2.3"
        commandGroupStarted = True
    End If

    EnsureImpositionStructure tplPage, tplLayer, outDoc, layout
    If startSlot <= 1 Then ClearArtworkLayers outDoc
    BuildImposedDocumentFromPageList srcDoc, outDoc, tplPage, tplLayer, layout, slots, sourcePages, startSlot, inputOrientation, rasterizeSource300, preserveVector

    outDoc.Activate
    If commandGroupStarted Then
        ActiveDocument.EndCommandGroup
        commandGroupStarted = False
    End If
    SetCorelOptimization False
    optimizationEnabled = False

    SaveOutputDocumentNoDialog outDoc, targetPath

    Set CreateMimakiImpositionFromPageListByJobIdV23 = outDoc

CleanUp:
    On Error Resume Next
    If optimizationEnabled Then SetCorelOptimization False
    If commandGroupStarted Then
        If Not outDoc Is Nothing Then outDoc.Activate
        ActiveDocument.EndCommandGroup
    End If
    If openedTemplateHere Then CloseDocumentWithoutSaving tplDoc
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "CreateMimakiImpositionFromPageListByJobIdV23", errDesc
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Resume CleanUp
End Function
Public Function GetCurrentMimakiOutputDocumentV21() As Document
    On Error Resume Next
    If Not gMimakiV21LastOutputDoc Is Nothing Then
        Set GetCurrentMimakiOutputDocumentV21 = gMimakiV21LastOutputDoc
    End If
End Function

Public Function GetMimakiTemplateLayoutNamesV21() As Variant
    Dim tplDoc As Document
    Dim openedTemplateHere As Boolean
    Dim pg As Page
    Dim ly As Layer
    Dim seen As Collection
    Dim result() As String
    Dim i As Long
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo EH

    Set seen = New Collection
    Set tplDoc = GetOpenDocumentByFullName(MimakiV21_GetTemplateDocumentPath())
    If tplDoc Is Nothing Then
        Set tplDoc = Application.OpenDocument(MimakiV21_GetTemplateDocumentPath())
        openedTemplateHere = True
    End If

    AddMasterLayoutNames tplDoc, seen

    For Each pg In tplDoc.Pages
        For Each ly In pg.Layers
            If LayerLooksLikeLayout(ly) Then
                On Error Resume Next
                seen.Add GetLayoutDisplayNameFromLayerName(ly.Name), UCase$(GetLayoutDisplayNameFromLayerName(ly.Name))
                On Error GoTo 0
            End If
        Next ly
    Next pg

    If seen.Count > 0 Then
        ReDim result(0 To seen.Count - 1)
        For i = 1 To seen.Count
            result(i - 1) = seen(i)
        Next i
        GetMimakiTemplateLayoutNamesV21 = result
    End If

CleanUp:
    On Error Resume Next
    If openedTemplateHere Then CloseDocumentWithoutSaving tplDoc
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "GetMimakiTemplateLayoutNamesV21", errDesc
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Resume CleanUp
End Function

Public Function GetMimakiLayoutSlotCountV21(ByVal layoutName As String) As Long
    Dim tplDoc As Document
    Dim openedTemplateHere As Boolean
    Dim layout As TLayoutInfo
    Dim slots() As TSlotInfo
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo EH

    Set tplDoc = GetOpenDocumentByFullName(MimakiV21_GetTemplateDocumentPath())
    If tplDoc Is Nothing Then
        Set tplDoc = Application.OpenDocument(MimakiV21_GetTemplateDocumentPath())
        openedTemplateHere = True
    End If

    NormalizeDocumentUnits tplDoc
    layout = ReadTemplateInfo(tplDoc, layoutName, slots)
    GetMimakiLayoutSlotCountV21 = layout.SlotCount

CleanUp:
    On Error Resume Next
    If openedTemplateHere Then CloseDocumentWithoutSaving tplDoc
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "GetMimakiLayoutSlotCountV21", errDesc
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Resume CleanUp
End Function

Public Function GetMimakiLayoutFirstSlotSizeV21(ByVal layoutName As String, ByRef slotWidth As Double, ByRef slotHeight As Double, ByRef graphicWidth As Double, ByRef graphicHeight As Double) As Boolean
    Dim tplDoc As Document
    Dim openedTemplateHere As Boolean
    Dim layout As TLayoutInfo
    Dim slots() As TSlotInfo
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo EH

    slotWidth = 0#
    slotHeight = 0#
    graphicWidth = 0#
    graphicHeight = 0#

    Set tplDoc = GetOpenDocumentByFullName(MimakiV21_GetTemplateDocumentPath())
    If tplDoc Is Nothing Then
        Set tplDoc = Application.OpenDocument(MimakiV21_GetTemplateDocumentPath())
        openedTemplateHere = True
    End If

    NormalizeDocumentUnits tplDoc
    layout = ReadTemplateInfo(tplDoc, layoutName, slots)
    If layout.SlotCount > 0 Then
        slotWidth = Abs(slots(1).RightX - slots(1).LeftX)
        slotHeight = Abs(slots(1).TopY - slots(1).BottomY)
        graphicWidth = slotWidth + (SLOT_BLEED_MM * 2#)
        graphicHeight = slotHeight + (SLOT_BLEED_MM * 2#)
        GetMimakiLayoutFirstSlotSizeV21 = True
    End If

CleanUp:
    On Error Resume Next
    If openedTemplateHere Then CloseDocumentWithoutSaving tplDoc
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "GetMimakiLayoutFirstSlotSizeV21", errDesc
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Resume CleanUp
End Function

Public Function EstimateMimakiOutputPagesV21(ByVal layoutName As String, ByVal sourceMode As Long, ByVal copiesCount As Long, ByVal startSlot As Long) As Long
    Dim capacity As Long
    Dim totalItems As Long

    capacity = GetMimakiLayoutSlotCountV21(layoutName)
    If capacity <= 0 Then Exit Function
    If startSlot < 1 Then startSlot = 1

    If sourceMode = mki21SmCopiesFromCurrentPage Then
        totalItems = copiesCount
    ElseIf Documents.Count > 0 Then
        totalItems = ActiveDocument.Pages.Count
    End If

    If totalItems <= 0 Then Exit Function
    EstimateMimakiOutputPagesV21 = CeilingDiv(startSlot + totalItems - 1, capacity)
End Function

Private Sub AddMasterLayoutNames(ByVal doc As Document, ByRef seen As Collection)
    Dim masterPage As Object
    Dim ly As Layer

    On Error Resume Next
    Set masterPage = CallByName(doc, "MasterPage", VbGet)
    On Error GoTo 0

    If masterPage Is Nothing Then Exit Sub

    For Each ly In masterPage.Layers
        If LayerLooksLikeLayout(ly) Then
            On Error Resume Next
            seen.Add GetLayoutDisplayNameFromLayerName(ly.Name), UCase$(GetLayoutDisplayNameFromLayerName(ly.Name))
            On Error GoTo 0
        End If
    Next ly
End Sub

Public Sub MimakiV21_ExportCurrentPDF()
    MimakiPublishCurrentProductionPDF
End Sub

Public Sub MimakiV21_ExportCurrentEPS()
    MimakiShowExportDialog cdrEPS, ".eps"
End Sub

Public Sub MimakiV21_DebugTemplateLayers()
    Dim tplDoc As Document
    Dim openedTemplateHere As Boolean
    Dim pg As Page
    Dim ly As Layer
    Dim masterPage As Object
    Dim msg As String

    On Error GoTo EH

    Set tplDoc = GetOpenDocumentByFullName(MimakiV21_GetTemplateDocumentPath())
    If tplDoc Is Nothing Then
        Set tplDoc = Application.OpenDocument(MimakiV21_GetTemplateDocumentPath())
        openedTemplateHere = True
    End If

    msg = "Template layers seen by macro:" & vbCrLf & vbCrLf

    For Each pg In tplDoc.Pages
        msg = msg & "Page " & CStr(pg.Index) & ":" & vbCrLf
        For Each ly In pg.Layers
            msg = msg & "  - " & ly.Name & vbCrLf
        Next ly
    Next pg

    On Error Resume Next
    Set masterPage = CallByName(tplDoc, "MasterPage", VbGet)
    On Error GoTo EH

    If Not masterPage Is Nothing Then
        msg = msg & vbCrLf & "MasterPage:" & vbCrLf
        For Each ly In masterPage.Layers
            msg = msg & "  - " & ly.Name & vbCrLf
        Next ly
    Else
        msg = msg & vbCrLf & "MasterPage not available through this Corel API call." & vbCrLf
    End If

    MsgBox msg, vbInformation, "Mimaki debug"

CleanUp:
    On Error Resume Next
    If openedTemplateHere Then CloseDocumentWithoutSaving tplDoc
    Exit Sub

EH:
    MsgBox "Debug failed: " & Err.Description, vbCritical, "Mimaki debug"
    Resume CleanUp
End Sub

' ============================================================
' Core build
' ============================================================

Private Sub BuildImposedDocument(ByVal srcDoc As Document, ByVal outDoc As Document, ByVal tplPage As Page, ByVal tplLayer As Layer, ByRef layout As TLayoutInfo, ByRef slots() As TSlotInfo, ByVal sourceMode As Long, ByVal copiesCount As Long, ByVal startSlot As Long, ByVal inputOrientation As Long, ByVal firstSourcePage As Long, ByVal lastSourcePage As Long, ByVal rasterizeSource300 As Boolean, ByVal preserveVector As Boolean)
    Dim totalItems As Long
    Dim lastSlot As Long
    Dim pagesNeeded As Long
    Dim itemIndex As Long
    Dim slotNumber As Long
    Dim outPageIndex As Long
    Dim slotIndex As Long
    Dim currentSlot As TSlotInfo
    Dim srcRange As ShapeRange
    Dim srcPage As Page
    Dim placed As Shape
    Dim placedWhite As Shape
    Dim pageItemCount As Long
    Dim pagePrototype() As Shape
    Dim pageWhitePrototype() As Shape
    Dim pagePrototypeSlot() As TSlotInfo
    Dim useFastCopyMode As Boolean

    If layout.SlotCount <= 0 Then Exit Sub
    If startSlot < 1 Then startSlot = 1
    NormalizeSourcePageRange srcDoc, firstSourcePage, lastSourcePage
    ClearTemporaryRasterLayersInDocument srcDoc, True

    If sourceMode = mki21SmCopiesFromCurrentPage Then
        srcDoc.ActivePage.Activate
        Set srcRange = GetSelectableRange(srcDoc, True)
        If srcRange Is Nothing Then
            Err.Raise vbObjectError + 501, "BuildImposedDocument", "No selected source shapes on the active source page."
        End If
        totalItems = copiesCount
    Else
        totalItems = CountSelectablePagesInRange(srcDoc, firstSourcePage, lastSourcePage)
    End If

    If totalItems <= 0 Then
        Err.Raise vbObjectError + 500, "BuildImposedDocument", "No source items found."
    End If

    lastSlot = startSlot + totalItems - 1
    pagesNeeded = CeilingDiv(lastSlot, layout.SlotCount)
    If pagesNeeded < 1 Then pagesNeeded = 1
    MimakiDebugLog "build summary layout=" & layout.LayoutName & " slotCount=" & CStr(layout.SlotCount) & " totalItems=" & CStr(totalItems) & " copies=" & CStr(copiesCount) & " startSlot=" & CStr(startSlot) & " pagesNeeded=" & CStr(pagesNeeded)

    EnsureDocumentPages outDoc, pagesNeeded, layout.PageWidth, layout.PageHeight
    EnsureImpositionStructure tplPage, tplLayer, outDoc, layout

    If sourceMode = mki21SmCopiesFromCurrentPage Then
        useFastCopyMode = SlotsHaveUniformOrientation(slots) And Not preserveVector
        If useFastCopyMode Then
            ReDim pagePrototype(1 To pagesNeeded)
            ReDim pageWhitePrototype(1 To pagesNeeded)
            ReDim pagePrototypeSlot(1 To pagesNeeded)
        End If

        For itemIndex = 1 To copiesCount
            slotNumber = startSlot + itemIndex - 1
            outPageIndex = ((slotNumber - 1) \ layout.SlotCount) + 1
            slotIndex = ((slotNumber - 1) Mod layout.SlotCount) + 1
            currentSlot = slots(slotIndex)
            If useFastCopyMode Then
                If pagePrototype(outPageIndex) Is Nothing And pageWhitePrototype(outPageIndex) Is Nothing Then
                    PlaceSourceContentIntoSlotSet srcDoc, outDoc, srcRange, outPageIndex, currentSlot, inputOrientation, placed, pageWhitePrototype(outPageIndex), preserveVector
                    Set pagePrototype(outPageIndex) = placed
                    pagePrototypeSlot(outPageIndex) = currentSlot
                Else
                    Set placed = DuplicatePlacedShapeIntoSlot(pagePrototype(outPageIndex), pagePrototypeSlot(outPageIndex), currentSlot)
                    DuplicatePlacedShapeIntoLayerSlot pageWhitePrototype(outPageIndex), pagePrototypeSlot(outPageIndex), currentSlot
                End If
            Else
                Set placedWhite = Nothing
                PlaceSourceContentIntoSlotSet srcDoc, outDoc, srcRange, outPageIndex, currentSlot, inputOrientation, placed, placedWhite, preserveVector
            End If
            If (itemIndex Mod 5) = 0 Then DoEvents
        Next itemIndex
    Else
        pageItemCount = 0
        For itemIndex = firstSourcePage To lastSourcePage
            Set srcPage = srcDoc.Pages(itemIndex)
            srcPage.Activate
            Set srcRange = GetPageSourceRange(srcDoc, srcPage, True)
            If Not srcRange Is Nothing Then
                pageItemCount = pageItemCount + 1
                slotNumber = startSlot + pageItemCount - 1
                outPageIndex = ((slotNumber - 1) \ layout.SlotCount) + 1
                slotIndex = ((slotNumber - 1) Mod layout.SlotCount) + 1
                currentSlot = slots(slotIndex)
                Set placedWhite = Nothing
                PlaceSourceContentIntoSlotSet srcDoc, outDoc, srcRange, outPageIndex, currentSlot, inputOrientation, placed, placedWhite, preserveVector
                If (pageItemCount Mod 5) = 0 Then DoEvents
            End If
        Next itemIndex
    End If
End Sub

Private Sub BuildImposedDocumentFromPageList(ByVal srcDoc As Document, ByVal outDoc As Document, ByVal tplPage As Page, ByVal tplLayer As Layer, ByRef layout As TLayoutInfo, ByRef slots() As TSlotInfo, ByRef sourcePages() As Long, ByVal startSlot As Long, ByVal inputOrientation As Long, ByVal rasterizeSource300 As Boolean, ByVal preserveVector As Boolean)
    Dim totalItems As Long
    Dim lastSlot As Long
    Dim pagesNeeded As Long
    Dim itemIndex As Long
    Dim pageNumber As Long
    Dim slotNumber As Long
    Dim outPageIndex As Long
    Dim slotIndex As Long
    Dim currentSlot As TSlotInfo
    Dim srcRange As ShapeRange
    Dim srcPage As Page
    Dim placed As Shape
    Dim placedWhite As Shape
    Dim pagePrototype() As Shape
    Dim pageWhitePrototype() As Shape
    Dim pagePrototypeSlot() As TSlotInfo
    Dim pagePrototypeSourcePage() As Long

    If layout.SlotCount <= 0 Then Exit Sub
    If startSlot < 1 Then startSlot = 1
    If Not LongArrayHasItems(sourcePages) Then
        Err.Raise vbObjectError + 500, "BuildImposedDocumentFromPageList", "No source pages in workflow."
    End If
    ClearTemporaryRasterLayersInDocument srcDoc, True

    totalItems = UBound(sourcePages) - LBound(sourcePages) + 1
    lastSlot = startSlot + totalItems - 1
    pagesNeeded = CeilingDiv(lastSlot, layout.SlotCount)
    If pagesNeeded < 1 Then pagesNeeded = 1
    MimakiDebugLog "build page-list summary layout=" & layout.LayoutName & " slotCount=" & CStr(layout.SlotCount) & " totalItems=" & CStr(totalItems) & " startSlot=" & CStr(startSlot) & " pagesNeeded=" & CStr(pagesNeeded)

    EnsureDocumentPages outDoc, pagesNeeded, layout.PageWidth, layout.PageHeight
    EnsureImpositionStructure tplPage, tplLayer, outDoc, layout
    ReDim pagePrototype(1 To pagesNeeded)
    ReDim pageWhitePrototype(1 To pagesNeeded)
    ReDim pagePrototypeSlot(1 To pagesNeeded)
    ReDim pagePrototypeSourcePage(1 To pagesNeeded)

    For itemIndex = LBound(sourcePages) To UBound(sourcePages)
        pageNumber = sourcePages(itemIndex)
        If pageNumber < 1 Or pageNumber > srcDoc.Pages.Count Then
            Err.Raise vbObjectError + 573, "BuildImposedDocumentFromPageList", "Source page is outside document range: " & CStr(pageNumber)
        End If

        Set srcPage = srcDoc.Pages(pageNumber)
        srcPage.Activate
        Set srcRange = GetPageSourceRange(srcDoc, srcPage, True)
        If srcRange Is Nothing Then
            Err.Raise vbObjectError + 500, "BuildImposedDocumentFromPageList", "No source graphics on page " & CStr(pageNumber) & "."
        End If

        slotNumber = startSlot + itemIndex - LBound(sourcePages)
        outPageIndex = ((slotNumber - 1) \ layout.SlotCount) + 1
        slotIndex = ((slotNumber - 1) Mod layout.SlotCount) + 1
        currentSlot = slots(slotIndex)
        If Not preserveVector And (Not pagePrototype(outPageIndex) Is Nothing Or Not pageWhitePrototype(outPageIndex) Is Nothing) And pagePrototypeSourcePage(outPageIndex) = pageNumber Then
            Set placed = DuplicatePlacedShapeIntoSlot(pagePrototype(outPageIndex), pagePrototypeSlot(outPageIndex), currentSlot)
            DuplicatePlacedShapeIntoLayerSlot pageWhitePrototype(outPageIndex), pagePrototypeSlot(outPageIndex), currentSlot
        Else
            Set placedWhite = Nothing
            PlaceSourceContentIntoSlotSet srcDoc, outDoc, srcRange, outPageIndex, currentSlot, inputOrientation, placed, placedWhite, preserveVector
            Set pagePrototype(outPageIndex) = placed
            Set pageWhitePrototype(outPageIndex) = placedWhite
            pagePrototypeSlot(outPageIndex) = currentSlot
            pagePrototypeSourcePage(outPageIndex) = pageNumber
        End If
        If ((itemIndex - LBound(sourcePages) + 1) Mod 5) = 0 Then DoEvents
    Next itemIndex
End Sub

Private Function PlaceSourceRangeIntoSlot(ByVal srcDoc As Document, ByVal outDoc As Document, ByVal srcRange As ShapeRange, ByVal outPageIndex As Long, ByRef slot As TSlotInfo, ByVal inputOrientation As Long, Optional ByVal preserveSourcePagePosition As Boolean = False, Optional ByVal rasterizeAfterPaste As Boolean = False) As Shape
    Dim outPage As Page
    Dim artLayer As Layer
    Dim pastedObj As Object
    Dim pastedShape As Shape
    Dim normalizedOffsetX As Double
    Dim normalizedOffsetY As Double
    Dim appliedRotation As Double
    Dim targetX As Double
    Dim targetY As Double
    Dim canvasWidth As Double
    Dim canvasHeight As Double
    Dim effectiveInputOrientation As Long

    If USE_TEMP_FILE_ENGINE Then
        Set PlaceSourceRangeIntoSlot = PlaceSourceRangeIntoSlotViaTempFile(srcDoc, outDoc, srcRange, outPageIndex, slot, inputOrientation, rasterizeAfterPaste)
        Exit Function
    End If

    srcDoc.Activate
    effectiveInputOrientation = ResolveEffectiveInputOrientation(inputOrientation, srcRange)
    If rasterizeAfterPaste Or CROP_SOURCE_TO_FIXED_CANVAS Then
        GetSourceCanvasSizeForSlot slot, effectiveInputOrientation, canvasWidth, canvasHeight
    ElseIf preserveSourcePagePosition Then
        GetSourceRangeNormalizedOffset srcDoc.ActivePage, srcRange, normalizedOffsetX, normalizedOffsetY
    End If
    CopySourceRangeForX8 srcRange, rasterizeAfterPaste, srcDoc.ActivePage, canvasWidth, canvasHeight, CROP_SOURCE_TO_FIXED_CANVAS

    EnsureDocumentPages outDoc, outPageIndex, outDoc.Pages(1).SizeWidth, outDoc.Pages(1).SizeHeight
    Set outPage = outDoc.Pages(outPageIndex)
    Set artLayer = EnsureArtworkLayer(outPage)
    outDoc.Activate
    outPage.Activate
    artLayer.Activate

    Set pastedObj = artLayer.Paste
    If pastedObj Is Nothing Then
        Exit Function
    End If

    Set pastedShape = NormalizePastedObjectToSingleShape(pastedObj)
    If pastedShape Is Nothing Then
        Err.Raise vbObjectError + 502, "PlaceSourceRangeIntoSlot", "Unsupported paste type: " & TypeName(pastedObj)
    End If

    If AUTO_ROTATE_TO_SLOT Then
        appliedRotation = RotateShapeToSlotOrientation(pastedShape, slot, effectiveInputOrientation)
        If preserveSourcePagePosition Then RotateOffset normalizedOffsetX, normalizedOffsetY, appliedRotation
    End If

    If rasterizeAfterPaste And Not IsBitmapShape(pastedShape) Then
        Set pastedShape = RasterizePlacedShapeForProduction(pastedShape)
    End If

    If rasterizeAfterPaste Or CROP_SOURCE_TO_FIXED_CANVAS Then
        CenterShapeOnPoint pastedShape, slot.CenterX, slot.CenterY
    ElseIf preserveSourcePagePosition Then
        targetX = slot.CenterX + (normalizedOffsetX * Abs(slot.RightX - slot.LeftX))
        targetY = slot.CenterY + (normalizedOffsetY * Abs(slot.TopY - slot.BottomY))
        CenterShapeOnPoint pastedShape, targetX, targetY
    Else
        CenterShapeOnPoint pastedShape, slot.CenterX, slot.CenterY
    End If

    Set PlaceSourceRangeIntoSlot = pastedShape
End Function

Private Sub PlaceSourceContentIntoSlotSet(ByVal srcDoc As Document, ByVal outDoc As Document, ByVal srcRange As ShapeRange, ByVal outPageIndex As Long, ByRef slot As TSlotInfo, ByVal inputOrientation As Long, ByRef placedPrint As Shape, ByRef placedWhite As Shape, Optional ByVal preserveVector As Boolean = False)
    Dim printRange As ShapeRange
    Dim whiteRange As ShapeRange

    Set placedPrint = Nothing
    Set placedWhite = Nothing
    If srcRange Is Nothing Then Exit Sub

    Set printRange = BuildFilteredSelectionRange(srcDoc, srcRange, False)
    Set whiteRange = BuildFilteredSelectionRange(srcDoc, srcRange, True)

    If Not printRange Is Nothing Then
        If preserveVector Then
            Set placedPrint = PlaceVectorPrintRangeIntoSlot(srcDoc, outDoc, printRange, outPageIndex, slot, inputOrientation)
        Else
            Set placedPrint = PlaceSourceRangeIntoSlot(srcDoc, outDoc, printRange, outPageIndex, slot, inputOrientation, PRESERVE_SOURCE_PAGE_OFFSET, True)
        End If
    End If
    If Not whiteRange Is Nothing Then
        Set placedWhite = PlaceWhiteRangeIntoSlot(srcDoc, outDoc, whiteRange, outPageIndex, slot, inputOrientation)
    End If

    If placedPrint Is Nothing And placedWhite Is Nothing Then
        Err.Raise vbObjectError + 500, "PlaceSourceContentIntoSlotSet", "No source shapes remained after Print Data / WHITE split."
    End If
End Sub

Private Function PlaceVectorPrintRangeIntoSlot(ByVal srcDoc As Document, ByVal outDoc As Document, ByVal srcRange As ShapeRange, ByVal outPageIndex As Long, ByRef slot As TSlotInfo, ByVal inputOrientation As Long) As Shape
    Dim sourceLayerNames As Collection
    Dim layerItem As Variant
    Dim sh As Shape
    Dim layerRange As ShapeRange
    Dim tempShape As Shape
    Dim pastedShape As Shape
    Dim pastedObject As Object
    Dim outPage As Page
    Dim outLayer As Layer
    Dim slotLayer As Layer
    Dim candidateLayer As Layer
    Dim sourceLayerName As String
    Dim outputLayerName As String
    Dim canvasWidth As Double
    Dim canvasHeight As Double
    Dim cropCenterX As Double
    Dim cropCenterY As Double
    Dim effectiveOrientation As Long
    Dim appliedRotation As Double
    Dim hasPlacedShape As Boolean
    Dim localErr As Long
    Dim localDesc As String

    If srcRange Is Nothing Then Exit Function
    If srcRange.Count <= 0 Then Exit Function
    On Error GoTo EH

    srcDoc.Activate
    effectiveOrientation = ResolveEffectiveInputOrientation(inputOrientation, srcRange)
    GetSourceCanvasSizeForSlot slot, effectiveOrientation, canvasWidth, canvasHeight
    GetSourceCropCenter srcDoc.ActivePage, srcRange, canvasWidth, canvasHeight, cropCenterX, cropCenterY
    Set sourceLayerNames = New Collection

    For Each sh In srcRange.Shapes
        If IsSourceCandidateShape(sh) Then
            sourceLayerName = Trim$(sh.Layer.Name)
            If Len(sourceLayerName) > 0 Then
                On Error Resume Next
                sourceLayerNames.Add sourceLayerName, UCase$(sourceLayerName)
                Err.Clear
                On Error GoTo EH
            End If
        End If
    Next sh

    If sourceLayerNames.Count = 0 Then
        Err.Raise vbObjectError + 580, "PlaceVectorPrintRangeIntoSlot", "No eligible source layers were found."
    End If

    For Each layerItem In sourceLayerNames
        sourceLayerName = CStr(layerItem)
        Set layerRange = BuildSourceLayerRange(srcDoc, srcRange, sourceLayerName)
        If layerRange Is Nothing Then
            Err.Raise vbObjectError + 581, "PlaceVectorPrintRangeIntoSlot", "Could not capture source layer: " & sourceLayerName
        End If
        If layerRange.Count = 0 Then
            Err.Raise vbObjectError + 581, "PlaceVectorPrintRangeIntoSlot", "Could not capture source layer: " & sourceLayerName
        End If

        Set tempShape = CreateTemporarySourceCanvasCopy(layerRange, srcDoc.ActivePage, canvasWidth, canvasHeight, False, True, cropCenterX, cropCenterY)
        If tempShape Is Nothing Then
            Err.Raise vbObjectError + 582, "PlaceVectorPrintRangeIntoSlot", "Could not create the vector crop canvas for layer: " & sourceLayerName
        End If

        tempShape.Copy
        tempShape.Delete
        Set tempShape = Nothing

        EnsureDocumentPages outDoc, outPageIndex, outDoc.Pages(1).SizeWidth, outDoc.Pages(1).SizeHeight
        Set outPage = outDoc.Pages(outPageIndex)
        outputLayerName = sourceLayerName
        If StrComp(outputLayerName, WHITE_LAYER_NAME, vbTextCompare) = 0 Then outputLayerName = "WHITE - Artwork"
        Set outLayer = EnsureLayer(outPage, outputLayerName)
        outLayer.Visible = True
        outLayer.Printable = True
        outLayer.Editable = True

        Set slotLayer = Nothing
        For Each candidateLayer In outPage.Layers
            If IsSlotsLayerName(candidateLayer.Name) Then
                Set slotLayer = candidateLayer
                Exit For
            End If
        Next candidateLayer
        If Not slotLayer Is Nothing Then MoveLayerAboveReference slotLayer, outLayer

        outDoc.Activate
        outPage.Activate
        outLayer.Activate
        Set pastedObject = outLayer.Paste
        If pastedObject Is Nothing Then
            Err.Raise vbObjectError + 583, "PlaceVectorPrintRangeIntoSlot", "Could not paste vector canvas for layer: " & sourceLayerName
        End If
        Set pastedShape = NormalizePastedObjectToSingleShape(pastedObject)
        If pastedShape Is Nothing Then
            Err.Raise vbObjectError + 584, "PlaceVectorPrintRangeIntoSlot", "Unsupported vector paste result for layer: " & sourceLayerName
        End If

        If AUTO_ROTATE_TO_SLOT Then appliedRotation = RotateShapeToSlotOrientation(pastedShape, slot, effectiveOrientation)
        CenterShapeOnPoint pastedShape, slot.CenterX, slot.CenterY
        If Not hasPlacedShape Then
            Set PlaceVectorPrintRangeIntoSlot = pastedShape
            hasPlacedShape = True
        End If
        Set pastedShape = Nothing
        Set pastedObject = Nothing
        Set layerRange = Nothing
        srcDoc.Activate
    Next layerItem

    ClearTemporaryRasterLayer srcDoc.ActivePage, True
    srcDoc.ClearSelection
    If Not hasPlacedShape Then
        Err.Raise vbObjectError + 585, "PlaceVectorPrintRangeIntoSlot", "No vector artwork was placed."
    End If
    Exit Function

EH:
    localErr = Err.Number
    localDesc = Err.Description
    On Error Resume Next
    If Not tempShape Is Nothing Then tempShape.Delete
    ClearTemporaryRasterLayer srcDoc.ActivePage, True
    srcDoc.ClearSelection
    On Error GoTo 0
    If localErr = 0 Then localErr = vbObjectError + 586
    Err.Raise localErr, "PlaceVectorPrintRangeIntoSlot", localDesc
End Function

Private Function BuildSourceLayerRange(ByVal srcDoc As Document, ByVal srcRange As ShapeRange, ByVal sourceLayerName As String) As ShapeRange
    Dim sh As Shape
    Dim hasAny As Boolean

    If srcDoc Is Nothing Or srcRange Is Nothing Then Exit Function
    srcDoc.Activate
    srcDoc.ClearSelection
    For Each sh In srcRange.Shapes
        If IsSourceCandidateShape(sh) Then
            If StrComp(Trim$(sh.Layer.Name), sourceLayerName, vbTextCompare) = 0 Then
                If Not hasAny Then
                    sh.CreateSelection
                    hasAny = True
                Else
                    CallByName sh, "AddToSelection", VbMethod
                End If
            End If
        End If
    Next sh
    If hasAny Then Set BuildSourceLayerRange = ActiveSelectionRange
End Function

Private Function PlaceWhiteRangeIntoSlot(ByVal srcDoc As Document, ByVal outDoc As Document, ByVal srcRange As ShapeRange, ByVal outPageIndex As Long, ByRef slot As TSlotInfo, ByVal inputOrientation As Long) As Shape
    Dim outPage As Page
    Dim whiteLayer As Layer
    Dim pastedObj As Object
    Dim pastedShape As Shape
    Dim normalizedOffsetX As Double
    Dim normalizedOffsetY As Double
    Dim appliedRotation As Double
    Dim targetX As Double
    Dim targetY As Double
    Dim effectiveInputOrientation As Long

    If srcRange Is Nothing Then Exit Function
    If srcRange.Count <= 0 Then Exit Function

    srcDoc.Activate
    effectiveInputOrientation = ResolveEffectiveInputOrientation(inputOrientation, srcRange)
    GetSourceRangeNormalizedOffset srcDoc.ActivePage, srcRange, normalizedOffsetX, normalizedOffsetY
    CopySourceRangeForX8 srcRange, False, srcDoc.ActivePage, 0#, 0#, False

    EnsureDocumentPages outDoc, outPageIndex, outDoc.Pages(1).SizeWidth, outDoc.Pages(1).SizeHeight
    Set outPage = outDoc.Pages(outPageIndex)
    Set whiteLayer = EnsureWhiteLayer(outPage)
    outDoc.Activate
    outPage.Activate
    whiteLayer.Activate

    Set pastedObj = whiteLayer.Paste
    If pastedObj Is Nothing Then Exit Function

    Set pastedShape = NormalizePastedObjectToSingleShape(pastedObj)
    If pastedShape Is Nothing Then
        Err.Raise vbObjectError + 502, "PlaceWhiteRangeIntoSlot", "Unsupported WHITE paste type: " & TypeName(pastedObj)
    End If

    If AUTO_ROTATE_TO_SLOT Then
        appliedRotation = RotateShapeToSlotOrientation(pastedShape, slot, effectiveInputOrientation)
        RotateOffset normalizedOffsetX, normalizedOffsetY, appliedRotation
    End If

    targetX = slot.CenterX + (normalizedOffsetX * Abs(slot.RightX - slot.LeftX))
    targetY = slot.CenterY + (normalizedOffsetY * Abs(slot.TopY - slot.BottomY))
    CenterShapeOnPoint pastedShape, targetX, targetY
    Set PlaceWhiteRangeIntoSlot = pastedShape
End Function

Private Function PlaceSourceRangeIntoSlotViaTempFile(ByVal srcDoc As Document, ByVal outDoc As Document, ByVal srcRange As ShapeRange, ByVal outPageIndex As Long, ByRef slot As TSlotInfo, ByVal inputOrientation As Long, ByVal rasterizeToBitmap As Boolean) As Shape
    Dim outPage As Page
    Dim artLayer As Layer
    Dim tempShape As Shape
    Dim importedShape As Shape
    Dim tempPath As String
    Dim canvasWidth As Double
    Dim canvasHeight As Double
    Dim appliedRotation As Double
    Dim exportAsBitmap As Boolean
    Dim tempStage As String
    Dim errNum As Long
    Dim errDesc As String
    Dim effectiveInputOrientation As Long
    Dim sourceWorkPage As Page

    On Error GoTo EH

    srcDoc.Activate
    Set sourceWorkPage = srcDoc.ActivePage
    effectiveInputOrientation = ResolveEffectiveInputOrientation(inputOrientation, srcRange)
    GetSourceCanvasSizeForSlot slot, effectiveInputOrientation, canvasWidth, canvasHeight
    exportAsBitmap = (rasterizeToBitmap Or FORCE_TEMP_FILE_ENGINE_BITMAP)
    MimakiDebugLog "slot begin page=" & CStr(outPageIndex) & " slot=" & slot.Name & " tempBitmap=" & CStr(exportAsBitmap) & " inputOri=" & CStr(inputOrientation) & " effectiveOri=" & CStr(effectiveInputOrientation)
    MimakiDebugLog "slot geometry mm " & Format$(Abs(slot.RightX - slot.LeftX), "0.###") & " x " & Format$(Abs(slot.TopY - slot.BottomY), "0.###") & " canvas with bleed mm " & Format$(canvasWidth, "0.###") & " x " & Format$(canvasHeight, "0.###")

    tempPath = BuildMimakiTempAssetPath(exportAsBitmap)
    MimakiDebugLog "temp path " & tempPath

    tempStage = "create temp crop bitmap"
    ClearTemporaryRasterLayer sourceWorkPage, True
    Set tempShape = CreateTemporarySourceCanvasCopy(srcRange, sourceWorkPage, canvasWidth, canvasHeight, exportAsBitmap)
    If tempShape Is Nothing Then
        Err.Raise vbObjectError + 509, "PlaceSourceRangeIntoSlotViaTempFile", "Temporary source shape was not created."
    End If
    MimakiDebugLog "temp shape size mm " & Format$(Abs(tempShape.RightX - tempShape.LeftX), "0.###") & " x " & Format$(Abs(tempShape.TopY - tempShape.BottomY), "0.###")

    tempStage = "export selected temp bitmap"
    MimakiDebugLog tempStage
    ExportTemporaryShapeToFile srcDoc, tempShape, tempPath, exportAsBitmap

    On Error Resume Next
    tempShape.Delete
    Set tempShape = Nothing
    ClearTemporaryRasterLayer sourceWorkPage, True
    On Error GoTo EH

    EnsureDocumentPages outDoc, outPageIndex, outDoc.Pages(1).SizeWidth, outDoc.Pages(1).SizeHeight
    Set outPage = outDoc.Pages(outPageIndex)
    Set artLayer = EnsureArtworkLayer(outPage)

    outDoc.Activate
    outPage.Activate
    artLayer.Activate

    tempStage = "import temp asset"
    MimakiDebugLog tempStage
    Set importedShape = ImportTemporaryShapeFile(outPage, artLayer, tempPath)
    If importedShape Is Nothing Then
        Err.Raise vbObjectError + 508, "PlaceSourceRangeIntoSlotViaTempFile", "Temporary source asset could not be imported."
    End If
    MimakiDebugLog "imported temp size mm before normalize " & Format$(Abs(importedShape.RightX - importedShape.LeftX), "0.###") & " x " & Format$(Abs(importedShape.TopY - importedShape.BottomY), "0.###")
    NormalizeImportedCanvasSize importedShape, canvasWidth, canvasHeight
    MimakiDebugLog "imported temp size mm after normalize " & Format$(Abs(importedShape.RightX - importedShape.LeftX), "0.###") & " x " & Format$(Abs(importedShape.TopY - importedShape.BottomY), "0.###")

    tempStage = "place temp asset"
    MimakiDebugLog tempStage
    If AUTO_ROTATE_TO_SLOT Then
        appliedRotation = RotateShapeToSlotOrientation(importedShape, slot, effectiveInputOrientation)
    End If
    CenterShapeOnPoint importedShape, slot.CenterX, slot.CenterY

    Set PlaceSourceRangeIntoSlotViaTempFile = importedShape

CleanUp:
    On Error Resume Next
    If Not tempShape Is Nothing Then tempShape.Delete
    ClearTemporaryRasterLayer sourceWorkPage, True
    If Len(tempPath) > 0 Then
        If FileExists(tempPath) Then Kill tempPath
    End If
    On Error GoTo 0
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.Description
    If Len(tempStage) > 0 Then errDesc = tempStage & ": " & CStr(errNum) & " " & errDesc
    MimakiDebugLog "FAILED " & errDesc
    Resume CleanUpRaise

CleanUpRaise:
    On Error Resume Next
    If Not tempShape Is Nothing Then tempShape.Delete
    ClearTemporaryRasterLayer sourceWorkPage, True
    If Len(tempPath) > 0 Then
        If FileExists(tempPath) Then Kill tempPath
    End If
    On Error GoTo 0
    If errNum >= vbObjectError Then
        Err.Raise errNum, "PlaceSourceRangeIntoSlotViaTempFile", errDesc
    Else
        Err.Raise vbObjectError + 515, "PlaceSourceRangeIntoSlotViaTempFile", errDesc
    End If
End Function

Private Sub ExportTemporaryShapeToFile(ByVal srcDoc As Document, ByVal tempShape As Shape, ByVal tempPath As String, ByVal rasterizeToBitmap As Boolean)
    Dim ex As ExportFilter
    Dim widthPx As Long
    Dim heightPx As Long
    Dim exportErr As Long
    Dim exportDesc As String

    If tempShape Is Nothing Then
        Err.Raise vbObjectError + 509, "ExportTemporaryShapeToFile", "Temporary source shape is empty."
    End If
    If Len(Trim$(tempPath)) = 0 Then
        Err.Raise vbObjectError + 510, "ExportTemporaryShapeToFile", "Temporary export path is empty."
    End If

    srcDoc.Activate
    srcDoc.ClearSelection
    tempShape.CreateSelection
    widthPx = ShapeWidthPixels(tempShape, RASTERIZE_SOURCE_DPI)
    heightPx = ShapeHeightPixels(tempShape, RASTERIZE_SOURCE_DPI)

    On Error Resume Next
    Err.Clear
    If rasterizeToBitmap Then
        MimakiDebugLog "try ExportBitmap TIFF direct"
        Set ex = srcDoc.ExportBitmap( _
            tempPath, _
            cdrTIFF, _
            cdrSelection, _
            cdrCMYKColorImage, _
            widthPx, heightPx, _
            RASTERIZE_SOURCE_DPI, RASTERIZE_SOURCE_DPI, _
            cdrNormalAntiAliasing, _
            False, False, True, False, _
            cdrCompressionNone)
        exportErr = Err.Number
        exportDesc = Err.Description
    Else
        MimakiDebugLog "try PublishToPDF temp"
        ConfigurePDFForTempSelection srcDoc
        srcDoc.PublishToPDF tempPath
        exportErr = Err.Number
        exportDesc = Err.Description
        On Error GoTo 0
        If Not FileExists(tempPath) Then
            Err.Raise vbObjectError + 512, "ExportTemporaryShapeToFile", "Temporary source PDF was not created. " & CStr(exportErr) & " " & exportDesc
        End If
        Exit Sub
    End If
    If ex Is Nothing Or exportErr <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError + 511, "ExportTemporaryShapeToFile", "Temporary source export failed. Last export error: " & CStr(exportErr) & " " & exportDesc
    End If

    Err.Clear
    MimakiDebugLog "finish temp export"
    ex.Finish
    If Err.Number <> 0 Then
        exportErr = Err.Number
        exportDesc = Err.Description
        On Error GoTo 0
        Err.Raise vbObjectError + 514, "ExportTemporaryShapeToFile", "Temporary source export finish failed. " & CStr(exportErr) & " " & exportDesc
    End If
    On Error GoTo 0

    If Not FileExists(tempPath) Then
        Err.Raise vbObjectError + 512, "ExportTemporaryShapeToFile", "Temporary source file was not created."
    End If
    MimakiDebugLog "temp export ok " & tempPath
End Sub

Private Sub ConfigurePDFForTempSelection(ByVal doc As Document)
    Dim pdfSettings As Object

    On Error Resume Next
    Set pdfSettings = CallByName(doc, "PDFSettings", VbGet)
    If pdfSettings Is Nothing Then Exit Sub

    CallByName pdfSettings, "PublishRange", VbLet, pdfPageRange
    CallByName pdfSettings, "PageRange", VbLet, CStr(doc.ActivePage.Index)
    CallByName pdfSettings, "SelectionOnly", VbLet, True
    CallByName pdfSettings, "UseColorProfile", VbLet, True
    CallByName pdfSettings, "EmbedFonts", VbLet, True
    On Error GoTo 0
End Sub

Private Function ImportTemporaryShapeFile(ByVal outPage As Page, ByVal artLayer As Layer, ByVal tempPath As String) As Shape
    Dim importFilter As ImportFilter
    Dim importLayer As Layer
    Dim importErr As Long
    Dim importDesc As String

    If outPage Is Nothing Then Exit Function
    If artLayer Is Nothing Then Exit Function
    If Not FileExists(tempPath) Then
        Err.Raise vbObjectError + 513, "ImportTemporaryShapeFile", "Temporary source file does not exist."
    End If

    outPage.Activate
    Set importLayer = EnsureLayer(outPage, TEMP_IMPORT_LAYER_NAME)
    PrepareTemporaryImportLayer importLayer
    ActiveDocument.ClearSelection

    On Error Resume Next
    Err.Clear
    MimakiDebugLog "try ImportEx temp direct"
    importLayer.Activate
    Set importFilter = importLayer.ImportEx(tempPath, TempFileImportFilter(tempPath))
    importErr = Err.Number
    importDesc = Err.Description
    If Not importFilter Is Nothing And importErr = 0 Then
        Err.Clear
        importFilter.Finish
        If Err.Number = 0 Then
            Set ImportTemporaryShapeFile = TakeImportedShapeFromLayer(importLayer, artLayer)
            If Not ImportTemporaryShapeFile Is Nothing Then Exit Function
        End If
        importErr = Err.Number
        importDesc = Err.Description
    End If
    MimakiDebugLog "ImportEx temp failed/empty " & CStr(importErr) & " " & importDesc

    Err.Clear
    MimakiDebugLog "try Import temp direct"
    PrepareTemporaryImportLayer importLayer
    importLayer.Activate
    importLayer.Import tempPath, TempFileImportFilter(tempPath)
    If Err.Number = 0 Then
        Set ImportTemporaryShapeFile = TakeImportedShapeFromLayer(importLayer, artLayer)
        If Not ImportTemporaryShapeFile Is Nothing Then Exit Function
    End If
    importErr = Err.Number
    importDesc = Err.Description
    MimakiDebugLog "Import temp failed/empty " & CStr(importErr) & " " & importDesc

    Err.Clear
    Set ImportTemporaryShapeFile = TakeImportedShapeFromLayer(importLayer, artLayer)
    On Error GoTo 0
End Function

Private Sub PrepareTemporaryImportLayer(ByVal importLayer As Layer)
    Dim i As Long

    If importLayer Is Nothing Then Exit Sub

    On Error Resume Next
    importLayer.Visible = True
    importLayer.Editable = True
    importLayer.Printable = False
    For i = importLayer.Shapes.Count To 1 Step -1
        importLayer.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Function TakeImportedShapeFromLayer(ByVal importLayer As Layer, ByVal artLayer As Layer) As Shape
    Dim importedShape As Shape

    If importLayer Is Nothing Then Exit Function
    If artLayer Is Nothing Then Exit Function

    On Error Resume Next
    If importLayer.Shapes.Count > 1 Then
        importLayer.Shapes.All.CreateSelection
        Set importedShape = ActiveSelection.Group
    ElseIf importLayer.Shapes.Count = 1 Then
        Set importedShape = importLayer.Shapes(1)
    Else
        Set importedShape = NormalizePastedObjectToSingleShape(ActiveSelectionRange)
    End If

    If Not importedShape Is Nothing Then
        Err.Clear
        CallByName importedShape, "MoveToLayer", VbMethod, artLayer
        If Err.Number <> 0 Then
            Err.Clear
            importedShape.CreateSelection
            ActiveSelection.Cut
            artLayer.Activate
            Set importedShape = NormalizePastedObjectToSingleShape(artLayer.Paste)
        End If
        artLayer.Activate
        importedShape.CreateSelection
        Set TakeImportedShapeFromLayer = importedShape
    End If
    PrepareTemporaryImportLayer importLayer
    On Error GoTo 0
End Function

Private Function TempFileImportFilter(ByVal tempPath As String) As Long
    Dim ext As String

    ext = LCase$(Mid$(tempPath, InStrRev(tempPath, ".") + 1))
    If ext = "tif" Or ext = "tiff" Then
        TempFileImportFilter = cdrTIFF
    ElseIf ext = "pdf" Then
        TempFileImportFilter = cdrPDF
    Else
        TempFileImportFilter = cdrAutoSense
    End If
End Function

Private Sub NormalizeImportedCanvasSize(ByVal sh As Shape, ByVal canvasWidth As Double, ByVal canvasHeight As Double)
    Dim oldRefPoint As Long
    Dim currentWidth As Double
    Dim currentHeight As Double
    Dim currentAspect As Double
    Dim targetAspect As Double
    Dim scaleFactor As Double
    Dim newWidth As Double
    Dim newHeight As Double

    If sh Is Nothing Then Exit Sub
    If canvasWidth <= 0# Or canvasHeight <= 0# Then Exit Sub
    currentWidth = Abs(sh.RightX - sh.LeftX)
    currentHeight = Abs(sh.TopY - sh.BottomY)
    If currentWidth <= 0# Or currentHeight <= 0# Then Exit Sub

    currentAspect = currentWidth / currentHeight
    targetAspect = canvasWidth / canvasHeight
    If Abs(currentAspect - targetAspect) > 0.05 Then
        MimakiDebugLog "skip temp size normalize, aspect mismatch " & Format$(currentAspect, "0.###") & " vs " & Format$(targetAspect, "0.###")
        Exit Sub
    End If

    scaleFactor = canvasWidth / currentWidth
    If scaleFactor <= 0# Then Exit Sub
    If Abs(1# - scaleFactor) < 0.005 Then Exit Sub
    newWidth = currentWidth * scaleFactor
    newHeight = currentHeight * scaleFactor

    On Error Resume Next
    oldRefPoint = ActiveDocument.ReferencePoint
    ActiveDocument.ReferencePoint = cdrCenter
    Err.Clear
    sh.SetSize newWidth, newHeight
    If Err.Number <> 0 Then
        Err.Clear
        CallByName sh, "SetSize", VbMethod, newWidth, newHeight
    End If
    ActiveDocument.ReferencePoint = oldRefPoint
    On Error GoTo 0
End Sub

Private Function NormalizeImportedObjectToSingleShape(ByVal importedObj As Object) As Shape
    Set NormalizeImportedObjectToSingleShape = NormalizePastedObjectToSingleShape(importedObj)
    If Not NormalizeImportedObjectToSingleShape Is Nothing Then Exit Function

    On Error Resume Next
    Set NormalizeImportedObjectToSingleShape = NormalizePastedObjectToSingleShape(ActiveSelectionRange)
    On Error GoTo 0
End Function

Private Function BuildMimakiTempAssetPath(ByVal rasterizeToBitmap As Boolean) As String
    Dim folderPath As String
    Dim extensionName As String

    folderPath = CombinePath(TEMP_WORK_ROOT_PATH, "MimakiImposer")
    EnsureFolderExists folderPath
    If rasterizeToBitmap Then
        extensionName = ".tif"
    Else
        extensionName = ".pdf"
    End If

    BuildMimakiTempAssetPath = CombinePath(folderPath, "mki_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int((999999# * Rnd) + 1)) & extensionName)
End Function

Private Function MillimetersToPixels(ByVal valueMm As Double, ByVal dpi As Long) As Long
    MillimetersToPixels = CLng((Abs(valueMm) / 25.4) * dpi)
    If MillimetersToPixels < 1 Then MillimetersToPixels = 1
End Function

Private Function ShapeWidthPixels(ByVal sh As Shape, ByVal dpi As Long) As Long
    Dim w As Double

    On Error Resume Next
    w = Abs(sh.RightX - sh.LeftX)
    If w <= 0# Then w = SOURCE_CARD_CANVAS_LONG_MM
    ShapeWidthPixels = CLng((w / 25.4) * dpi)
    If ShapeWidthPixels < 1 Then ShapeWidthPixels = 1
    On Error GoTo 0
End Function

Private Function ShapeHeightPixels(ByVal sh As Shape, ByVal dpi As Long) As Long
    Dim h As Double

    On Error Resume Next
    h = Abs(sh.TopY - sh.BottomY)
    If h <= 0# Then h = SOURCE_CARD_CANVAS_SHORT_MM
    ShapeHeightPixels = CLng((h / 25.4) * dpi)
    If ShapeHeightPixels < 1 Then ShapeHeightPixels = 1
    On Error GoTo 0
End Function
Private Sub ClearTemporaryRasterLayersInDocument(ByVal doc As Document, Optional ByVal deleteLayer As Boolean = False)
    Dim pg As Page

    If doc Is Nothing Then Exit Sub

    On Error Resume Next
    For Each pg In doc.Pages
        ClearTemporaryRasterLayer pg, deleteLayer
    Next pg
    On Error GoTo 0
End Sub

Private Sub ClearTemporaryRasterLayer(ByVal pg As Page, Optional ByVal deleteLayer As Boolean = False)
    Dim tempLayer As Layer
    Dim sh As Shape
    Dim i As Long

    If pg Is Nothing Then Exit Sub

    On Error Resume Next
    Set tempLayer = pg.Layers(TEMP_RASTER_LAYER_NAME)
    If tempLayer Is Nothing Then Exit Sub

    tempLayer.Visible = True
    tempLayer.Editable = True
    For i = tempLayer.Shapes.Count To 1 Step -1
        Set sh = tempLayer.Shapes(i)
        If Not sh Is Nothing Then sh.Delete
    Next i
    tempLayer.Printable = False
    If deleteLayer Then tempLayer.Delete
    On Error GoTo 0
End Sub

Private Sub MimakiDebugLog(ByVal messageText As String)
    Dim folderPath As String
    Dim filePath As String
    Dim fileNo As Integer

    On Error Resume Next
    folderPath = CombinePath(TEMP_WORK_ROOT_PATH, "MimakiImposer")
    EnsureFolderExists folderPath
    filePath = CombinePath(folderPath, "mimaki_debug.log")
    fileNo = FreeFile
    Open filePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Close #fileNo
    On Error GoTo 0
End Sub

Private Function NormalizePastedObjectToSingleShape(ByVal pastedObj As Object) As Shape
    On Error Resume Next

    If pastedObj Is Nothing Then Exit Function

    If TypeName(pastedObj) = "ShapeRange" Then
        If pastedObj.Count > 1 Then
            pastedObj.CreateSelection
            Set NormalizePastedObjectToSingleShape = ActiveSelection.Group
            If NormalizePastedObjectToSingleShape Is Nothing Then
                Set NormalizePastedObjectToSingleShape = pastedObj.Group
            End If
        ElseIf pastedObj.Count = 1 Then
            Set NormalizePastedObjectToSingleShape = pastedObj.Shapes(1)
        End If
    ElseIf TypeName(pastedObj) = "Shape" Then
        Set NormalizePastedObjectToSingleShape = pastedObj
    End If

    On Error GoTo 0
End Function

Private Function RasterizePlacedShapeForProduction(ByVal sh As Shape) As Shape
    Dim obj As Object
    Dim sr As ShapeRange

    If sh Is Nothing Then Exit Function

    On Error GoTo EH
    sh.CreateSelection
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        Err.Raise vbObjectError + 504, "RasterizePlacedShapeForProduction", "Raster selection is empty."
    End If
    If sr.Count = 0 Then
        Err.Raise vbObjectError + 504, "RasterizePlacedShapeForProduction", "Raster selection is empty."
    End If

    ' CorelDRAW 2026 v27.2 ConvertToBitmapEx has one Resolution argument.
    Set obj = CallByName(sr, "ConvertToBitmapEx", VbMethod, cdrCMYKColorImage, False, False, RASTERIZE_SOURCE_DPI, cdrNormalAntiAliasing, True, True, 95)
    If obj Is Nothing Then
        Err.Raise vbObjectError + 504, "RasterizePlacedShapeForProduction", "600 dpi temp raster failed. Corel did not return a bitmap shape."
    End If

    Set RasterizePlacedShapeForProduction = NormalizePastedObjectToSingleShape(obj)
    If RasterizePlacedShapeForProduction Is Nothing Then
        Err.Raise vbObjectError + 505, "RasterizePlacedShapeForProduction", "600 dpi temp raster returned unsupported object type: " & TypeName(obj)
    End If
    Exit Function

EH:
    If Err.Number = 0 Then
        Err.Raise vbObjectError + 504, "RasterizePlacedShapeForProduction", "600 dpi ConvertToBitmapEx failed."
    Else
        Err.Raise Err.Number, "RasterizePlacedShapeForProduction", Err.Description
    End If
End Function

Private Sub CopySourceRangeForX8(ByVal srcRange As ShapeRange, Optional ByVal rasterizeBeforeCopy As Boolean = False, Optional ByVal srcPage As Object = Nothing, Optional ByVal canvasWidth As Double = 0#, Optional ByVal canvasHeight As Double = 0#, Optional ByVal cropBeforeCopy As Boolean = False)
    Dim tempShape As Shape
    Dim workPage As Page

    If srcRange.Count <= 0 Then
        Err.Raise vbObjectError + 503, "CopySourceRangeForX8", "Source range is empty."
    End If

    If rasterizeBeforeCopy Or cropBeforeCopy Then
        On Error Resume Next
        If srcPage Is Nothing Then
            Set workPage = ActiveDocument.ActivePage
        Else
            Set workPage = srcPage
        End If
        On Error GoTo 0
        ClearTemporaryRasterLayer workPage, True
        Set tempShape = CreateTemporarySourceCanvasCopy(srcRange, workPage, canvasWidth, canvasHeight, rasterizeBeforeCopy)
        tempShape.Copy
        tempShape.Delete
        ClearTemporaryRasterLayer workPage, True
        Exit Sub
    End If

    If srcRange.Count = 1 Then
        srcRange.Shapes(1).Copy
    Else
        ActiveDocument.ClearSelection
        srcRange.CreateSelection
        ActiveSelection.Copy
    End If
End Sub

Private Function CreateTemporaryRasterizedSourceCanvasCopy(ByVal srcRange As ShapeRange, ByVal srcPage As Object, ByVal canvasWidth As Double, ByVal canvasHeight As Double) As Shape
    Set CreateTemporaryRasterizedSourceCanvasCopy = CreateTemporarySourceCanvasCopy(srcRange, srcPage, canvasWidth, canvasHeight, True)
End Function

Private Function CreateTemporarySourceCanvasCopy(ByVal srcRange As ShapeRange, ByVal srcPage As Object, ByVal canvasWidth As Double, ByVal canvasHeight As Double, ByVal rasterizeCanvas As Boolean, Optional ByVal useExplicitCropCenter As Boolean = False, Optional ByVal explicitCenterX As Double = 0#, Optional ByVal explicitCenterY As Double = 0#) As Shape
    Dim dupObj As Object
    Dim dupShape As Shape
    Dim frameShape As Shape
    Dim rasterSource As Shape
    Dim tempLayer As Layer
    Dim workPage As Page
    Dim cropCenterX As Double
    Dim cropCenterY As Double

    If srcRange.Count <= 0 Then
        Err.Raise vbObjectError + 503, "CreateTemporarySourceCanvasCopy", "Source range is empty."
    End If

    On Error GoTo EH

    If srcPage Is Nothing Then
        Set workPage = ActiveDocument.ActivePage
    Else
        Set workPage = srcPage
    End If
    If canvasWidth <= 0# Or canvasHeight <= 0# Then
        canvasWidth = SOURCE_CARD_CANVAS_LONG_MM
        canvasHeight = SOURCE_CARD_CANVAS_SHORT_MM
    End If

    workPage.Activate
    Set tempLayer = EnsureLayer(workPage, TEMP_RASTER_LAYER_NAME)
    tempLayer.Visible = True
    tempLayer.Editable = True
    tempLayer.Printable = True
    If useExplicitCropCenter Then
        cropCenterX = explicitCenterX
        cropCenterY = explicitCenterY
    Else
        GetSourceCropCenter workPage, srcRange, canvasWidth, canvasHeight, cropCenterX, cropCenterY
    End If

    Set frameShape = tempLayer.CreateRectangle2(cropCenterX - (canvasWidth / 2#), cropCenterY + (canvasHeight / 2#), canvasWidth, canvasHeight)
    frameShape.Name = "MIMAKI_TEMP_SLOT_CANVAS"
    If rasterizeCanvas Then
        frameShape.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
    Else
        frameShape.Fill.ApplyNoFill
    End If
    frameShape.Outline.SetNoOutline
    On Error Resume Next
    frameShape.OrderToBack
    On Error GoTo EH

    If srcRange.Count = 1 Then
        Set dupShape = srcRange.Shapes(1).Duplicate(0#, 0#)
    Else
        ActiveDocument.ClearSelection
        srcRange.CreateSelection
        Err.Clear
        Set dupObj = CallByName(ActiveSelectionRange, "Duplicate", VbMethod, 0#, 0#)
        Set dupShape = NormalizePastedObjectToSingleShape(dupObj)
    End If

    If dupShape Is Nothing Then
        Err.Raise vbObjectError + 506, "CreateTemporarySourceCanvasCopy", "Could not create temporary source copy."
    End If

    If Not TryAddShapeToPowerClip(dupShape, frameShape) Then
        Err.Raise vbObjectError + 507, "CreateTemporarySourceCanvasCopy", "PowerClip slot canvas failed. Cannot safely crop oversized source artwork."
    End If

    If rasterizeCanvas Then
        Set rasterSource = frameShape
        Set CreateTemporarySourceCanvasCopy = RasterizePlacedShapeForProduction(rasterSource)
        MimakiDebugLog "temp rasterized canvas size mm " & Format$(Abs(CreateTemporarySourceCanvasCopy.RightX - CreateTemporarySourceCanvasCopy.LeftX), "0.###") & " x " & Format$(Abs(CreateTemporarySourceCanvasCopy.TopY - CreateTemporarySourceCanvasCopy.BottomY), "0.###")
    Else
        Set CreateTemporarySourceCanvasCopy = frameShape
    End If
    Exit Function

EH:
    Dim localErrNum As Long
    Dim localErrDesc As String

    localErrNum = Err.Number
    localErrDesc = Err.Description
    On Error Resume Next
    If Not dupShape Is Nothing Then dupShape.Delete
    If Not frameShape Is Nothing Then frameShape.Delete
    On Error GoTo 0
    Err.Raise localErrNum, "CreateTemporarySourceCanvasCopy", localErrDesc
End Function

Private Function TryAddShapeToPowerClip(ByVal contentShape As Shape, ByVal containerShape As Shape) As Boolean
    If contentShape Is Nothing Or containerShape Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    CallByName contentShape, "AddToPowerClip", VbMethod, containerShape, False
    If Err.Number = 0 Then
        TryAddShapeToPowerClip = True
        Exit Function
    End If
    On Error GoTo 0
End Function

Private Function IsBitmapShape(ByVal sh As Shape) As Boolean
    If sh Is Nothing Then Exit Function

    On Error Resume Next
    IsBitmapShape = (sh.Type = cdrBitmapShape)
    On Error GoTo 0
End Function

Private Function DuplicatePlacedShapeIntoSlot(ByVal sourceShape As Shape, ByRef sourceSlot As TSlotInfo, ByRef targetSlot As TSlotInfo) As Shape
    Dim dup As Shape
    Dim dx As Double
    Dim dy As Double

    If sourceShape Is Nothing Then Exit Function
    Set dup = sourceShape.Duplicate(0#, 0#)
    dx = targetSlot.CenterX - sourceSlot.CenterX
    dy = targetSlot.CenterY - sourceSlot.CenterY
    dup.Move dx, dy
    Set DuplicatePlacedShapeIntoSlot = dup
End Function

Private Sub DuplicatePlacedShapeIntoLayerSlot(ByVal sourceShape As Shape, ByRef sourceSlot As TSlotInfo, ByRef targetSlot As TSlotInfo)
    Dim dup As Shape

    If sourceShape Is Nothing Then Exit Sub
    Set dup = DuplicatePlacedShapeIntoSlot(sourceShape, sourceSlot, targetSlot)
End Sub

Private Function SlotsHaveUniformOrientation(ByRef slots() As TSlotInfo) As Boolean
    Dim i As Long
    Dim firstOrientation As Long
    Dim currentOrientation As Long

    On Error Resume Next
    If UBound(slots) < LBound(slots) Then Exit Function
    On Error GoTo 0

    firstOrientation = SlotOrientationValue(slots(LBound(slots)))
    If firstOrientation = 0 Then Exit Function

    For i = LBound(slots) + 1 To UBound(slots)
        currentOrientation = SlotOrientationValue(slots(i))
        If currentOrientation <> firstOrientation Then Exit Function
    Next i

    SlotsHaveUniformOrientation = True
End Function

Private Function SlotOrientationValue(ByRef slot As TSlotInfo) As Long
    Dim slotWidth As Double
    Dim slotHeight As Double

    slotWidth = Abs(slot.RightX - slot.LeftX)
    slotHeight = Abs(slot.TopY - slot.BottomY)

    If slotWidth <= 0 Or slotHeight <= 0 Then Exit Function
    If Abs(slotWidth - slotHeight) <= 0.2 Then Exit Function

    If slotWidth > slotHeight Then
        SlotOrientationValue = 1
    Else
        SlotOrientationValue = 2
    End If
End Function

' ============================================================
' Template handling
' ============================================================

Private Function ReadTemplateInfo(ByVal tplDoc As Document, ByVal layoutName As String, ByRef slots() As TSlotInfo) As TLayoutInfo
    Dim pg As Page
    Dim ly As Layer
    Dim layoutNameNorm As String

    layoutNameNorm = NormalizeLayoutName(layoutName)

    Set ly = FindMasterLayoutLayerByDisplayName(tplDoc, layoutNameNorm)
    If Not ly Is Nothing Then
        ReadTemplateInfo = BuildLayoutInfoFromLayer(ly, tplDoc.Pages(1).SizeWidth, tplDoc.Pages(1).SizeHeight, 1, layoutNameNorm, slots)
        If ReadTemplateInfo.SlotCount > 0 Then Exit Function
    End If

    For Each pg In tplDoc.Pages
        Set ly = FindLayoutLayerOnPage(pg, layoutNameNorm)
        If Not ly Is Nothing Then
            ReadTemplateInfo = BuildLayoutInfoFromLayer(ly, pg.SizeWidth, pg.SizeHeight, pg.Index, layoutNameNorm, slots)
            If ReadTemplateInfo.SlotCount > 0 Then
                Exit Function
            End If
        End If
    Next pg
End Function

Private Function BuildLayoutInfoFromLayer(ByVal ly As Layer, ByVal pageWidth As Double, ByVal pageHeight As Double, ByVal pageIndex As Long, ByVal layoutName As String, ByRef slots() As TSlotInfo) As TLayoutInfo
    Dim sh As Shape
    Dim slotCount As Long

    slotCount = 0
    Erase slots

    For Each sh In ly.Shapes
        If IsTemplateSlotShape(sh) Then
            slotCount = slotCount + 1
            ReDim Preserve slots(1 To slotCount)
            slots(slotCount) = ShapeToSlotInfo(sh, layoutName)
        End If
    Next sh

    If slotCount > 0 Then
        SortSlots slots
        BuildLayoutInfoFromLayer.PageWidth = pageWidth
        BuildLayoutInfoFromLayer.PageHeight = pageHeight
        BuildLayoutInfoFromLayer.SlotCount = slotCount
        BuildLayoutInfoFromLayer.SourcePageIndex = pageIndex
        BuildLayoutInfoFromLayer.LayoutName = layoutName
        BuildLayoutInfoFromLayer.SlotLayerName = NormalizeLayerToken(ly.Name)
    End If
End Function

Private Function ShapeToSlotInfo(ByVal sh As Shape, ByVal layoutName As String) As TSlotInfo
    Dim s As TSlotInfo

    s.Name = sh.Name
    s.LayerName = layoutName
    s.LeftX = sh.LeftX
    s.RightX = sh.RightX
    s.TopY = sh.TopY
    s.BottomY = sh.BottomY
    s.CenterX = (sh.LeftX + sh.RightX) / 2#
    s.CenterY = (sh.TopY + sh.BottomY) / 2#

    ShapeToSlotInfo = s
End Function

Private Function LayerLooksLikeLayout(ByVal ly As Layer) As Boolean
    If Not IsSlotsLayerName(ly.Name) Then Exit Function
    LayerLooksLikeLayout = LayerHasTemplateSlots(ly)
End Function

Private Function LayerHasTemplateSlots(ByVal ly As Layer) As Boolean
    Dim sh As Shape

    If ly Is Nothing Then Exit Function
    For Each sh In ly.Shapes
        If IsTemplateSlotShape(sh) Then
            LayerHasTemplateSlots = True
            Exit Function
        End If
    Next sh
End Function

Private Sub SortSlots(ByRef slots() As TSlotInfo)
    Dim i As Long, j As Long
    Dim tmp As TSlotInfo

    On Error Resume Next
    If UBound(slots) <= LBound(slots) Then Exit Sub
    On Error GoTo 0

    For i = LBound(slots) To UBound(slots) - 1
        For j = i + 1 To UBound(slots)
            If ShouldSwapSlot(slots(i), slots(j)) Then
                tmp = slots(i)
                slots(i) = slots(j)
                slots(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function ShouldSwapSlot(ByRef a As TSlotInfo, ByRef b As TSlotInfo) As Boolean
    If a.CenterY > b.CenterY Then
        ShouldSwapSlot = True
    ElseIf a.CenterY = b.CenterY Then
        If a.CenterX < b.CenterX Then ShouldSwapSlot = True
    End If
End Function

Private Sub EnsureImpositionStructure(ByVal tplPage As Page, ByVal tplLayer As Layer, ByVal outDoc As Document, ByRef layout As TLayoutInfo)
    Dim p As Long

    For p = 1 To outDoc.Pages.Count
        EnsureImpositionStructureForPage tplPage, tplLayer, outDoc.Pages(p), layout
    Next p
End Sub

Private Sub EnsureImpositionStructureForPage(ByVal tplPage As Page, ByVal tplLayer As Layer, ByVal outPage As Page, ByRef layout As TLayoutInfo)
    Dim dstTemplateLayer As Layer
    Dim artLayer As Layer
    Dim whiteLayer As Layer

    outPage.SetSize layout.PageWidth, layout.PageHeight

    Set dstTemplateLayer = EnsureLayer(outPage, layout.LayoutName)
    Set whiteLayer = EnsureLayer(outPage, WHITE_LAYER_NAME)
    Set artLayer = EnsureLayer(outPage, PRINT_DATA_LAYER_NAME)

    If dstTemplateLayer.Shapes.Count = 0 Then
        dstTemplateLayer.Editable = True
        CopyLayerShapes tplLayer, dstTemplateLayer
    End If

    dstTemplateLayer.Visible = True
    dstTemplateLayer.Printable = False
    MoveLayerAboveReference dstTemplateLayer, artLayer
    dstTemplateLayer.Editable = False

    artLayer.Visible = True
    artLayer.Printable = True
    artLayer.Editable = True

    whiteLayer.Visible = True
    whiteLayer.Printable = True
    whiteLayer.Editable = True
End Sub

Private Sub CopyLayerShapes(ByVal srcLayer As Layer, ByVal dstLayer As Layer)
    Dim pasted As Object
    Dim srcWasEditable As Boolean
    Dim dstWasEditable As Boolean

    On Error GoTo EH

    If srcLayer.Shapes.Count = 0 Then Exit Sub

    srcWasEditable = srcLayer.Editable
    dstWasEditable = dstLayer.Editable
    srcLayer.Editable = True
    dstLayer.Editable = True

    srcLayer.Shapes.All.Copy
    dstLayer.Activate
    Set pasted = dstLayer.Paste
    If Not pasted Is Nothing Then
        If TypeName(pasted) = "ShapeRange" Then
            If pasted.Count > 1 Then pasted.Group
        End If
    End If

CleanUp:
    On Error Resume Next
    srcLayer.Editable = srcWasEditable
    dstLayer.Editable = dstWasEditable
    On Error GoTo 0
    Exit Sub

EH:
    Resume CleanUp
End Sub

Private Sub MoveLayerAboveReference(ByVal targetLayer As Layer, ByVal referenceLayer As Layer)
    If targetLayer Is Nothing Then Exit Sub
    If referenceLayer Is Nothing Then Exit Sub

    On Error Resume Next
    targetLayer.Editable = True
    CallByName targetLayer, "MoveAbove", VbMethod, referenceLayer
    If Err.Number <> 0 Then
        Err.Clear
        CallByName targetLayer, "OrderAbove", VbMethod, referenceLayer
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Layer / document helpers
' ============================================================

Private Function EnsureLayer(ByVal pg As Page, ByVal layerName As String) As Layer
    Dim ly As Layer
    Set ly = FindLayerByName(pg, layerName)
    If ly Is Nothing Then
        Set ly = pg.CreateLayer(layerName)
    End If
    Set EnsureLayer = ly
End Function

Private Function EnsureArtworkLayer(ByVal pg As Page) As Layer
    Set EnsureArtworkLayer = EnsureLayer(pg, PRINT_DATA_LAYER_NAME)
End Function

Private Function EnsureWhiteLayer(ByVal pg As Page) As Layer
    Set EnsureWhiteLayer = EnsureLayer(pg, WHITE_LAYER_NAME)
End Function

Private Sub ClearArtworkLayers(ByVal doc As Document)
    Dim pg As Page
    Dim artLayer As Layer
    Dim whiteLayer As Layer

    On Error Resume Next
    For Each pg In doc.Pages
        Set artLayer = FindLayerByName(pg, PRINT_DATA_LAYER_NAME)
        If Not artLayer Is Nothing Then
            artLayer.Editable = True
            artLayer.Visible = True
            artLayer.Shapes.All.Delete
        End If
        Set whiteLayer = FindLayerByName(pg, WHITE_LAYER_NAME)
        If Not whiteLayer Is Nothing Then
            whiteLayer.Editable = True
            whiteLayer.Visible = True
            whiteLayer.Shapes.All.Delete
        End If
    Next pg
    On Error GoTo 0
End Sub

Private Function FindLayerByName(ByVal pg As Page, ByVal layerName As String) As Layer
    Dim ly As Layer
    For Each ly In pg.Layers
        If LayerNameMatches(ly.Name, layerName) Then
            Set FindLayerByName = ly
            Exit Function
        End If
    Next ly
End Function

Private Function FindLayoutLayerOnPage(ByVal pg As Page, ByVal layoutName As String) As Layer
    Dim ly As Layer

    If pg Is Nothing Then Exit Function
    For Each ly In pg.Layers
        If IsSlotsLayerName(ly.Name) Then
            If LayoutNamesEquivalent(GetLayoutDisplayNameFromLayerName(ly.Name), layoutName) Then
                Set FindLayoutLayerOnPage = ly
                Exit Function
            End If
        End If
    Next ly
End Function

Private Function FindTemplateLayerByName(ByVal doc As Document, ByVal layerName As String) As Layer
    Dim pg As Page
    Dim ly As Layer

    For Each pg In doc.Pages
        Set ly = FindLayoutLayerOnPage(pg, layerName)
        If Not ly Is Nothing Then
            Set FindTemplateLayerByName = ly
            Exit Function
        End If
    Next pg

    Set FindTemplateLayerByName = FindMasterLayoutLayerByDisplayName(doc, layerName)
End Function

Private Function FindMasterLayoutLayerByDisplayName(ByVal doc As Document, ByVal layoutName As String) As Layer
    Dim masterPage As Object
    Dim ly As Layer

    On Error Resume Next
    Set masterPage = CallByName(doc, "MasterPage", VbGet)
    On Error GoTo 0

    If masterPage Is Nothing Then Exit Function

    For Each ly In masterPage.Layers
        If IsSlotsLayerName(ly.Name) Then
            If LayoutNamesEquivalent(GetLayoutDisplayNameFromLayerName(ly.Name), layoutName) Then
                Set FindMasterLayoutLayerByDisplayName = ly
                Exit Function
            End If
        End If
    Next ly
End Function

Private Function FindMasterLayerByName(ByVal doc As Document, ByVal layerName As String) As Layer
    Dim masterPage As Object
    Dim ly As Layer

    On Error Resume Next
    Set masterPage = CallByName(doc, "MasterPage", VbGet)
    On Error GoTo 0

    If masterPage Is Nothing Then Exit Function

    For Each ly In masterPage.Layers
        If LayerNameMatches(ly.Name, layerName) Then
            Set FindMasterLayerByName = ly
            Exit Function
        End If
    Next ly
End Function

Private Function LayerNameMatches(ByVal actualName As String, ByVal wantedName As String) As Boolean
    LayerNameMatches = (NormalizeLayerToken(actualName) = NormalizeLayerToken(wantedName))
End Function

Private Function NormalizeLayerToken(ByVal rawName As String) As String
    NormalizeLayerToken = NormalizeLayoutName(GetLayerBaseName(rawName))
End Function

Private Function GetLayerBaseName(ByVal layerName As String) As String
    Dim p As Long
    Dim s As String

    s = Trim$(layerName)
    p = InStr(1, s, " (", vbTextCompare)
    If p > 1 Then s = Left$(s, p - 1)
    GetLayerBaseName = Trim$(s)
End Function

Private Function GetLayoutSlotsLayerName(ByVal layoutName As String) As String
    GetLayoutSlotsLayerName = Trim$(NormalizeLayoutName(layoutName)) & " " & SLOT_LAYER_SUFFIX
End Function

Private Function IsSlotsLayerName(ByVal layerName As String) As Boolean
    Dim normalizedName As String

    normalizedName = NormalizeLayerToken(layerName)
    If Len(normalizedName) <= Len(SLOT_LAYER_SUFFIX) Then Exit Function
    IsSlotsLayerName = (Right$(normalizedName, Len(SLOT_LAYER_SUFFIX)) = SLOT_LAYER_SUFFIX)
End Function

Private Function GetLayoutDisplayNameFromLayerName(ByVal layerName As String) As String
    Dim normalizedName As String

    normalizedName = NormalizeLayerToken(layerName)
    If Right$(normalizedName, Len(SLOT_LAYER_SUFFIX)) = SLOT_LAYER_SUFFIX Then
        GetLayoutDisplayNameFromLayerName = Trim$(Left$(normalizedName, Len(normalizedName) - Len(SLOT_LAYER_SUFFIX)))
    Else
        GetLayoutDisplayNameFromLayerName = normalizedName
    End If
End Function

Private Sub PrepareOutputDocument(ByVal outDoc As Document, ByVal pageWidth As Double, ByVal pageHeight As Double)
    outDoc.Unit = cdrMillimeter
    outDoc.ReferencePoint = cdrCenter
    outDoc.Pages(1).SetSize pageWidth, pageHeight
End Sub

Private Sub EnsureDocumentPages(ByVal doc As Document, ByVal requiredPages As Long, ByVal pageWidth As Double, ByVal pageHeight As Double)
    Dim p As Long

    If requiredPages < 1 Then requiredPages = 1

    If doc.Pages.Count < requiredPages Then
        doc.AddPages requiredPages - doc.Pages.Count
    End If

    For p = 1 To doc.Pages.Count
        doc.Pages(p).SetSize pageWidth, pageHeight
    Next p
End Sub

Private Function GetSelectableRange(ByVal srcDoc As Document, Optional ByVal preferSelection As Boolean = False) As ShapeRange
    Dim sr As ShapeRange

    On Error Resume Next
    srcDoc.Activate

    If preferSelection Then
        Set sr = ActiveSelectionRange
        Set GetSelectableRange = BuildFilteredSelectionRange(srcDoc, sr, False, True)
        Exit Function
    End If

    If srcDoc.SelectableShapes.Count > 0 Then
        Set GetSelectableRange = BuildFilteredSelectionRange(srcDoc, srcDoc.SelectableShapes.All, False, True)
    End If
End Function

Private Function GetPageSourceRange(ByVal doc As Document, ByVal pg As Page, Optional ByVal groupPageContent As Boolean = True) As ShapeRange
    On Error GoTo EH
    doc.Activate
    pg.Activate
    If pg.Shapes.Count <= 0 Then Exit Function
    Set GetPageSourceRange = SelectPageSourceShapes(doc, pg)
    Exit Function

EH:
    Err.Raise Err.Number, "GetPageSourceRange", "Page " & CStr(pg.Index) & ": " & Err.Description
End Function

Private Function SelectPageSourceShapes(ByVal doc As Document, ByVal pg As Page) As ShapeRange
    Dim sh As Shape
    Dim hasAny As Boolean

    If pg Is Nothing Then Exit Function

    doc.Activate
    pg.Activate
    ActiveDocument.ClearSelection
    For Each sh In pg.Shapes
        If IsSourceCandidateShape(sh) Then
            If Not hasAny Then
                sh.CreateSelection
                hasAny = True
            Else
                CallByName sh, "AddToSelection", VbMethod
            End If
        End If
    Next sh

    If hasAny Then Set SelectPageSourceShapes = ActiveSelectionRange
End Function

Private Function BuildFilteredSelectionRange(ByVal doc As Document, ByVal srcRange As ShapeRange, ByVal includeWhite As Boolean, Optional ByVal includeAllSourceShapes As Boolean = False) As ShapeRange
    Dim sh As Shape
    Dim hasAny As Boolean

    If srcRange Is Nothing Then Exit Function
    If srcRange.Count <= 0 Then Exit Function

    doc.Activate
    ActiveDocument.ClearSelection
    For Each sh In srcRange.Shapes
        If IsSourceCandidateShape(sh) Then
            If includeAllSourceShapes Or ShapeMatchesWhiteFilter(sh, includeWhite) Then
                If Not hasAny Then
                    sh.CreateSelection
                    hasAny = True
                Else
                    CallByName sh, "AddToSelection", VbMethod
                End If
            End If
        End If
    Next sh

    If hasAny Then Set BuildFilteredSelectionRange = ActiveSelectionRange
End Function

Private Function ShapeMatchesWhiteFilter(ByVal sh As Shape, ByVal includeWhite As Boolean) As Boolean
    If includeWhite Then
        ShapeMatchesWhiteFilter = IsWhiteSpotShape(sh)
    Else
        ShapeMatchesWhiteFilter = Not IsWhiteSpotShape(sh)
    End If
End Function

Private Function GetPageSourceShapeCount(ByVal pg As Page) As Long
    Dim sh As Shape

    If pg Is Nothing Then Exit Function
    For Each sh In pg.Shapes
        If IsSourceCandidateShape(sh) Then GetPageSourceShapeCount = GetPageSourceShapeCount + 1
    Next sh
End Function

Private Function IsSourceCandidateShape(ByVal sh As Shape) As Boolean
    Dim layerName As String

    If sh Is Nothing Then Exit Function
    On Error Resume Next
    layerName = sh.Layer.Name
    On Error GoTo 0
    If LayerNameMatches(layerName, TEMP_RASTER_LAYER_NAME) Then Exit Function
    If LayerNameMatches(layerName, TEMP_IMPORT_LAYER_NAME) Then Exit Function
    If IsSlotsLayerName(layerName) Then Exit Function
    IsSourceCandidateShape = True
End Function

Private Function IsWhiteSpotShape(ByVal sh As Shape) As Boolean
    If sh Is Nothing Then Exit Function
    On Error Resume Next
    If Not sh.Fill Is Nothing Then
        If SpotColorMatches(sh.Fill.UniformColor) Then
            IsWhiteSpotShape = True
            Exit Function
        End If
    End If
    If Not sh.Outline Is Nothing Then
        If SpotColorMatches(sh.Outline.Color) Then
            IsWhiteSpotShape = True
            Exit Function
        End If
    End If
    On Error GoTo 0
End Function

Private Function SpotColorMatches(ByVal clr As Object) As Boolean
    SpotColorMatches = ColorNameMatches(clr, WHITE_SPOT_NAME)
End Function

Private Function IsTemplateSlotShape(ByVal sh As Shape) As Boolean
    If sh Is Nothing Then Exit Function

    On Error Resume Next
    If Not sh.Outline Is Nothing Then
        If ColorNameMatches(sh.Outline.Color, SLOT_OUTLINE_SPOT_NAME) Then
            IsTemplateSlotShape = True
            Exit Function
        End If
    End If
    On Error GoTo 0
End Function

Private Function ColorNameMatches(ByVal clr As Object, ByVal expectedName As String) As Boolean
    Dim colorName As String

    On Error Resume Next
    colorName = Trim$(CStr(CallByName(clr, "Name", VbGet)))
    If Len(colorName) = 0 Then colorName = Trim$(CStr(CallByName(clr, "SpotName", VbGet)))
    On Error GoTo 0
    ColorNameMatches = (UCase$(colorName) = UCase$(expectedName))
End Function

Private Sub GroupSourcePagesInRange(ByVal doc As Document, ByVal firstPage As Long, ByVal lastPage As Long)
    Dim i As Long
    Dim groupedRange As ShapeRange

    NormalizeSourcePageRange doc, firstPage, lastPage
    doc.Activate
    For i = firstPage To lastPage
        If PageHasShapes(doc.Pages(i)) Then
            Set groupedRange = GroupPageContentAsSingleSource(doc, doc.Pages(i))
            If groupedRange Is Nothing Then
                Err.Raise vbObjectError + 563, "GroupSourcePagesInRange", "Page " & CStr(i) & ": source page could not be prepared as one group."
            End If
        End If
    Next i
End Sub

Private Function GroupPageContentAsSingleSource(ByVal doc As Document, ByVal pg As Page) As ShapeRange
    Dim sr As ShapeRange
    Dim activeSr As ShapeRange
    Dim groupedShape As Shape

    On Error GoTo EH
    doc.Activate
    pg.Activate
    If pg.Shapes.Count <= 0 Then Exit Function

    Set sr = pg.Shapes.All
    If sr Is Nothing Then Exit Function
    If sr.Count <= 0 Then Exit Function

    ActiveDocument.ClearSelection
    sr.CreateSelection
    Set activeSr = ActiveSelectionRange
    If activeSr Is Nothing Then
        Err.Raise vbObjectError + 560, "GroupPageContentAsSingleSource", "Could not select source objects."
    End If
    If activeSr.Count <= 0 Then
        Err.Raise vbObjectError + 561, "GroupPageContentAsSingleSource", "Source selection is empty."
    End If

    If activeSr.Count > 1 Then
        Set groupedShape = activeSr.Group
        If groupedShape Is Nothing Then
            Err.Raise vbObjectError + 562, "GroupPageContentAsSingleSource", "Could not group source objects."
        End If
        groupedShape.CreateSelection
    End If

    Set GroupPageContentAsSingleSource = ActiveSelectionRange
    Exit Function

EH:
    Err.Raise Err.Number, "GroupPageContentAsSingleSource", "Page " & CStr(pg.Index) & ": " & Err.Description
End Function

Private Function CountSelectablePages(ByVal doc As Document) As Long
    Dim pg As Page
    For Each pg In doc.Pages
        If GetPageSourceShapeCount(pg) > 0 Then CountSelectablePages = CountSelectablePages + 1
    Next pg
End Function

Private Function CountSelectablePagesInRange(ByVal doc As Document, ByVal firstPage As Long, ByVal lastPage As Long) As Long
    Dim i As Long

    NormalizeSourcePageRange doc, firstPage, lastPage
    For i = firstPage To lastPage
        If GetPageSourceShapeCount(doc.Pages(i)) > 0 Then CountSelectablePagesInRange = CountSelectablePagesInRange + 1
    Next i
End Function

Private Sub NormalizeSourcePageRange(ByVal doc As Document, ByRef firstPage As Long, ByRef lastPage As Long)
    If firstPage < 1 Then firstPage = 1
    If lastPage <= 0 Or lastPage > doc.Pages.Count Then lastPage = doc.Pages.Count
    If firstPage > doc.Pages.Count Then firstPage = doc.Pages.Count
    If lastPage < firstPage Then lastPage = firstPage
End Sub

Private Function PageHasShapes(ByVal pg As Page) As Boolean
    On Error Resume Next
    PageHasShapes = (pg.Shapes.Count > 0)
End Function

Private Function GetOpenDocumentByFullName(ByVal fullName As String) As Document
    Dim d As Document
    Dim wantName As String
    Dim fileOnly As String

    wantName = LCase$(Trim$(fullName))
    fileOnly = LCase$(Mid$(wantName, InStrRev(wantName, "\") + 1))

    For Each d In Documents
        If LCase$(Trim$(d.FullFileName)) = wantName Then
            Set GetOpenDocumentByFullName = d
            Exit Function
        End If
        If Len(fileOnly) > 0 Then
            If LCase$(Trim$(d.Title)) = fileOnly Then
                Set GetOpenDocumentByFullName = d
                Exit Function
            End If
        End If
    Next d
End Function

Private Sub CloseDocumentWithoutSaving(ByVal doc As Document)
    If doc Is Nothing Then Exit Sub

    On Error Resume Next
    CallByName doc, "Close", VbMethod, False
    If Err.Number <> 0 Then
        Err.Clear
        CallByName doc, "Close", VbMethod, 2
    End If
    If Err.Number <> 0 Then
        Err.Clear
        doc.Close
    End If
    On Error GoTo 0
End Sub

Private Sub NormalizeDocumentUnits(ByVal doc As Document)
    On Error Resume Next
    doc.Unit = cdrMillimeter
    doc.ReferencePoint = cdrCenter
End Sub

Private Sub SetCorelOptimization(ByVal enabled As Boolean)
    On Error Resume Next
    CallByName Application, "Optimization", VbLet, enabled
    If Not enabled Then CallByName Application, "Refresh", VbMethod
    On Error GoTo 0
End Sub

Private Function BuildTargetDocPath(ByVal srcDoc As Document, ByVal docNumber As Long, ByVal sideCode As String, Optional ByVal outputFolder As String = "") As String
    BuildTargetDocPath = BuildTargetDocPathByJobId(srcDoc, DEFAULT_LAYOUT_NAME, Format$(docNumber, "00"), sideCode, outputFolder)
End Function

Private Function BuildTargetDocPathByJobId(ByVal srcDoc As Document, ByVal layoutName As String, ByVal jobId As String, ByVal sideCode As String, Optional ByVal outputFolder As String = "") As String
    Dim folderName As String
    Dim outputProductName As String
    Dim safeJobId As String

    folderName = NormalizeOutputFolder(outputFolder, srcDoc)
    outputProductName = GetOutputProductName(layoutName)
    safeJobId = NormalizeJobId(jobId)
    BuildTargetDocPathByJobId = CombinePath(folderName, safeJobId & "_" & outputProductName & "_" & UCase$(NormalizeSideCode(sideCode)) & ".cdr")
End Function

Private Function NormalizeJobId(ByVal jobId As String) As String
    NormalizeJobId = SanitizeFileName(Trim$(jobId))
    If Len(NormalizeJobId) = 0 Then NormalizeJobId = DEFAULT_JOB_ID
End Function

Private Function GetOutputProductName(ByVal layoutName As String) As String
    Dim productName As String

    productName = NormalizeLayoutName(layoutName)
    If Left$(productName, 9) = "TEMPLATE_" Then productName = Mid$(productName, 10)
    If Right$(productName, Len(SLOT_LAYER_SUFFIX)) = SLOT_LAYER_SUFFIX Then
        productName = Trim$(Left$(productName, Len(productName) - Len(SLOT_LAYER_SUFFIX)))
    End If

    If LayoutNamesEquivalent(productName, "KARTA") Then
        GetOutputProductName = TARGET_DOC_PREFIX
    Else
        GetOutputProductName = LCase$(SanitizeFileName(productName))
    End If

    If Len(Trim$(GetOutputProductName)) = 0 Then GetOutputProductName = TARGET_DOC_PREFIX
End Function

Private Function GetSourceDocumentBaseName(ByVal srcDoc As Document) As String
    Dim fullName As String
    Dim fileName As String

    fullName = Trim$(srcDoc.FullFileName)
    If Len(fullName) > 0 Then
        fileName = Mid$(fullName, InStrRev(fullName, "\") + 1)
    Else
        fileName = srcDoc.Title
    End If

    GetSourceDocumentBaseName = SanitizeFileName(RemoveFileExtension(fileName))
    If Len(Trim$(GetSourceDocumentBaseName)) = 0 Then GetSourceDocumentBaseName = "mimaki_output"
End Function

Private Sub SaveOutputDocumentNoDialog(ByVal outDoc As Document, ByVal targetPath As String)
    Dim opt As New StructSaveAsOptions
    Dim folderPath As String

    If outDoc Is Nothing Then Err.Raise vbObjectError + 520, "SaveOutputDocumentNoDialog", "Output document is missing."
    targetPath = Trim$(targetPath)
    If Len(targetPath) = 0 Then Err.Raise vbObjectError + 520, "SaveOutputDocumentNoDialog", "Output document target path is empty."
    If LCase$(Right$(targetPath, 4)) <> ".cdr" Then targetPath = targetPath & ".cdr"

    folderPath = GetFolderFromFilePath(targetPath)
    If Len(folderPath) = 0 Then Err.Raise vbObjectError + 520, "SaveOutputDocumentNoDialog", "Output document target folder is empty."
    EnsureFolderExists folderPath

    opt.Filter = cdrCDR
    opt.Overwrite = True
    opt.Range = cdrAllPages
    opt.Version = cdrCurrentVersion

    outDoc.Activate
    outDoc.SaveAs targetPath, opt
End Sub

Private Function NormalizeOutputFolder(ByVal outputFolder As String, ByVal srcDoc As Document) As String
    NormalizeOutputFolder = GetSourceFolder(srcDoc)
    If Len(Trim$(NormalizeOutputFolder)) = 0 Then NormalizeOutputFolder = Trim$(outputFolder)
    If Len(Trim$(NormalizeOutputFolder)) = 0 Then NormalizeOutputFolder = MimakiV21_GetDefaultOutputRoot()
End Function

Private Function GetSourceFolder(ByVal srcDoc As Document) As String
    Dim p As String
    Dim slashPos As Long

    p = Trim$(srcDoc.FullFileName)
    If Len(p) > 0 Then
        slashPos = InStrRev(p, "\")
        If slashPos > 1 Then GetSourceFolder = Left$(p, slashPos - 1)
    End If
End Function

Private Function GetFolderFromFilePath(ByVal filePath As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(filePath, "\")
    If slashPos > 1 Then GetFolderFromFilePath = Left$(filePath, slashPos - 1)
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir$(filePath)) > 0)
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(Trim$(folderPath)) = 0 Then Exit Sub
    If FolderExists(folderPath) Then Exit Sub
    MkDir folderPath
End Sub

Private Function LongArrayHasItems(ByRef values() As Long) As Boolean
    On Error GoTo EH
    LongArrayHasItems = (UBound(values) >= LBound(values))
    Exit Function
EH:
    LongArrayHasItems = False
End Function

Private Function BuildRepeatedPageList(ByVal pageNumber As Long, ByVal itemCount As Long) As Long()
    Dim result() As Long
    Dim i As Long

    If itemCount <= 0 Then Err.Raise vbObjectError + 500, "BuildRepeatedPageList", "Page list is empty."
    ReDim result(1 To itemCount)
    For i = 1 To itemCount
        result(i) = pageNumber
    Next i
    BuildRepeatedPageList = result
End Function

Private Function BuildSequentialPageList(ByVal firstPage As Long, ByVal lastPage As Long) As Long()
    Dim result() As Long
    Dim i As Long
    Dim n As Long

    If lastPage < firstPage Then Err.Raise vbObjectError + 500, "BuildSequentialPageList", "Page list is empty."
    ReDim result(1 To (lastPage - firstPage + 1))
    For i = firstPage To lastPage
        n = n + 1
        result(n) = i
    Next i
    BuildSequentialPageList = result
End Function

Private Function BuildAlternatingPageList(ByVal firstPage As Long, ByVal lastPage As Long, ByVal oddPositions As Boolean) As Long()
    Dim result() As Long
    Dim i As Long
    Dim n As Long
    Dim pos As Long
    Dim countItems As Long

    For i = firstPage To lastPage
        pos = i - firstPage + 1
        If ((pos Mod 2) = 1) = oddPositions Then countItems = countItems + 1
    Next i
    If countItems <= 0 Then Err.Raise vbObjectError + 500, "BuildAlternatingPageList", "Page list is empty."

    ReDim result(1 To countItems)
    For i = firstPage To lastPage
        pos = i - firstPage + 1
        If ((pos Mod 2) = 1) = oddPositions Then
            n = n + 1
            result(n) = i
        End If
    Next i
    BuildAlternatingPageList = result
End Function
Private Function GetFileNameOnly(ByVal filePath As String) As String
    Dim p As Long
    p = InStrRev(filePath, "\")
    If p > 0 Then
        GetFileNameOnly = Mid$(filePath, p + 1)
    Else
        GetFileNameOnly = filePath
    End If
End Function

Private Function BuildAvailableFilePath(ByVal filePath As String) As String
    Dim basePath As String
    Dim ext As String
    Dim candidate As String
    Dim counter As Long

    If Not FileExists(filePath) Then
        BuildAvailableFilePath = filePath
        Exit Function
    End If

    basePath = RemoveFileExtension(filePath)
    ext = Mid$(filePath, Len(basePath) + 1)

    counter = 1
    Do
        candidate = basePath & "_" & Format$(counter, "00") & ext
        counter = counter + 1
    Loop While FileExists(candidate)

    BuildAvailableFilePath = candidate
End Function

Private Function CombinePath(ByVal folderPath As String, ByVal fileName As String) As String
    If Right$(folderPath, 1) = "\" Then
        CombinePath = folderPath & fileName
    Else
        CombinePath = folderPath & "\" & fileName
    End If
End Function

Private Function NormalizeLayoutName(ByVal rawName As String) As String
    NormalizeLayoutName = UCase$(Trim$(rawName))
End Function

Private Function LayoutNamesEquivalent(ByVal leftName As String, ByVal rightName As String) As Boolean
    Dim a As String
    Dim b As String

    a = NormalizeLayoutName(leftName)
    b = NormalizeLayoutName(rightName)

    If a = b Then
        LayoutNamesEquivalent = True
    ElseIf (a = "KARTA" And b = "KARTY") Or (a = "KARTY" And b = "KARTA") Then
        LayoutNamesEquivalent = True
    End If
End Function

Private Function NormalizeSideCode(ByVal rawCode As String) As String
    Dim s As String
    s = UCase$(Trim$(rawCode))
    If Left$(s, 1) = "Z" Then
        NormalizeSideCode = "Z"
    Else
        NormalizeSideCode = "P"
    End If
End Function

Private Function CeilingDiv(ByVal dividend As Long, ByVal divisor As Long) As Long
    If divisor <= 0 Then
        CeilingDiv = 0
    Else
        CeilingDiv = (dividend + divisor - 1) \ divisor
    End If
End Function

Private Sub GetSourceCanvasSizeForSlot(ByRef slot As TSlotInfo, ByVal inputOrientation As Long, ByRef canvasWidth As Double, ByRef canvasHeight As Double)
    Dim longSide As Double
    Dim shortSide As Double
    Dim slotWidth As Double
    Dim slotHeight As Double
    Dim bleedTotal As Double

    slotWidth = Abs(slot.RightX - slot.LeftX)
    slotHeight = Abs(slot.TopY - slot.BottomY)
    bleedTotal = SLOT_BLEED_MM * 2#

    longSide = MaxDouble(slotWidth, slotHeight) + bleedTotal
    shortSide = MinDouble(slotWidth, slotHeight) + bleedTotal

    Select Case inputOrientation
        Case mki21IoLandscape
            canvasWidth = longSide
            canvasHeight = shortSide
        Case mki21IoPortrait
            canvasWidth = shortSide
            canvasHeight = longSide
        Case Else
            If slotWidth >= slotHeight Then
                canvasWidth = longSide
                canvasHeight = shortSide
            Else
                canvasWidth = shortSide
                canvasHeight = longSide
            End If
    End Select
End Sub

Private Function SlotLooksLikePvcCard(ByVal slotWidth As Double, ByVal slotHeight As Double) As Boolean
    Dim longSide As Double
    Dim shortSide As Double

    longSide = MaxDouble(slotWidth, slotHeight)
    shortSide = MinDouble(slotWidth, slotHeight)

    SlotLooksLikePvcCard = (Abs(longSide - 86#) <= CARD_SLOT_MATCH_TOLERANCE_MM And Abs(shortSide - 55#) <= CARD_SLOT_MATCH_TOLERANCE_MM)
End Function

Private Function ResolveEffectiveInputOrientation(ByVal requestedOrientation As Long, ByVal srcRange As ShapeRange) As Long
    Const SOURCE_ORIENTATION_TOLERANCE_MM As Double = 3#
    Dim sourceLeft As Double
    Dim sourceRight As Double
    Dim sourceTop As Double
    Dim sourceBottom As Double
    Dim sourceWidth As Double
    Dim sourceHeight As Double

    ResolveEffectiveInputOrientation = requestedOrientation
    If ResolveEffectiveInputOrientation <> mki21IoLandscape And ResolveEffectiveInputOrientation <> mki21IoPortrait Then
        ResolveEffectiveInputOrientation = mki21IoAuto
    End If

    If Not GetSourceRangeBounds(srcRange, sourceLeft, sourceRight, sourceTop, sourceBottom) Then Exit Function

    sourceWidth = Abs(sourceRight - sourceLeft)
    sourceHeight = Abs(sourceTop - sourceBottom)
    If sourceWidth <= 0# Or sourceHeight <= 0# Then Exit Function

    If sourceWidth > sourceHeight + SOURCE_ORIENTATION_TOLERANCE_MM Then
        ResolveEffectiveInputOrientation = mki21IoLandscape
    ElseIf sourceHeight > sourceWidth + SOURCE_ORIENTATION_TOLERANCE_MM Then
        ResolveEffectiveInputOrientation = mki21IoPortrait
    End If
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then
        MaxDouble = a
    Else
        MaxDouble = b
    End If
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a <= b Then
        MinDouble = a
    Else
        MinDouble = b
    End If
End Function

Private Sub GetSourceCropCenter(ByVal pg As Page, ByVal srcRange As ShapeRange, ByVal canvasWidth As Double, ByVal canvasHeight As Double, ByRef centerX As Double, ByRef centerY As Double)
    Dim pageLeft As Double
    Dim pageRight As Double
    Dim pageTop As Double
    Dim pageBottom As Double
    Dim pageWidth As Double
    Dim pageHeight As Double
    Dim sourceLeft As Double
    Dim sourceRight As Double
    Dim sourceTop As Double
    Dim sourceBottom As Double

    GetPageBounds pg, pageLeft, pageRight, pageTop, pageBottom
    centerX = (pageLeft + pageRight) / 2#
    centerY = (pageTop + pageBottom) / 2#
    pageWidth = Abs(pageRight - pageLeft)
    pageHeight = Abs(pageTop - pageBottom)

    If SourcePageMatchesCropCanvas(pageWidth, pageHeight, canvasWidth, canvasHeight) Then Exit Sub

    If GetSourceRangeBounds(srcRange, sourceLeft, sourceRight, sourceTop, sourceBottom) Then
        centerX = (sourceLeft + sourceRight) / 2#
        centerY = (sourceTop + sourceBottom) / 2#
    End If
End Sub

Private Function SourcePageMatchesCropCanvas(ByVal pageWidth As Double, ByVal pageHeight As Double, ByVal canvasWidth As Double, ByVal canvasHeight As Double) As Boolean
    Const PAGE_CANVAS_TOLERANCE_MM As Double = 15#

    SourcePageMatchesCropCanvas = (Abs(MaxDouble(pageWidth, pageHeight) - MaxDouble(canvasWidth, canvasHeight)) <= PAGE_CANVAS_TOLERANCE_MM And _
                                   Abs(MinDouble(pageWidth, pageHeight) - MinDouble(canvasWidth, canvasHeight)) <= PAGE_CANVAS_TOLERANCE_MM)
End Function

Private Function GetSourceRangeBounds(ByVal srcRange As ShapeRange, ByRef sourceLeft As Double, ByRef sourceRight As Double, ByRef sourceTop As Double, ByRef sourceBottom As Double) As Boolean
    Dim sh As Shape

    If srcRange Is Nothing Then Exit Function
    If srcRange.Count <= 0 Then Exit Function

    For Each sh In srcRange.Shapes
        If Not sh Is Nothing Then
            If Not GetSourceRangeBounds Then
                sourceLeft = sh.LeftX
                sourceRight = sh.RightX
                sourceTop = sh.TopY
                sourceBottom = sh.BottomY
                GetSourceRangeBounds = True
            Else
                If sh.LeftX < sourceLeft Then sourceLeft = sh.LeftX
                If sh.RightX > sourceRight Then sourceRight = sh.RightX
                If sh.TopY > sourceTop Then sourceTop = sh.TopY
                If sh.BottomY < sourceBottom Then sourceBottom = sh.BottomY
            End If
        End If
    Next sh
End Function

Private Sub CenterShapeOnPoint(ByVal sh As Shape, ByVal targetX As Double, ByVal targetY As Double)
    Dim curX As Double
    Dim curY As Double
    Dim dx As Double
    Dim dy As Double

    curX = (sh.LeftX + sh.RightX) / 2#
    curY = (sh.TopY + sh.BottomY) / 2#
    dx = targetX - curX
    dy = targetY - curY
    sh.Move dx, dy
End Sub

Private Sub GetSourceRangeNormalizedOffset(ByVal pg As Page, ByVal srcRange As ShapeRange, ByRef offsetX As Double, ByRef offsetY As Double)
    Dim sh As Shape
    Dim sourceLeft As Double
    Dim sourceRight As Double
    Dim sourceTop As Double
    Dim sourceBottom As Double
    Dim sourceCenterX As Double
    Dim sourceCenterY As Double
    Dim pageLeft As Double
    Dim pageRight As Double
    Dim pageTop As Double
    Dim pageBottom As Double
    Dim pageCenterX As Double
    Dim pageCenterY As Double
    Dim pageWidth As Double
    Dim pageHeight As Double
    Dim hasBounds As Boolean

    offsetX = 0#
    offsetY = 0#

    On Error Resume Next
    If srcRange Is Nothing Then Exit Sub
    If srcRange.Count <= 0 Then Exit Sub

    For Each sh In srcRange.Shapes
        If Not sh Is Nothing Then
            If Not hasBounds Then
                sourceLeft = sh.LeftX
                sourceRight = sh.RightX
                sourceTop = sh.TopY
                sourceBottom = sh.BottomY
                hasBounds = True
            Else
                If sh.LeftX < sourceLeft Then sourceLeft = sh.LeftX
                If sh.RightX > sourceRight Then sourceRight = sh.RightX
                If sh.TopY > sourceTop Then sourceTop = sh.TopY
                If sh.BottomY < sourceBottom Then sourceBottom = sh.BottomY
            End If
        End If
    Next sh
    If Not hasBounds Then Exit Sub

    sourceCenterX = (sourceLeft + sourceRight) / 2#
    sourceCenterY = (sourceTop + sourceBottom) / 2#
    GetPageBounds pg, pageLeft, pageRight, pageTop, pageBottom
    pageCenterX = (pageLeft + pageRight) / 2#
    pageCenterY = (pageTop + pageBottom) / 2#
    pageWidth = Abs(pageRight - pageLeft)
    pageHeight = Abs(pageTop - pageBottom)
    If pageWidth <= 0# Or pageHeight <= 0# Then Exit Sub

    offsetX = (sourceCenterX - pageCenterX) / pageWidth
    offsetY = (sourceCenterY - pageCenterY) / pageHeight
    On Error GoTo 0
End Sub

Private Sub GetPageBounds(ByVal pg As Page, ByRef leftX As Double, ByRef rightX As Double, ByRef topY As Double, ByRef bottomY As Double)
    On Error Resume Next
    leftX = pg.LeftX
    rightX = pg.RightX
    topY = pg.TopY
    bottomY = pg.BottomY
    If Err.Number <> 0 Then
        Err.Clear
        leftX = -pg.SizeWidth / 2#
        rightX = pg.SizeWidth / 2#
        topY = pg.SizeHeight / 2#
        bottomY = -pg.SizeHeight / 2#
    End If
    On Error GoTo 0
End Sub

Private Function GetNextCardSetupOrientation() As Long
    Dim lastApplied As Long

    lastApplied = CLng(Val(GetSetting(CONFIG_APP, CONFIG_SECTION, "CardSetupLastApplied", "0")))
    If lastApplied = mki21IoPortrait Then
        GetNextCardSetupOrientation = mki21IoLandscape
    Else
        GetNextCardSetupOrientation = mki21IoPortrait
    End If
End Function

Private Sub EnsureCardBoundaryGuide(ByVal pg As Page)
    Dim guideShape As Shape
    Dim leftX As Double
    Dim rightX As Double
    Dim topY As Double
    Dim bottomY As Double

    If pg Is Nothing Then Exit Sub

    pg.Activate
    ClearCardBoundaryGuides pg
    GetPageBounds pg, leftX, rightX, topY, bottomY
    Set guideShape = ActivePage.GuidesLayer.CreateGuideAngle(leftX, 0#, 90#)
    If Not guideShape Is Nothing Then guideShape.Name = CARD_GUIDE_SHAPE_NAME & "_L"
    Set guideShape = ActivePage.GuidesLayer.CreateGuideAngle(rightX, 0#, 90#)
    If Not guideShape Is Nothing Then guideShape.Name = CARD_GUIDE_SHAPE_NAME & "_R"
    Set guideShape = ActivePage.GuidesLayer.CreateGuideAngle(0#, topY, 0#)
    If Not guideShape Is Nothing Then guideShape.Name = CARD_GUIDE_SHAPE_NAME & "_T"
    Set guideShape = ActivePage.GuidesLayer.CreateGuideAngle(0#, bottomY, 0#)
    If Not guideShape Is Nothing Then guideShape.Name = CARD_GUIDE_SHAPE_NAME & "_B"
End Sub

Private Sub ClearCardBoundaryGuides(ByVal pg As Page)
    Dim sh As Shape
    Dim i As Long

    If pg Is Nothing Then Exit Sub

    On Error Resume Next
    pg.Activate
    ActivePage.GuidesLayer.Editable = True
    For i = ActivePage.Guides(cdrAllGuides).Count To 1 Step -1
        Set sh = ActivePage.Guides(cdrAllGuides)(i)
        If Not sh Is Nothing Then sh.Delete
    Next i
    ActivePage.Guides.All.Delete
    ActivePage.Guides(cdrAllGuides).All.Delete
    ActivePage.GuidesLayer.Shapes.All.Delete
    ActiveDocument.MasterPage.GuidesLayer.Editable = True
    ActiveDocument.MasterPage.Guides.All.Delete
    ActiveDocument.MasterPage.Guides(cdrAllGuides).All.Delete
    ActiveDocument.MasterPage.GuidesLayer.Shapes.All.Delete
    On Error GoTo 0
End Sub

Private Sub RotateOffset(ByRef offsetX As Double, ByRef offsetY As Double, ByVal rotationAngle As Double)
    Dim oldX As Double
    Dim oldY As Double
    Dim normalizedAngle As Long

    oldX = offsetX
    oldY = offsetY
    normalizedAngle = CLng(rotationAngle)

    Select Case normalizedAngle
        Case 90, -270
            offsetX = -oldY
            offsetY = oldX
        Case -90, 270
            offsetX = oldY
            offsetY = -oldX
        Case 180, -180
            offsetX = -oldX
            offsetY = -oldY
    End Select
End Sub

Private Function RotateShapeToSlotOrientation(ByVal sh As Shape, ByRef slot As TSlotInfo, ByVal inputOrientation As Long) As Double
    Dim shapeWidth As Double
    Dim shapeHeight As Double
    Dim slotWidth As Double
    Dim slotHeight As Double
    Dim shapeIsLandscape As Boolean
    Dim slotIsLandscape As Boolean
    Dim tolerance As Double

    tolerance = 0.2

    shapeWidth = Abs(sh.RightX - sh.LeftX)
    shapeHeight = Abs(sh.TopY - sh.BottomY)
    slotWidth = Abs(slot.RightX - slot.LeftX)
    slotHeight = Abs(slot.TopY - slot.BottomY)

    If shapeWidth <= 0 Or shapeHeight <= 0 Then Exit Function
    If slotWidth <= 0 Or slotHeight <= 0 Then Exit Function

    If Abs(slotWidth - slotHeight) <= tolerance Then Exit Function

    If inputOrientation = mki21IoLandscape Then
        shapeIsLandscape = True
    ElseIf inputOrientation = mki21IoPortrait Then
        shapeIsLandscape = False
    Else
        If Abs(shapeWidth - shapeHeight) <= tolerance Then Exit Function
        shapeIsLandscape = (shapeWidth > shapeHeight)
    End If

    slotIsLandscape = (slotWidth > slotHeight)

    If shapeIsLandscape <> slotIsLandscape Then
        sh.Rotate 90#
        RotateShapeToSlotOrientation = 90#
    End If
End Function

' ============================================================
' Prompt helpers
' ============================================================

Private Function PromptText(ByVal prompt As String, ByVal defaultValue As String, ByVal title As String) As String
    PromptText = Trim$(InputBox(prompt, title, defaultValue))
End Function

Private Function PromptLong(ByVal prompt As String, ByVal defaultValue As Long, ByVal title As String) As Long
    PromptLong = CLng(Val(InputBox(prompt, title, CStr(defaultValue))))
End Function

Private Function PromptTargetNumber(ByVal defaultValue As Long) As Long
    PromptTargetNumber = PromptLong("Enter document number XX.", defaultValue, "Document number")
End Function

Private Function PromptOpenExisting(ByRef docNumber As Long, ByVal sideCode As String, ByVal srcDoc As Document) As Long
    Dim targetPath As String
    Dim resp As VbMsgBoxResult
    Dim newNumber As Long

    Do
        targetPath = BuildTargetDocPath(srcDoc, docNumber, sideCode)
        If Not FileExists(targetPath) Then
            PromptOpenExisting = 0
            Exit Function
        End If

        resp = MsgBox("Target file exists:" & vbCrLf & targetPath & vbCrLf & vbCrLf & _
                      "Yes = open existing file" & vbCrLf & _
                      "No = choose another number" & vbCrLf & _
                      "Cancel = stop", _
                      vbYesNoCancel + vbQuestion, "Target document")

        If resp = vbYes Then
            PromptOpenExisting = 1
            Exit Function
        ElseIf resp = vbNo Then
            newNumber = PromptTargetNumber(docNumber + 1)
            If newNumber <= 0 Then
                PromptOpenExisting = -1
                Exit Function
            End If
            docNumber = newNumber
        Else
            PromptOpenExisting = -1
            Exit Function
        End If
    Loop
End Function

Private Function AskSourceMode() As Long
    Dim resp As VbMsgBoxResult

    resp = MsgBox("Yes = copy one graphic from the current page into multiple positions." & vbCrLf & _
                  "No = place the content from all pages one by one.", _
                  vbYesNoCancel + vbQuestion, "Source mode")
    If resp = vbCancel Then
        AskSourceMode = 0
    ElseIf resp = vbYes Then
        AskSourceMode = mki21SmCopiesFromCurrentPage
    Else
        AskSourceMode = mki21SmAllPagesAsItems
    End If
End Function

' ============================================================
' Export helpers kept for later manual use
' ============================================================

Private Sub SaveMimakiProductionFiles(ByVal outDoc As Document, ByRef cdrPath As String, ByRef pdfPath As String)
    Dim basePath As String

    On Error GoTo EH

    If Len(Trim$(outDoc.FullFileName)) = 0 Then
        Err.Raise vbObjectError + 520, "SaveMimakiProductionFiles", "Output document has no CDR file name yet. Run imposition first, then use Save."
    End If

    cdrPath = outDoc.FullFileName
    basePath = RemoveFileExtension(cdrPath)
    pdfPath = BuildSinglePagePdfPath(basePath & ".pdf", 1)

    outDoc.Save
    ExportProductionPDF outDoc, basePath & ".pdf"
    Exit Sub

EH:
    Err.Raise Err.Number, "SaveMimakiProductionFiles", Err.Description
End Sub

Private Sub ExportMimakiProductionPDFOnly(ByVal outDoc As Document, ByVal outputFolder As String, ByRef pdfPath As String)
    Dim baseName As String

    On Error GoTo EH

    outputFolder = Trim$(outputFolder)
    If Len(outputFolder) = 0 Then
        Err.Raise vbObjectError + 530, "ExportMimakiProductionPDFOnly", "Output folder is empty."
    End If

    baseName = GetDocumentBaseName(outDoc)
    pdfPath = BuildSinglePagePdfPath(CombinePath(outputFolder, baseName & ".pdf"), 1)

    ExportProductionPDF outDoc, CombinePath(outputFolder, baseName & ".pdf")
    Exit Sub

EH:
    Err.Raise Err.Number, "ExportMimakiProductionPDFOnly", Err.Description
End Sub

Private Sub CopyProductionPDFSetToFolder(ByVal basePdfPath As String, ByVal pageCount As Long, ByVal targetFolder As String, ByRef firstCopiedPath As String)
    Dim sourcePath As String
    Dim targetPath As String
    Dim whiteSourcePath As String
    Dim whiteTargetPath As String
    Dim pageIndex As Long

    targetFolder = Trim$(targetFolder)
    If Len(targetFolder) = 0 Then
        Err.Raise vbObjectError + 531, "CopyProductionPDFSetToFolder", "RasterLink hotfolder path is empty."
    End If

    If Not FolderExists(targetFolder) Then
        Err.Raise vbObjectError + 532, "CopyProductionPDFSetToFolder", "RasterLink hotfolder is not available: " & targetFolder
    End If

    If pageCount <= 1 Then
        sourcePath = BuildSinglePagePdfPath(basePdfPath, 1)
        targetPath = CombinePath(targetFolder, GetFileNameOnly(sourcePath))
        targetPath = BuildAvailableFilePath(targetPath)
        CopyOneProductionPDF sourcePath, targetPath
        firstCopiedPath = targetPath
        whiteSourcePath = BuildWhiteSinglePagePdfPath(basePdfPath, 1)
        If FileExists(whiteSourcePath) Then
            whiteTargetPath = CombinePath(targetFolder, GetFileNameOnly(whiteSourcePath))
            whiteTargetPath = BuildAvailableFilePath(whiteTargetPath)
            CopyOneProductionPDF whiteSourcePath, whiteTargetPath
        End If
    Else
        For pageIndex = 1 To pageCount
            sourcePath = BuildSinglePagePdfPath(basePdfPath, pageIndex)
            targetPath = CombinePath(targetFolder, GetFileNameOnly(sourcePath))
            targetPath = BuildAvailableFilePath(targetPath)
            CopyOneProductionPDF sourcePath, targetPath
            If pageIndex = 1 Then firstCopiedPath = targetPath
            whiteSourcePath = BuildWhiteSinglePagePdfPath(basePdfPath, pageIndex)
            If FileExists(whiteSourcePath) Then
                whiteTargetPath = CombinePath(targetFolder, GetFileNameOnly(whiteSourcePath))
                whiteTargetPath = BuildAvailableFilePath(whiteTargetPath)
                CopyOneProductionPDF whiteSourcePath, whiteTargetPath
            End If
        Next pageIndex
    End If
End Sub

Private Sub CopyOneProductionPDF(ByVal sourcePath As String, ByVal targetPath As String)
    If Not FileExists(sourcePath) Then
        Err.Raise vbObjectError + 533, "CopyOneProductionPDF", "Production PDF was not created: " & sourcePath
    End If
    FileCopy sourcePath, targetPath
End Sub

Private Function ResolveMimakiOutputDocument() As Document
    Dim activeDoc As Document

    If Documents.Count > 0 Then
        Set activeDoc = ActiveDocument
        If Not activeDoc Is Nothing Then
            If DocumentLooksLikeMimakiOutput(activeDoc) Then
                Set ResolveMimakiOutputDocument = activeDoc
                Exit Function
            End If
        End If
    End If

    If Not gMimakiV21LastOutputDoc Is Nothing Then
        Set ResolveMimakiOutputDocument = gMimakiV21LastOutputDoc
    End If
End Function

Private Function DocumentLooksLikeMimakiOutput(ByVal doc As Document) As Boolean
    Dim pg As Page
    Dim artLayer As Layer
    Dim whiteLayer As Layer
    Dim titleText As String

    On Error Resume Next

    titleText = UCase$(Trim$(doc.Title))
    If InStr(1, titleText, UCase$(TARGET_DOC_PREFIX), vbTextCompare) > 0 Then
        DocumentLooksLikeMimakiOutput = True
        Exit Function
    End If

    For Each pg In doc.Pages
        Set artLayer = FindLayerByName(pg, PRINT_DATA_LAYER_NAME)
        If Not artLayer Is Nothing Then
            DocumentLooksLikeMimakiOutput = True
            Exit Function
        End If
        Set whiteLayer = FindLayerByName(pg, WHITE_LAYER_NAME)
        If Not whiteLayer Is Nothing Then
            DocumentLooksLikeMimakiOutput = True
            Exit Function
        End If
    Next pg
End Function

Private Sub CaptureLayerStatesForProductionExport(ByVal doc As Document, ByVal targetLayerName As String, ByRef layers As Collection, ByRef visibleStates As Collection, ByRef printableStates As Collection)
    Dim pg As Page
    Dim ly As Layer
    Dim keepLayer As Boolean

    For Each pg In doc.Pages
        For Each ly In pg.Layers
            layers.Add ly
            visibleStates.Add CBool(ly.Visible)
            printableStates.Add CBool(ly.Printable)

            keepLayer = LayerNameMatches(ly.Name, targetLayerName)
            ly.Visible = keepLayer
            ly.Printable = keepLayer
        Next ly
    Next pg
End Sub

Private Sub RestoreLayerStates(ByRef layers As Collection, ByRef visibleStates As Collection, ByRef printableStates As Collection)
    Dim i As Long
    Dim ly As Layer

    On Error Resume Next
    For i = 1 To layers.Count
        Set ly = layers(i)
        ly.Visible = CBool(visibleStates(i))
        ly.Printable = CBool(printableStates(i))
    Next i
    On Error GoTo 0
End Sub

Private Sub ExportProductionPDF(ByVal outDoc As Document, ByVal pdfPath As String)
    Dim pageIndex As Long
    Dim onePagePath As String
    Dim whitePagePath As String
    Dim layers As Collection
    Dim visibleStates As Collection
    Dim printableStates As Collection

    On Error GoTo EH

    outDoc.Activate
    outDoc.ClearSelection

    For pageIndex = 1 To outDoc.Pages.Count
        onePagePath = BuildSinglePagePdfPath(pdfPath, pageIndex)
        ExportOneProductionPagePDF outDoc, pageIndex, onePagePath, PRINT_DATA_LAYER_NAME, layers, visibleStates, printableStates
        If PageHasWhiteContent(outDoc.Pages(pageIndex)) Then
            whitePagePath = BuildWhiteSinglePagePdfPath(pdfPath, pageIndex)
            ExportOneProductionPagePDF outDoc, pageIndex, whitePagePath, WHITE_LAYER_NAME, layers, visibleStates, printableStates
        End If
    Next pageIndex
    Exit Sub

EH:
    On Error Resume Next
    If Not layers Is Nothing Then RestoreLayerStates layers, visibleStates, printableStates
    On Error GoTo 0
    Err.Raise Err.Number, "ExportProductionPDF", Err.Description
End Sub

Private Sub ExportOneProductionPagePDF(ByVal outDoc As Document, ByVal pageIndex As Long, ByVal pdfPath As String, ByVal targetLayerName As String, ByRef layers As Collection, ByRef visibleStates As Collection, ByRef printableStates As Collection)
    Set layers = New Collection
    Set visibleStates = New Collection
    Set printableStates = New Collection
    CaptureLayerStatesForProductionExport outDoc, targetLayerName, layers, visibleStates, printableStates
    outDoc.Pages(pageIndex).Activate
    ConfigurePDFForPrintCMYK outDoc, CStr(pageIndex)
    outDoc.ClearSelection
    outDoc.PublishToPDF pdfPath
    RestoreLayerStates layers, visibleStates, printableStates
    Set layers = Nothing
    Set visibleStates = Nothing
    Set printableStates = Nothing
End Sub

Private Function PageHasWhiteContent(ByVal pg As Page) As Boolean
    Dim whiteLayer As Layer

    Set whiteLayer = FindLayerByName(pg, WHITE_LAYER_NAME)
    If whiteLayer Is Nothing Then Exit Function
    PageHasWhiteContent = (whiteLayer.Shapes.Count > 0)
End Function

Private Sub ExportProductionEPS(ByVal outDoc As Document, ByVal epsPath As String)
    Dim ex As ExportFilter

    Set ex = outDoc.ExportEx(epsPath, cdrEPS, cdrAllPages)
    ConfigureExportFilterForPrintCMYK ex
    ex.Finish
End Sub

Private Sub ConfigurePDFForPrintCMYK(ByVal outDoc As Document, Optional ByVal pageRange As String = "")
    Dim pdfSettings As Object

    On Error Resume Next
    outDoc.ClearSelection
    Set pdfSettings = CallByName(outDoc, "PDFSettings", VbGet)
    If pdfSettings Is Nothing Then Exit Sub

    CallByName pdfSettings, "PublishRange", VbLet, pdfPageRange
    If Len(Trim$(pageRange)) > 0 Then
        CallByName pdfSettings, "PageRange", VbLet, pageRange
    Else
        CallByName pdfSettings, "PageRange", VbLet, "1-" & CStr(outDoc.Pages.Count)
    End If
    CallByName pdfSettings, "SelectionOnly", VbLet, False
    CallByName pdfSettings, "UseColorProfile", VbLet, True
    CallByName pdfSettings, "PreserveSpotColors", VbLet, True
    CallByName pdfSettings, "EmbedFonts", VbLet, True
    On Error GoTo 0
End Sub

Private Sub ConfigureExportFilterForPrintCMYK(ByVal ex As ExportFilter)
    On Error Resume Next
    CallByName ex, "ColorMode", VbLet, 1
    CallByName ex, "UseColorProfile", VbLet, True
    CallByName ex, "EmbedFonts", VbLet, True
    On Error GoTo 0
End Sub

Private Sub MimakiPublishCurrentProductionPDF()
    Dim outDoc As Document
    Dim pdfPath As String
    Dim outputFolder As String

    On Error GoTo EH

    Set outDoc = GetCurrentMimakiOutputDocumentV21()
    If outDoc Is Nothing Then
        MsgBox "Najprv vytvor vystupny dokument.", vbExclamation, "Mimaki PDF"
        Exit Sub
    End If

    outDoc.Activate
    outputFolder = GetDocumentFolderOrFallback(outDoc)
    ExportMimakiProductionPDFOnly outDoc, outputFolder, pdfPath
    Exit Sub

EH:
    MsgBox FormatMimakiV21ErrorForUser(Err.Number, Err.Description), vbCritical, "Mimaki PDF"
End Sub

Private Sub MimakiShowExportDialog(ByVal filter As cdrFilter, ByVal ext As String)
    Dim outDoc As Document
    Dim ex As ExportFilter
    Dim defaultPath As String

    Set outDoc = GetCurrentMimakiOutputDocumentV21()
    If outDoc Is Nothing Then
        MsgBox "Najprv vytvor vystupny dokument.", vbExclamation, "Mimaki rozklad"
        Exit Sub
    End If

    outDoc.Activate
    defaultPath = BuildDefaultExportPath(outDoc, ext)
    Set ex = outDoc.ExportEx(defaultPath, filter, cdrAllPages)
    If ex.HasDialog Then
        If ex.ShowDialog(0) Then
            ex.Finish
        End If
    Else
        ex.Finish
    End If
End Sub

Private Function RemoveFileExtension(ByVal filePath As String) As String
    Dim p As Long

    p = InStrRev(filePath, ".")
    If p > 0 Then
        RemoveFileExtension = Left$(filePath, p - 1)
    Else
        RemoveFileExtension = filePath
    End If
End Function

Private Function BuildSinglePagePdfPath(ByVal basePdfPath As String, ByVal pageIndex As Long) As String
    BuildSinglePagePdfPath = RemoveFileExtension(basePdfPath) & "_p" & Format$(pageIndex, "00") & ".pdf"
End Function

Private Function BuildWhiteSinglePagePdfPath(ByVal basePdfPath As String, ByVal pageIndex As Long) As String
    BuildWhiteSinglePagePdfPath = RemoveFileExtension(basePdfPath) & "_p" & Format$(pageIndex, "00") & "_W.pdf"
End Function

Private Function GetDocumentBaseName(ByVal doc As Document) As String
    Dim fullName As String
    Dim fileName As String

    fullName = Trim$(doc.FullFileName)
    If Len(fullName) > 0 Then
        fileName = Mid$(fullName, InStrRev(fullName, "\") + 1)
    Else
        fileName = doc.Title
    End If

    GetDocumentBaseName = SanitizeFileName(RemoveFileExtension(fileName))
    If Len(Trim$(GetDocumentBaseName)) = 0 Then GetDocumentBaseName = "mimaki_output"
End Function

Private Function BuildDefaultExportPath(ByVal doc As Document, ByVal ext As String) As String
    Dim baseName As String
    Dim folderName As String

    baseName = SanitizeFileName(doc.Title)
    If Len(Trim$(baseName)) = 0 Then baseName = "mimaki_output"

    folderName = Environ$("TEMP")
    If Len(Trim$(folderName)) = 0 Then folderName = "C:\Temp"

    BuildDefaultExportPath = CombinePath(folderName, baseName & ext)
End Function

Private Function GetDocumentFolderOrFallback(ByVal doc As Document) As String
    Dim fullName As String
    Dim p As Long

    On Error Resume Next
    fullName = Trim$(doc.FullFileName)
    On Error GoTo 0

    p = InStrRev(fullName, "\")
    If p > 1 Then
        GetDocumentFolderOrFallback = Left$(fullName, p - 1)
        Exit Function
    End If

    GetDocumentFolderOrFallback = Environ$("TEMP")
    If Len(Trim$(GetDocumentFolderOrFallback)) = 0 Then GetDocumentFolderOrFallback = "C:\Temp"
End Function

Private Function SanitizeFileName(ByVal rawName As String) As String
    Dim badChars As Variant
    Dim i As Long
    Dim s As String

    s = Trim$(rawName)
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i
    SanitizeFileName = s
End Function





