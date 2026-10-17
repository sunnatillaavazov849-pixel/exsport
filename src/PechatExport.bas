Attribute VB_Name = "PechatExport"
'==============================================================================
'  Pechat Eksport - CorelDRAW uchun makros (plagin)
'
'  Rasmni pechatga (banner, orakal, bekprint) eksport qiladi:
'    - rang rejimi: CMYK yoki RGB
'    - CMYK da rang to'qlashmasligi uchun ranglarni ochroq qilish
'    - format: TIF, PNG, JPG, JPEG
'    - razreshenie (dpi)
'    - fayl nomi avtomatik:  "150 x 200 cm_banner_cmyk.tif"
'
'  Ishga tushirish: PechatgaEksport makrosini tugmaga biriktiring.
'==============================================================================
Option Explicit

Public Const APP_NAME As String = "CorelPechatExport"

' Asosiy makros - shu makrosni panelga tugma qilib qo'ying
Public Sub PechatgaEksport()
    If Documents.Count = 0 Then
        MsgBox "Avval CorelDRAW hujjatini oching.", vbExclamation, "Pechat eksport"
        Exit Sub
    End If
    frmExport.Show
End Sub

' Eksport qilinadigan maydon chegarasi (santimetrda)
Public Function GetExportBounds(ByVal onlySelection As Boolean, _
        x As Double, y As Double, w As Double, h As Double) As Boolean
    Dim doc As Document, oldUnit As cdrUnit
    Set doc = ActiveDocument
    oldUnit = doc.Unit
    doc.Unit = cdrCentimeter

    If onlySelection Then
        If ActiveSelectionRange.Count > 0 Then
            ActiveSelectionRange.GetBoundingBox x, y, w, h, True
            GetExportBounds = True
        End If
    Else
        x = ActivePage.LeftX
        y = ActivePage.BottomY
        w = ActivePage.SizeWidth
        h = ActivePage.SizeHeight
        GetExportBounds = True
    End If

    doc.Unit = oldUnit
End Function

' 150 -> "150", 150.25 -> "150.3"
Public Function FormatCm(ByVal v As Double) As String
    v = Round(v, 1)
    If v = Int(v) Then
        FormatCm = CStr(CLng(v))
    Else
        FormatCm = Replace(CStr(v), ",", ".")
    End If
End Function

' "150 x 200 cm_banner_cmyk.tif"
Public Function BuildFileName(ByVal w As Double, ByVal h As Double, ByVal printType As String, _
        ByVal isCMYK As Boolean, ByVal ext As String) As String
    BuildFileName = FormatCm(w) & " x " & FormatCm(h) & " cm_" & LCase$(printType) & "_" & _
                    IIf(isCMYK, "cmyk", "rgb") & "." & LCase$(ext)
End Function

' Eksport. Muvaffaqiyatli bo'lsa saqlangan fayl yo'lini qaytaradi, aks holda "".
Public Function ExportPechat(ByVal isCMYK As Boolean, ByVal brighten As Boolean, ByVal brightPct As Long, _
        ByVal printType As String, ByVal ext As String, ByVal dpi As Long, _
        ByVal folder As String, ByVal onlySelection As Boolean) As String
    Dim doc As Document, oldUnit As cdrUnit
    Dim origSel As ShapeRange, exportSr As ShapeRange
    Dim tmpLayer As Layer, lensShape As Shape
    Dim x As Double, y As Double, w As Double, h As Double
    Dim wPx As Long, hPx As Long
    Dim fullPath As String, flt As ExportFilter
    Dim filterType As cdrFilter, rng As cdrExportRange, imgType As cdrImageType

    Set doc = ActiveDocument
    If Not GetExportBounds(onlySelection, x, y, w, h) Then
        MsgBox "Eksport uchun obyekt tanlanmagan.", vbExclamation, "Pechat eksport"
        Exit Function
    End If

    Select Case LCase$(ext)
        Case "tif", "tiff": filterType = cdrTIFF
        Case "png": filterType = cdrPNG: isCMYK = False   ' PNG CMYK ni qo'llamaydi
        Case Else: filterType = cdrJPEG
    End Select
    imgType = IIf(isCMYK, cdrCMYKColorImage, cdrRGBColorImage)

    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    fullPath = folder & BuildFileName(w, h, printType, isCMYK, ext)
    If Len(Dir$(fullPath)) > 0 Then
        If MsgBox("Bu fayl allaqachon bor:" & vbCrLf & fullPath & vbCrLf & vbCrLf & _
                  "Ustidan yozilsinmi?", vbYesNo + vbQuestion, "Pechat eksport") <> vbYes Then Exit Function
        Kill fullPath
    End If

    ' Piksel o'lchami: sm -> dyuym -> piksel
    wPx = CLng(w / 2.54 * dpi)
    hPx = CLng(h / 2.54 * dpi)

    Set origSel = ActiveSelectionRange
    oldUnit = doc.Unit
    On Error GoTo Fail
    doc.Unit = cdrCentimeter

    ' Ochroq qilish: vaqtinchalik qatlamda "Brighten" linzasi bilan to'rtburchak
    If brighten And brightPct > 0 Then
        Set tmpLayer = ActivePage.CreateLayer("__ochroq_tmp")
        Set lensShape = tmpLayer.CreateRectangle2(x, y, w, h)
        lensShape.Fill.ApplyNoFill
        lensShape.Outline.SetNoOutline
        lensShape.CreateLens cdrLensBrighten, brightPct
    End If

    If onlySelection Then
        Set exportSr = CreateShapeRange
        exportSr.AddRange origSel
        If Not lensShape Is Nothing Then exportSr.Add lensShape
        exportSr.CreateSelection
        rng = cdrSelection
    Else
        rng = cdrCurrentPage
    End If

    Set flt = doc.ExportBitmap(fullPath, filterType, rng, imgType, wPx, hPx, dpi, dpi, _
                               cdrNormalAntiAliasing, False, False, True, False, cdrCompressionNone)
    flt.Finish
    ExportPechat = fullPath

Cleanup:
    On Error Resume Next
    If Not lensShape Is Nothing Then lensShape.Delete
    If Not tmpLayer Is Nothing Then tmpLayer.Delete
    doc.Unit = oldUnit
    If origSel.Count > 0 Then origSel.CreateSelection Else doc.ClearSelection
    Exit Function

Fail:
    MsgBox "Eksportda xatolik: " & Err.Description, vbCritical, "Pechat eksport"
    ExportPechat = ""
    Resume Cleanup
End Function
