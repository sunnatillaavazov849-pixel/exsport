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
'  Hech narsa tanlanmagan bo'lsa  -> butun sahifa eksport qilinadi.
'  Bir nechta obyekt tanlangan bo'lsa -> har biri alohida, o'z o'lchami bilan.
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

' Obyekt (yoki sr = Nothing bo'lsa butun sahifa) chegarasi, santimetrda
Public Sub GetBoundsCm(ByVal sr As ShapeRange, x As Double, y As Double, w As Double, h As Double)
    Dim doc As Document, oldUnit As cdrUnit
    Set doc = ActiveDocument
    oldUnit = doc.Unit
    doc.Unit = cdrCentimeter

    If sr Is Nothing Then
        x = ActivePage.LeftX
        y = ActivePage.BottomY
        w = ActivePage.SizeWidth
        h = ActivePage.SizeHeight
    Else
        sr.GetBoundingBox x, y, w, h, True
    End If

    doc.Unit = oldUnit
End Sub

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

' Shu nomli fayl bo'lsa, "(2)", "(3)" ... qo'shib bo'sh nom topadi
Private Function UniquePath(ByVal folder As String, ByVal fileName As String) As String
    Dim baseName As String, ext As String, n As Long, p As String
    baseName = Left$(fileName, InStrRev(fileName, ".") - 1)
    ext = Mid$(fileName, InStrRev(fileName, "."))
    p = folder & fileName
    n = 2
    Do While Len(Dir$(p)) > 0
        p = folder & baseName & " (" & n & ")" & ext
        n = n + 1
    Loop
    UniquePath = p
End Function

' Eksport. Saqlangan fayllar ro'yxatini (har biri yangi qatorda) qaytaradi, xato bo'lsa "".
Public Function ExportPechat(ByVal isCMYK As Boolean, ByVal brightPct As Long, _
        ByVal printType As String, ByVal ext As String, ByVal dpi As Long, _
        ByVal folder As String) As String
    Dim doc As Document, origSel As ShapeRange, one As ShapeRange
    Dim i As Long, path As String, result As String

    Set doc = ActiveDocument
    If LCase$(ext) = "png" Then isCMYK = False   ' PNG CMYK ni qo'llamaydi
    If Right$(folder, 1) <> "\" Then folder = folder & "\"

    Set origSel = ActiveSelectionRange

    If origSel.Count = 0 Then
        ' Hech narsa tanlanmagan - butun sahifa
        result = ExportOne(doc, Nothing, isCMYK, brightPct, printType, ext, dpi, folder)
    Else
        ' Har bir tanlangan obyekt alohida
        For i = 1 To origSel.Count
            Set one = CreateShapeRange
            one.Add origSel(i)
            path = ExportOne(doc, one, isCMYK, brightPct, printType, ext, dpi, folder)
            If Len(path) = 0 Then Exit For
            result = result & path & vbCrLf
        Next i
    End If

    On Error Resume Next
    If origSel.Count > 0 Then origSel.CreateSelection Else doc.ClearSelection
    ExportPechat = result
End Function

' Bitta obyektni (sr = Nothing bo'lsa sahifani) eksport qiladi
Private Function ExportOne(ByVal doc As Document, ByVal sr As ShapeRange, ByVal isCMYK As Boolean, _
        ByVal brightPct As Long, ByVal printType As String, ByVal ext As String, _
        ByVal dpi As Long, ByVal folder As String) As String
    Dim oldUnit As cdrUnit, exportSr As ShapeRange
    Dim tmpLayer As Layer, lensShape As Shape
    Dim x As Double, y As Double, w As Double, h As Double
    Dim fullPath As String, flt As ExportFilter
    Dim filterType As cdrFilter, rng As cdrExportRange, imgType As cdrImageType

    Select Case LCase$(ext)
        Case "tif", "tiff": filterType = cdrTIFF
        Case "png": filterType = cdrPNG
        Case Else: filterType = cdrJPEG
    End Select
    imgType = IIf(isCMYK, cdrCMYKColorImage, cdrRGBColorImage)

    GetBoundsCm sr, x, y, w, h
    fullPath = UniquePath(folder, BuildFileName(w, h, printType, isCMYK, ext))

    oldUnit = doc.Unit
    On Error GoTo Fail
    doc.Unit = cdrCentimeter

    ' Ochroq qilish: vaqtinchalik qatlamda "Brighten" linzasi bilan to'rtburchak
    If brightPct > 0 Then
        Set tmpLayer = ActivePage.CreateLayer("__ochroq_tmp")
        Set lensShape = tmpLayer.CreateRectangle2(x, y, w, h)
        lensShape.Fill.ApplyNoFill
        lensShape.Outline.SetNoOutline
        lensShape.CreateLens cdrLensBrighten, brightPct
    End If

    If sr Is Nothing Then
        rng = cdrCurrentPage
    Else
        Set exportSr = CreateShapeRange
        exportSr.AddRange sr
        If Not lensShape Is Nothing Then exportSr.Add lensShape
        exportSr.CreateSelection
        rng = cdrSelection
    End If

    ' Piksel o'lchami: sm -> dyuym -> piksel
    Set flt = doc.ExportBitmap(fullPath, filterType, rng, imgType, _
                               CLng(w / 2.54 * dpi), CLng(h / 2.54 * dpi), dpi, dpi, _
                               cdrNormalAntiAliasing, False, False, True, False, cdrCompressionNone)
    flt.Finish
    ExportOne = fullPath

Cleanup:
    On Error Resume Next
    If Not lensShape Is Nothing Then lensShape.Delete
    If Not tmpLayer Is Nothing Then tmpLayer.Delete
    doc.Unit = oldUnit
    Exit Function

Fail:
    MsgBox "Eksportda xatolik: " & Err.Description, vbCritical, "Pechat eksport"
    ExportOne = ""
    Resume Cleanup
End Function
