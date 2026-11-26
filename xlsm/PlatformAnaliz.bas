Attribute VB_Name = "PlatformAnaliz"
Option Explicit

' PLATFORM DEGERLENDIRME - created by EYALTIRIKLI
' Hexagon raporunu (.xlsx) okur, her platformu PASS / FAIL degerlendirir.
' Kural (her platform, N = o platformdaki nokta sayisi):
'   |sapma| <= .005             iceride
'   .005 < |sapma| <= .010      "OUT PROFILE .010" noktasi
'   |sapma| > .010              "OUT PROFILE .020" noktasi
'   PASS = .020 noktasi yok  ve  .010 noktasi < N x %5 (yuvarlama yok:
'          39 nokta -> 1,95: 1 nokta gecer, 2 nokta FAIL)  ve  |ortalama| <= .004
'   Tam .004 / .005 / .010 kabul.

Private Const TOL_IN As Double = 0.005
Private Const TOL_OUT As Double = 0.01
Private Const TOL_AVG As Double = 0.004
Private Const EPS As Double = 0.000000001
Private Const HEAD_ROW As Long = 6
Private Const MAX_ROWS As Long = 2000
Private Const BASLIK As String = "PLATFORM EVALUATION"

' ---------------------------------------------------------------- dugmeler

Public Sub RaporYukle()
    Dim f As Variant
    f = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
                                    "Select Hexagon report")
    If VarType(f) = vbBoolean Then Exit Sub
    Analiz CStr(f)
End Sub

Public Sub SonucuKaydet()
    Dim f As Variant, yeni As Workbook, s As Worksheet, k As Long
    If Len(CStr(AnalizSayfasi().Cells(HEAD_ROW + 1, 1).Value)) = 0 Then
        MsgBox "Load a report first (LOAD REPORT).", vbInformation, BASLIK
        Exit Sub
    End If
    f = Application.GetSaveAsFilename(InitialFileName:=DosyaAdi(), _
                                      FileFilter:="Excel (*.xlsx),*.xlsx")
    If VarType(f) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(Array(AnalizSayfasi().Name, NoktaSayfasi().Name)).Copy
    Set yeni = ActiveWorkbook
    For Each s In yeni.Worksheets
        For k = s.Shapes.Count To 1 Step -1
            s.Shapes(k).Delete
        Next k
    Next s
    Application.DisplayAlerts = False
    yeni.SaveAs Filename:=CStr(f), FileFormat:=51
    Application.DisplayAlerts = True
    yeni.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Saved:" & vbCrLf & CStr(f), vbInformation, BASLIK
End Sub

Public Sub Temizle()
    Dim sh As Worksheet, adr As Variant, i As Long
    Set sh = AnalizSayfasi()
    Application.ScreenUpdating = False
    adr = Array("B3", "F3", "B4", "F4")
    For i = 0 To 3
        sh.Range(adr(i)).Value = ""
    Next i
    sh.Range("A2").Value = "No report loaded."
    sh.Range(sh.Cells(HEAD_ROW + 1, 1), sh.Cells(HEAD_ROW + MAX_ROWS, 8)).Clear
    With sh.Range("H3")
        .Value = "-"
        .Interior.Color = RGB(166, 166, 166)
        .Font.Color = RGB(255, 255, 255)
    End With
    sh.Range("H5").Value = ""
    With NoktaSayfasi()
        .Cells.Clear
        .Cells(1, 1).Value = "POINT"
    End With
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------- analiz

Public Sub Analiz(ByVal yol As String)
    Dim wb As Workbook, ws As Worksheet, sm As Worksheet
    Dim v As Variant, msg As String
    Dim info(1 To 10) As String
    On Error GoTo Hata
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set wb = Workbooks.Open(Filename:=yol, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set ws = SayfaBul(wb, "platform")
    If ws Is Nothing Then
        msg = "'Platform' sheet not found."
        GoTo Bitir
    End If
    v = AlanOku(ws)
    Set sm = SayfaBul(wb, "summary")
    If sm Is Nothing Then Set sm = wb.Worksheets(1)
    BilgiOku sm, info
    wb.Close SaveChanges:=False
    Set wb = Nothing
    msg = Degerlendir(v, info, yol)
Bitir:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    AnalizSayfasi().Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Len(msg) > 0 Then MsgBox msg, vbExclamation, BASLIK
    Exit Sub
Hata:
    msg = "Error: " & Err.Description
    Resume Bitir
End Sub

Private Function Degerlendir(v As Variant, info() As String, ByVal yol As String) As String
    Dim nr As Long, nc As Long, r As Long, c As Long, i As Long
    Dim bas As Long, son As Long, npt As Long, np As Long, d As Double, a As Double
    nr = UBound(v, 1)
    nc = UBound(v, 2)

    ' veri: A sutununda 1, 2, 3 ... olan satirlar (alttaki tolerans satirlari degil)
    For r = 1 To nr
        If TamSayi(v(r, 1), 1) Then
            For c = 2 To nc
                If Sayi(v(r, c), d) Then
                    bas = r
                    Exit For
                End If
            Next c
            If bas > 0 Then Exit For
        End If
    Next r
    If bas = 0 Then
        Degerlendir = "No data found on the Platform sheet" & vbCrLf & _
                      "(point numbers 1, 2, 3 ... expected in column A)."
        Exit Function
    End If
    son = bas
    Do While son + 1 <= nr
        If Not TamSayi(v(son + 1, 1), son - bas + 2) Then Exit Do
        son = son + 1
    Loop
    npt = son - bas + 1

    ' platformlar: B'den saga sayi iceren her sutun.  Basliginda "SPF"
    ' gecen sutun (her noktadaki en buyuk deger) platform degildir, atlanir.
    Dim cols() As Long, h As Long, spf As Boolean
    ReDim cols(1 To nc)
    For c = 2 To nc
        spf = False
        For h = bas - 1 To Application.Max(1, bas - 3) Step -1
            If InStr(1, Norm(v(h, c)), "spf") = 1 Then spf = True
        Next h
        If spf Then GoTo SonrakiSutun
        For r = bas To son
            If Sayi(v(r, c), d) Then
                np = np + 1
                cols(np) = c
                Exit For
            End If
        Next r
SonrakiSutun:
    Next c
    If np = 0 Then
        Degerlendir = "No numeric deviation values on the Platform sheet."
        Exit Function
    End If

    Dim n() As Long, c1() As Long, c2() As Long, lim() As Double, c1Ok() As Boolean
    Dim ort() As Double, mx() As Double, ok() As Boolean, ortOk() As Boolean
    Dim pts() As Variant, top As Double, gecen As Long
    ReDim n(1 To np): ReDim c1(1 To np): ReDim c2(1 To np): ReDim lim(1 To np): ReDim c1Ok(1 To np)
    ReDim ort(1 To np): ReDim mx(1 To np): ReDim ok(1 To np): ReDim ortOk(1 To np)
    ReDim pts(1 To npt, 1 To np)
    For i = 1 To np
        top = 0
        For r = bas To son
            If Sayi(v(r, cols(i)), d) Then
                pts(r - bas + 1, i) = d
                n(i) = n(i) + 1
                top = top + d
                a = Abs(d)
                If a > mx(i) Then mx(i) = a
                If a > TOL_OUT + EPS Then
                    c2(i) = c2(i) + 1
                ElseIf a > TOL_IN + EPS Then
                    c1(i) = c1(i) + 1
                End If
            End If
        Next r
        If n(i) > 0 Then ort(i) = top / n(i)
        lim(i) = n(i) * 5 / 100                     ' %5, yuvarlanmaz
        c1Ok(i) = (c1(i) * 100 < n(i) * 5)          ' tam sayilarla: c1 < N x %5
        ortOk(i) = (Abs(ort(i)) <= TOL_AVG + EPS)
        ok(i) = (n(i) > 0 And c2(i) = 0 And c1Ok(i) And ortOk(i))
        If ok(i) Then gecen = gecen + 1
    Next i

    ' ---- ANALIZ sayfasi
    Dim sh As Worksheet, rg As Range, o() As Variant, adr As Variant, genel As Boolean
    Set sh = AnalizSayfasi()
    ' Part No, Serial No (ilk 5 hane), CMM, Date
    sh.Range("B3").Value = "'" & info(1)
    sh.Range("F3").Value = "'" & Left$(info(2), 5)
    sh.Range("B4").Value = "'" & info(5)
    sh.Range("F4").Value = "'" & info(8)
    sh.Range("A2").Value = "File: " & DosyaKisa(yol) & "     Analyzed: " & Format(Now, "dd.mm.yyyy hh:nn")

    sh.Range(sh.Cells(HEAD_ROW + 1, 1), sh.Cells(HEAD_ROW + MAX_ROWS, 8)).Clear
    ReDim o(1 To np, 1 To 8)
    For i = 1 To np
        o(i, 1) = "PLATFORM #" & i
        o(i, 2) = c1(i)
        o(i, 3) = c2(i)
        o(i, 4) = ort(i)
        o(i, 5) = n(i)
        o(i, 6) = lim(i)
        o(i, 7) = mx(i)
        o(i, 8) = IIf(ok(i), "PASS", "FAIL")
    Next i
    Set rg = sh.Cells(HEAD_ROW + 1, 1).Resize(np, 8)
    rg.Value = o
    Cerceve rg
    rg.Columns(1).HorizontalAlignment = xlLeft
    rg.Columns(1).Font.Bold = True
    rg.Columns(4).NumberFormat = "0.00000"
    rg.Columns(6).NumberFormat = "0.00"
    rg.Columns(7).NumberFormat = "0.0000"
    For i = 1 To np
        r = HEAD_ROW + i
        Boya sh.Cells(r, 8), ok(i)
        If Not c1Ok(i) Then sh.Cells(r, 2).Interior.Color = RGB(255, 124, 128)
        If c2(i) > 0 Then sh.Cells(r, 3).Interior.Color = RGB(255, 124, 128)
        If Not ortOk(i) Then sh.Cells(r, 4).Interior.Color = RGB(255, 124, 128)
    Next i

    genel = (gecen = np)
    r = HEAD_ROW + np + 2
    sh.Cells(r, 1).Value = "OVERALL RESULT"
    sh.Cells(r, 2).Value = gecen & " / " & np & " platforms PASS"
    sh.Cells(r, 8).Value = IIf(genel, "PASS", "FAIL")
    Cerceve sh.Cells(r, 1)
    sh.Cells(r, 1).Font.Bold = True
    sh.Cells(r, 2).Font.Name = "Calibri"
    Cerceve sh.Cells(r, 8)
    Boya sh.Cells(r, 8), genel
    sh.Cells(r + 1, 1).Value = "Rule: |dev| <= .005 in profile;  points with .005 < |dev| <= .010 " & _
        "< 5% of total points;  no point with |dev| > .010;  |average| <= .004."
    With sh.Cells(r + 1, 1).Font
        .Italic = True
        .Size = 9
        .Color = RGB(89, 89, 89)
    End With
    With sh.Range("H3")
        .Value = IIf(genel, "PASS", "FAIL")
        .Interior.Color = IIf(genel, RGB(0, 176, 80), RGB(255, 0, 0))
        .Font.Color = RGB(255, 255, 255)
    End With
    sh.Range("H5").Value = gecen & " / " & np & " platforms PASS"

    ' ---- NOKTALAR sayfasi
    Dim sn As Worksheet
    Set sn = NoktaSayfasi()
    sn.Cells.Clear
    sn.Cells(1, 1).Value = "POINT"
    Cerceve sn.Cells(1, 1)
    sn.Cells(1, 1).Font.Bold = True
    sn.Cells(1, 1).Interior.Color = RGB(217, 225, 242)
    For i = 1 To np
        sn.Cells(1, i + 1).Value = "#" & i
        Boya sn.Cells(1, i + 1), ok(i)
    Next i
    Cerceve sn.Range(sn.Cells(1, 2), sn.Cells(1, np + 1))
    For r = 1 To npt
        sn.Cells(r + 1, 1).Value = r
    Next r
    Cerceve sn.Range(sn.Cells(2, 1), sn.Cells(npt + 1, 1))
    sn.Range(sn.Cells(2, 1), sn.Cells(npt + 1, 1)).Font.Bold = True
    Set rg = sn.Range(sn.Cells(2, 2), sn.Cells(npt + 1, np + 1))
    rg.Value = pts
    Cerceve rg
    rg.NumberFormat = "0.0000"
    rg.Interior.Color = RGB(198, 239, 206)
    For i = 1 To np
        For r = 1 To npt
            If Not IsEmpty(pts(r, i)) Then
                a = Abs(pts(r, i))
                If a > TOL_OUT + EPS Then
                    rg.Cells(r, i).Interior.Color = RGB(255, 124, 128)
                ElseIf a > TOL_IN + EPS Then
                    rg.Cells(r, i).Interior.Color = RGB(255, 217, 102)
                End If
            Else
                rg.Cells(r, i).Interior.Pattern = xlNone
            End If
        Next r
    Next i
    sn.Columns(1).ColumnWidth = 8
    sn.Range(sn.Columns(2), sn.Columns(np + 1)).ColumnWidth = 9
End Function

' ---------------------------------------------------------------- sayfalar

' ANALYSIS ve POINTS sayfalari adlarindan bulunur (yoksa 1. ve 2. sayfa)
Private Function AnalizSayfasi() As Worksheet
    Set AnalizSayfasi = BuSayfa("anal", 1)
End Function

Private Function NoktaSayfasi() As Worksheet
    Set NoktaSayfasi = BuSayfa("point", 2)
End Function

Private Function BuSayfa(ByVal bas As String, ByVal sira As Long) As Worksheet
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If InStr(1, Norm(s.Name), bas) = 1 Then
            Set BuSayfa = s
            Exit Function
        End If
    Next s
    Set BuSayfa = ThisWorkbook.Worksheets(sira)
End Function

' ---------------------------------------------------------------- okuma

Private Function SayfaBul(wb As Workbook, ByVal kelime As String) As Worksheet
    Dim s As Worksheet
    For Each s In wb.Worksheets
        If Norm(s.Name) = kelime Then
            Set SayfaBul = s
            Exit Function
        End If
    Next s
    For Each s In wb.Worksheets
        If InStr(1, Norm(s.Name), kelime) > 0 Then
            Set SayfaBul = s
            Exit Function
        End If
    Next s
End Function

Private Function AlanOku(ws As Worksheet) As Variant
    Dim lr As Long, lc As Long
    With ws.UsedRange
        lr = .Row + .Rows.Count - 1
        lc = .Column + .Columns.Count - 1
    End With
    If lr < 2 Then lr = 2
    If lc < 2 Then lc = 2
    AlanOku = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value
End Function

' Summary: her etiketin degeri altindaki ilk dolu hucre (3 satira kadar),
' yoksa sagindaki.  Part No bulunamazsa B3.
Private Sub BilgiOku(sm As Worksheet, info() As String)
    Dim lab As Variant, v As Variant, r As Long, c As Long, k As Long, x As String
    Dim bulundu(0 To 9) As Boolean
    lab = Array("partno", "serialno", "operation", "operator", "cmm", "ncws", _
                "spindle", "date", "company", "location")
    v = sm.Range("A1:AN60").Value
    For r = 1 To 60
        For c = 1 To 40
            If Not IsEmpty(v(r, c)) Then
                x = Norm(v(r, c))
                If Len(x) > 0 Then
                    For k = 0 To 9
                        If Not bulundu(k) And x = lab(k) Then
                            bulundu(k) = True
                            info(k + 1) = EtiketDegeri(v, r, c, lab)
                        End If
                    Next k
                End If
            End If
        Next c
    Next r
    If Len(info(1)) = 0 Then info(1) = Metin(v(3, 2))
End Sub

Private Function EtiketDegeri(v As Variant, ByVal r As Long, ByVal c As Long, lab As Variant) As String
    Dim k As Long, t As String
    For k = 1 To 3
        If r + k <= 60 Then
            t = Metin(v(r + k, c))
            If Len(t) > 0 Then
                If Not EtiketMi(t, lab) Then EtiketDegeri = t
                If Len(EtiketDegeri) > 0 Then Exit Function
                Exit For
            End If
        End If
    Next k
    For k = 1 To 3
        If c + k <= 40 Then
            t = Metin(v(r, c + k))
            If Len(t) > 0 Then
                If Not EtiketMi(t, lab) Then EtiketDegeri = t
                Exit Function
            End If
        End If
    Next k
End Function

Private Function EtiketMi(ByVal t As String, lab As Variant) As Boolean
    Dim k As Long, x As String
    x = Norm(t)
    For k = 0 To UBound(lab)
        If x = lab(k) Then
            EtiketMi = True
            Exit Function
        End If
    Next k
End Function

' ---------------------------------------------------------------- yardimcilar

Private Function Metin(ByVal x As Variant) As String
    If IsError(x) Or IsEmpty(x) Or IsNull(x) Then Exit Function
    If VarType(x) = vbDate Then
        Metin = Format(x, "dd.mm.yyyy hh:nn")
    Else
        Metin = Trim$(CStr(x))
    End If
End Function

Private Function Norm(ByVal x As Variant) As String
    Dim s As String, i As Long, ch As String
    s = LCase$(Metin(x))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9]" Then Norm = Norm & ch
    Next i
End Function

Private Function Sayi(ByVal x As Variant, ByRef d As Double) As Boolean
    Dim s As String
    Select Case VarType(x)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, 20
            d = CDbl(x)
            Sayi = True
        Case vbString
            s = Replace(Trim$(x), ",", ".")
            If Len(s) = 0 Then Exit Function
            If s Like "*[!0-9.eE+-]*" Then Exit Function
            If Not s Like "*[0-9]*" Then Exit Function
            d = Val(s)
            Sayi = True
    End Select
End Function

Private Function TamSayi(ByVal x As Variant, ByVal istenen As Long) As Boolean
    Dim d As Double
    If Not Sayi(x, d) Then Exit Function
    If Abs(d - Round(d, 0)) > 0.000000001 Then Exit Function
    TamSayi = (CLng(d) = istenen)
End Function

Private Sub Cerceve(rg As Range)
    With rg
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        If .Columns.Count > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
        If .Rows.Count > 1 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub

Private Sub Boya(cel As Range, ByVal gecti As Boolean)
    cel.Interior.Color = IIf(gecti, RGB(0, 176, 80), RGB(255, 0, 0))
    cel.Font.Color = RGB(255, 255, 255)
    cel.Font.Bold = True
End Sub

Private Function DosyaKisa(ByVal yol As String) As String
    Dim k As Long
    k = InStrRev(Replace(yol, "\", "/"), "/")
    DosyaKisa = Mid$(yol, k + 1)
End Function

Private Function DosyaAdi() As String
    Dim s As String, ch As Variant
    s = Trim$(CStr(AnalizSayfasi().Range("B3").Value))
    If Len(Trim$(CStr(AnalizSayfasi().Range("F3").Value))) > 0 Then s = s & "_" & Trim$(CStr(AnalizSayfasi().Range("F3").Value))
    If Len(s) = 0 Then s = "REPORT"
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "_")
    Next ch
    DosyaAdi = s & "_PLATFORM.xlsx"
End Function
