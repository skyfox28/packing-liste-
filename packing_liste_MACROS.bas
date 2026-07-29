Attribute VB_Name = "Module1"
' ===========================================================================
' PACKING LISTE — Macros VBA
' ===========================================================================
' Installation (une seule fois) :
'   1. Ouvrez packing_liste.xlsx dans Excel.
'   2. Alt+F11 pour ouvrir l'éditeur VBA.
'   3. Dans l'arborescence à gauche, double-cliquez sur la feuille "Scan"
'      (sous "Microsoft Excel Objects") et collez le premier bloc ci-dessous
'      ("PARTIE 1 — À coller dans le module de la feuille Scan").
'   4. Clic droit sur le projet > Insert > Module, et collez le second bloc
'      ("PARTIE 2 — À coller dans un module standard").
'   5. Fichier > Enregistrer sous > type "Classeur Excel (.xlsm)".
'      (Un .xlsx ne peut pas contenir de macros.)
'   6. À l'ouverture, autorisez les macros si Excel le demande.
'
' Pour lancer la synthèse : Alt+F8, sélectionnez GenerateSynthese, Exécuter.
' (Ou : Développeur > Insérer > Bouton, dessinez-le sur la feuille Synthese,
' et affectez-lui la macro GenerateSynthese pour un simple clic.)
' ===========================================================================


' =====================================================================
' PARTIE 1 — À coller dans le module de la feuille "Scan"
' (double-cliquer sur l'onglet "Scan" dans l'explorateur de projets VBA)
' =====================================================================
'
' Option Explicit
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     ' Scan en colonne A (ligne >= 5) -> curseur saute sur la Quantité (G)
'     ' Quantité validée en G -> curseur redescend sur A, ligne suivante
'     On Error GoTo CleanUp
'     If Target.Cells.Count > 1 Then Exit Sub
'     If Target.Row < 5 Then Exit Sub
'
'     Application.EnableEvents = False
'
'     If Target.Column = 1 Then                    ' colonne A : code scanné
'         If Target.Value <> "" Then
'             Me.Cells(Target.Row, "G").Select
'         End If
'     ElseIf Target.Column = 7 Then                 ' colonne G : quantité
'         If Target.Value <> "" Then
'             Me.Cells(Target.Row + 1, "A").Select
'         End If
'     End If
'
' CleanUp:
'     Application.EnableEvents = True
' End Sub


' =====================================================================
' PARTIE 2 — À coller dans un module standard (Insert > Module)
' =====================================================================

Option Explicit

Sub GenerateSynthese()
    ' Reconstruit les feuilles "Synthese" et "Detail_palettes" à partir des
    ' données de la feuille "Scan", selon les mêmes règles que l'export
    ' Excel de l'application web (voir Spec_export_synthese_palettes.md) :
    ' QteStd = quantité la plus fréquente par article ; tri des articles par
    ' quantité totale décroissante ; palettes complètes (par lot puis n°
    ' scan) avant les incomplètes ; lignes sans code article ignorées.
    Dim wsScan As Worksheet, wsSynth As Worksheet, wsDetail As Worksheet
    On Error Resume Next
    Set wsScan = ThisWorkbook.Sheets("Scan")
    Set wsSynth = ThisWorkbook.Sheets("Synthese")
    Set wsDetail = ThisWorkbook.Sheets("Detail_palettes")
    On Error GoTo 0
    If wsScan Is Nothing Or wsSynth Is Nothing Or wsDetail Is Nothing Then
        MsgBox "Feuilles introuvables (attendu : Scan, Synthese, Detail_palettes).", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim lastRow As Long
    lastRow = wsScan.Cells(wsScan.Rows.Count, "A").End(xlUp).Row

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 5 To lastRow
        Dim codeArt As String
        codeArt = Trim(CStr(wsScan.Cells(r, "D").Value))
        If codeArt <> "" Then
            Dim des As String, lot As String, dluo As String, qty As Double
            des = CStr(wsScan.Cells(r, "C").Value)
            lot = CStr(wsScan.Cells(r, "F").Value)
            If IsDate(wsScan.Cells(r, "E").Value) Then
                dluo = Format(wsScan.Cells(r, "E").Value, "dd/mm/yyyy")
            Else
                dluo = ""
            End If
            qty = 0
            If IsNumeric(wsScan.Cells(r, "G").Value) Then qty = wsScan.Cells(r, "G").Value

            Dim g As Object
            If Not groups.Exists(codeArt) Then
                Set g = CreateObject("Scripting.Dictionary")
                g("code") = codeArt
                g("des") = des
                g("pal") = 0
                g("qty") = 0
                Dim freq As Object
                Set freq = CreateObject("Scripting.Dictionary")
                g("freq") = freq
                Dim items As Collection
                Set items = New Collection
                g("items") = items
                groups.Add codeArt, g
            End If
            Set g = groups(codeArt)
            g("pal") = g("pal") + 1
            g("qty") = g("qty") + qty

            Dim freqD As Object
            Set freqD = g("freq")
            Dim qKey As String
            qKey = CStr(qty)
            If freqD.Exists(qKey) Then
                freqD(qKey) = freqD(qKey) + 1
            Else
                freqD(qKey) = 1
            End If

            Dim itemColl As Collection
            Set itemColl = g("items")
            Dim itm As Object
            Set itm = CreateObject("Scripting.Dictionary")
            itm("seq") = r - 4
            itm("lot") = lot
            itm("dluo") = dluo
            itm("qty") = qty
            itemColl.Add itm
        End If
    Next r

    ' QteStd = mode des quantités (en cas d'égalité, la plus grande valeur)
    Dim key As Variant
    For Each key In groups.Keys
        Set g = groups(key)
        Set freqD = g("freq")
        Dim bestQ As Double, bestCount As Long
        bestQ = 0: bestCount = -1
        Dim fk As Variant
        For Each fk In freqD.Keys
            Dim cnt As Long
            cnt = freqD(fk)
            If cnt > bestCount Or (cnt = bestCount And CDbl(fk) > bestQ) Then
                bestCount = cnt
                bestQ = CDbl(fk)
            End If
        Next fk
        g("qteStd") = bestQ
    Next key

    ' Tri des articles par quantité totale décroissante
    Dim n As Long
    n = groups.Count
    Dim keysArr() As Variant
    If n > 0 Then
        ReDim keysArr(1 To n)
        Dim i As Long
        i = 1
        For Each key In groups.Keys
            keysArr(i) = key
            i = i + 1
        Next key
        Dim a As Long, b As Long, tmp As Variant
        For a = 1 To n - 1
            For b = 1 To n - a
                If groups(keysArr(b))("qty") < groups(keysArr(b + 1))("qty") Then
                    tmp = keysArr(b): keysArr(b) = keysArr(b + 1): keysArr(b + 1) = tmp
                End If
            Next b
        Next a
    Else
        ReDim keysArr(1 To 0)
    End If

    WriteSynthSheet wsSynth, groups, keysArr, n
    WriteDetailSheet wsDetail, groups, keysArr, n

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Synthèse mise à jour : " & n & " article(s).", vbInformation
End Sub

Private Function HexColor(ByVal rgbHex As String) As Long
    HexColor = RGB(CLng("&H" & Mid(rgbHex, 1, 2)), CLng("&H" & Mid(rgbHex, 3, 2)), CLng("&H" & Mid(rgbHex, 5, 2)))
End Function

Private Sub StyleRange(rng As Range, bg As String, fg As String, bold As Boolean, hAlign As Long)
    If bg <> "" Then rng.Interior.Color = HexColor(bg) Else rng.Interior.ColorIndex = xlNone
    If fg <> "" Then rng.Font.Color = HexColor(fg) Else rng.Font.ColorIndex = xlAutomatic
    rng.Font.Bold = bold
    rng.Font.Name = "Arial"
    rng.HorizontalAlignment = hAlign
    rng.VerticalAlignment = xlCenter
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = HexColor("B0B8C8")
        .Weight = xlThin
    End With
End Sub

Private Sub WriteSynthSheet(ws As Worksheet, groups As Object, keysArr() As Variant, n As Long)
    ws.Cells.Clear
    ws.Range("A1:E1").Merge
    ws.Range("A1").Value = "SYNTHÈSE PACKING LISTE — " & Format(Date, "dd/mm/yyyy")
    StyleRange ws.Range("A1:E1"), "1F3864", "FFFFFF", True, xlCenter
    ws.Range("A1").Font.Size = 13

    Dim totalPal As Long, totalQty As Double, k As Long
    totalPal = 0: totalQty = 0
    For k = 1 To n
        totalPal = totalPal + groups(keysArr(k))("pal")
        totalQty = totalQty + groups(keysArr(k))("qty")
    Next k
    ws.Range("A2:E2").Merge
    ws.Range("A2").Value = "Total : " & totalPal & " palettes  •  " & totalQty & " unités"
    StyleRange ws.Range("A2:E2"), "D9E1F2", "1F3864", True, xlCenter

    Dim headers As Variant, c As Long
    headers = Array("Code article", "Désignation", "Qté / palette", "Palettes", "Quantité totale")
    For c = 0 To 4
        ws.Cells(4, c + 1).Value = headers(c)
    Next c
    StyleRange ws.Range("A4:E4"), "2E5496", "FFFFFF", True, xlCenter

    Dim r As Long, g As Object
    For k = 1 To n
        r = 4 + k
        Set g = groups(keysArr(k))
        ws.Cells(r, 1).Value = g("code")
        ws.Cells(r, 2).Value = g("des")
        ws.Cells(r, 3).Value = g("qteStd")
        ws.Cells(r, 4).Value = g("pal")
        ws.Cells(r, 5).Value = g("qty")
        StyleRange ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)), IIf(k Mod 2 = 0, "EAF0FA", "FFFFFF"), "", False, xlCenter
        ws.Cells(r, 2).HorizontalAlignment = xlLeft
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).NumberFormat = "#,##0"
    Next k

    Dim totalRow As Long
    totalRow = 5 + n
    ws.Cells(totalRow, 1).Value = "TOTAL"
    If n > 0 Then
        ws.Cells(totalRow, 4).Formula = "=SUM(D5:D" & (4 + n) & ")"
        ws.Cells(totalRow, 5).Formula = "=SUM(E5:E" & (4 + n) & ")"
    Else
        ws.Cells(totalRow, 4).Value = 0
        ws.Cells(totalRow, 5).Value = 0
    End If
    StyleRange ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, 5)), "FFE699", "", True, xlCenter
    ws.Cells(totalRow, 2).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(totalRow, 4), ws.Cells(totalRow, 5)).NumberFormat = "#,##0"

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 16
    ws.Activate
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub WriteDetailSheet(ws As Worksheet, groups As Object, keysArr() As Variant, n As Long)
    ws.Cells.Clear
    ws.Range("A1:F1").Merge
    ws.Range("A1").Value = "DÉTAIL PALETTES PAR ARTICLE"
    StyleRange ws.Range("A1:F1"), "1F3864", "FFFFFF", True, xlCenter
    ws.Range("A1").Font.Size = 13

    ws.Range("A2:F2").Merge
    ws.Range("A2").Value = "Tri : palettes complètes d'abord (par lot puis n° scan), incomplètes en fin de bloc"
    StyleRange ws.Range("A2:F2"), "D9E1F2", "1F3864", True, xlCenter

    Dim headers As Variant, c As Long
    headers = Array("N° palette", "Code article", "Désignation", "Lot", "DLUO", "Quantité")
    For c = 0 To 5
        ws.Cells(4, c + 1).Value = headers(c)
    Next c
    StyleRange ws.Range("A4:F4"), "2E5496", "FFFFFF", True, xlCenter

    Dim r As Long
    r = 5
    Dim k As Long, totalQty As Double
    totalQty = 0
    Dim segStarts As Collection, segEnds As Collection
    Set segStarts = New Collection
    Set segEnds = New Collection

    For k = 1 To n
        Dim g As Object
        Set g = groups(keysArr(k))
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Merge
        ws.Cells(r, 1).Value = g("code") & " — " & g("des") & "   (" & g("pal") & " pal. / " & g("qty") & " u. / std " & g("qteStd") & ")"
        StyleRange ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)), "C9D6EA", "1F3864", True, xlLeft
        r = r + 1

        Dim segStart As Long
        segStart = r

        Dim itemColl As Collection
        Set itemColl = g("items")
        Dim m As Long
        m = itemColl.Count
        If m > 0 Then
            Dim arr() As Object
            ReDim arr(1 To m)
            Dim idx As Long
            idx = 1
            Dim it As Variant
            For Each it In itemColl
                Set arr(idx) = it
                idx = idx + 1
            Next it

            ' Tri par insertion : complètes (qté >= QteStd) avant incomplètes,
            ' puis par lot croissant, puis par n° scan croissant
            Dim x As Long, y As Long
            For x = 2 To m
                Dim cur As Object
                Set cur = arr(x)
                Dim curIncomplete As Boolean
                curIncomplete = (cur("qty") < g("qteStd"))
                y = x - 1
                Do While y >= 1
                    Dim cmpIncomplete As Boolean
                    cmpIncomplete = (arr(y)("qty") < g("qteStd"))
                    Dim doSwap As Boolean
                    doSwap = False
                    If cmpIncomplete <> curIncomplete Then
                        If curIncomplete = False And cmpIncomplete = True Then doSwap = True
                    ElseIf StrComp(CStr(arr(y)("lot")), CStr(cur("lot")), vbTextCompare) > 0 Then
                        doSwap = True
                    ElseIf StrComp(CStr(arr(y)("lot")), CStr(cur("lot")), vbTextCompare) = 0 And arr(y)("seq") > cur("seq") Then
                        doSwap = True
                    End If
                    If doSwap Then
                        Set arr(y + 1) = arr(y)
                        y = y - 1
                    Else
                        Exit Do
                    End If
                Loop
                Set arr(y + 1) = cur
            Next x

            For x = 1 To m
                Dim p As Object
                Set p = arr(x)
                Dim incomplete As Boolean
                incomplete = (p("qty") < g("qteStd"))
                ws.Cells(r, 1).Value = p("seq")
                ws.Cells(r, 2).Value = g("code")
                ws.Cells(r, 3).Value = g("des")
                ws.Cells(r, 4).Value = p("lot")
                ws.Cells(r, 5).Value = p("dluo")
                ws.Cells(r, 6).Value = p("qty")
                totalQty = totalQty + p("qty")
                Dim rowBg As String
                If incomplete Then
                    rowBg = "FCE4D6"
                ElseIf (r - segStart) Mod 2 = 1 Then
                    rowBg = "EAF0FA"
                Else
                    rowBg = "FFFFFF"
                End If
                StyleRange ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)), rowBg, "", False, xlCenter
                ws.Cells(r, 3).HorizontalAlignment = xlLeft
                ws.Cells(r, 1).NumberFormat = "#,##0"
                ws.Cells(r, 6).NumberFormat = "#,##0"
                r = r + 1
            Next x
        End If

        If r > segStart Then
            segStarts.Add segStart
            segEnds.Add r - 1
        End If
    Next k

    Dim totalRow As Long
    totalRow = r
    ws.Cells(totalRow, 1).Value = "TOTAL"
    If segStarts.Count > 0 Then
        Dim f As String
        f = "=SUM("
        Dim si As Long
        For si = 1 To segStarts.Count
            If si > 1 Then f = f & ","
            f = f & "F" & segStarts(si) & ":F" & segEnds(si)
        Next si
        f = f & ")"
        ws.Cells(totalRow, 6).Formula = f
    Else
        ws.Cells(totalRow, 6).Value = 0
    End If
    StyleRange ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, 6)), "FFE699", "", True, xlCenter
    ws.Cells(totalRow, 3).HorizontalAlignment = xlLeft
    ws.Cells(totalRow, 6).NumberFormat = "#,##0"

    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 11
    ws.Activate
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
End Sub
