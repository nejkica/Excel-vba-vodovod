Attribute VB_Name = "Module2"
Public Function VrziVpodizvajalca() As Object
'
' Makro1 Makro
'
'Disable DoEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    zapIzDat = zsSituacije + 2
    zapIzDatS = StringEnice(zapIzDat)
    
    zacPolje = "A13"
    
    'Dim wbP As Workbook
    Dim wbT As Workbook
    Dim stVrstic As Integer
    
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
    Range(zacPolje).Select
    
    wbP.Sheets("PRO").Activate
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
    Range(zacPolje).Select
    
    UserFormMain.LabelOdbelavaPodatkov = "odpiram izvozno datoteko"
    DoEvents
    '--------------------odpremo še izvoz_0 datoteko 10
    Set wbT = Workbooks.Open("1040_sit_" & zapIzDat & "_izvoz_0.xls", UpdateLinks:=False)
    Application.WindowState = xlMinimized
    wbT.Sheets("PRO").Activate
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    DoEvents
    UserFormMain.LabelOdbelavaPodatkov = "izvozna datoteka odprta - sledi branje ..."

'++++++++++++++++++++++++++++++++++++++++++++++++++od tukaj naprej delamo
    
        
    Dim dict As Object
    Dim dictP As Object
    Dim dictlen As Integer
    Set dict = CreateObject("Scripting.Dictionary")

        
    'grem v podizvajalèevo datoteko (stolpec E) in zapišem wbs kjer je cena
    Set dict = PripraviDict(wbP)
    
    dictlen = dict.Count
    
    'potem grem v izvorno datoteko iskat kolièine (stolpec G)
    Set dictP = KoncniDict(wbT, dict)
    
    'nazadnje grem še enkrat skozi, da zapišem vrednosti v celice
    ZapisiVcelice "PRO", wbP, dictP
    ZapisiVcelice "Nepredvidena", wbP, dictP
    
    wbT.Close savechanges:=False
    
    Application.WindowState = xlNormal
    
    wbP.Save
    'wbP.Close
'Enable
    UserFormMain.LabelOdbelavaPodatkov.Caption = "Konèan uvoz podizvajalca " & podizvajalec
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    'Set VrziVpodizvajalca = wbP

End Function

Private Function ZapisiVcelice(zavihek As String, ByRef wb As Workbook, dict As Object)
    wb.Sheets(zavihek).Activate
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
    stVrst = Range("A14:A30000").Cells.SpecialCells(xlCellTypeConstants).Count
    zacPolje = "A13"
    
    UserFormMain.LabelOdbelavaPodatkov = "zapisujem v celice - zavihek " & zavihek
    UserFormMain.ProgressBar1.Value = 0
    DoEvents
    For i = 1 To stVrst
        odstotek = i * 100 / stVrst
        If odstotek Mod 5 = 0 Then
            DoEvents
            UserFormMain.ProgressBar1.Value = odstotek
        End If
        wbs = wb.Worksheets(zavihek).Range(zacPolje).Offset(i, 0).Value

        kolP = dict(wbs)
        
        If Not (IsEmpty(kolP) Or kolP = 0) Then
            wb.Worksheets(zavihek).Range(zacPolje).Offset(i, 6).Value = kolP
        End If
        
    Next i
    
End Function
Public Function StringEnice(ByVal zs As Integer) As String
    
    tekst = CStr(zs)
    tekstLen = Len(tekst)
    
    If tekstLen = 1 Then
        StringEnice = "0" & tekst
    Else
        StringEnice = tekst
    End If
End Function
Private Function PripraviDict(ByRef wb As Workbook) As Object
    
    wb.Sheets("PRO").Activate
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
    zacPolje = "A13"
    Dim idDict As Object
    Set idDict = CreateObject("Scripting.Dictionary")
    wbs = Empty
    kol = Empty
    cena = Empty
    
    Range(zacPolje).Select
    stV = Range("A13:A30000").Cells.SpecialCells(xlCellTypeConstants).Count
        
    
    UserFormMain.LabelOdbelavaPodatkov = "branje zavihka PRO poteka"
    UserFormMain.ProgressBar1.Value = 0
    DoEvents
    For i = 1 To stV
        odstotek = i * 100 / stV
        If odstotek Mod 5 = 0 Then
            DoEvents
            UserFormMain.ProgressBar1.Value = odstotek
        End If
        
        wbs = wb.Worksheets("PRO").Range(zacPolje).Offset(i, 0).Value
        cena = wb.Worksheets("PRO").Range(zacPolje).Offset(i, 4).Value
        
        If IsEmpty(wbs) Then GoTo nadaljuj0
        
        If IsEmpty(cena) Then GoTo nadaljuj0
        
        idDict.Add wbs, 0

        'Debug.Print wbs & " " & idDict(wbs)
        
nadaljuj0:
    Next i
    
    wb.Sheets("Nepredvidena").Activate
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
    Range(zacPolje).Select
    stVrst = Range("A13:A30000").Cells.SpecialCells(xlCellTypeConstants).Count
    
    
    UserFormMain.LabelOdbelavaPodatkov = "branje zavihka Nepredvidena poteka"
    UserFormMain.ProgressBar1.Value = 0
    DoEvents
    For j = 1 To stVrst
        odstotek = i * 100 / stVrst
        If odstotek Mod 5 = 0 Then
            DoEvents
            UserFormMain.ProgressBar1.Value = odstotek
        End If
        
        wbs = wb.Worksheets("Nepredvidena").Range(zacPolje).Offset(i, 0).Value
        cena = wb.Worksheets("Nepredvidena").Range(zacPolje).Offset(i, 4).Value
        
        If IsEmpty(wbs) Then GoTo nadaljuj1
        
        If IsEmpty(cena) Then GoTo nadaljuj1
        
        idDict.Add wbs, 0

nadaljuj1:
    Next j

    Set PripraviDict = idDict
End Function

Private Function KoncniDict(ByRef wb As Workbook, dict As Object) As Object
    
    wb.Sheets("PRO").Activate
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
    zacPolje = "A13"
    stV = Range("A13:A30000").Cells.SpecialCells(xlCellTypeConstants).Count
    Dim idDict As Object
    Set idDict = CreateObject("Scripting.Dictionary")
    wbs = Empty
    kol = Empty
    cena = Empty
    
    
    UserFormMain.LabelOdbelavaPodatkov = "delam koèno matriko"
    UserFormMain.ProgressBar1.Value = 0
    DoEvents
    For i = 1 To stV
        odstotek = i * 100 / stV
        If odstotek Mod 5 = 0 Then
            DoEvents
            UserFormMain.ProgressBar1.Value = odstotek
        End If
        wbs = wb.Worksheets("PRO").Range(zacPolje).Offset(i, 0).Value
        kol = wb.Worksheets("PRO").Range(zacPolje).Offset(i, 6).Value
        
        If IsEmpty(wbs) Then GoTo nadaljuj0
        
        If dict.exists(wbs) Then
            idDict(wbs) = kol
'            If Not kol = 0 Then
'                Debug.Print wbs & " " & idDict(wbs)
'            End If
        End If
        
nadaljuj0:
    Next i
    
    Set KoncniDict = idDict
End Function
