Attribute VB_Name = "Module1"
Public wbP As Workbook
Public zsSituacije As Integer

Public Function ObdelajSrajcko(ByVal staraDatPodizvajalca As String) As Object
    'Disable DoEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
UserFormMain.ProgressBar1.Value = 0
    Dim razsekanoImePoti() As String
    razsekanoImePoti = Split(staraDatPodizvajalca, "\")
    imeDatoteke = razsekanoImePoti(UBound(razsekanoImePoti))
    imeDatotekeL = Len(imeDatoteke)
    pot = Left(staraDatPodizvajalca, Len(staraDatPodizvajalca) - imeDatotekeL)
    razsekanoImeDatoteke = Split(imeDatoteke, "_")
    Dim podizvajalec As String
    podizvajalec = razsekanoImeDatoteke(2)
'Debug.Print (podizvajalec)
    zapStSituacije = CInt(razsekanoImeDatoteke(0))
    zapNvSituacije = zapStSituacije + 1
    
    zapStSituacijeS = StringEnice(zapStSituacije)
    zapNvSituacijeS = StringEnice(zapNvSituacije)
    
    Dim datum() As String
    datum = Split(Date, ".")
    datum(1) = StringEnice(CInt(datum(1)))
    datumS = datum(2) & "-" & datum(1) & "-" & datum(0)
    
    
    
    Dim wbP As Workbook
    
    
    '--------------------------------save as in dodamo datum 8->9
UserFormMain.LabelOdbelavaPodatkov = "poteka Shrani kot - nova situacija"
UserFormMain.ProgressBar1.Value = 10
DoEvents
    Set wbP = Workbooks.Open(staraDatPodizvajalca, UpdateLinks:=False)
    Application.WindowState = xlMinimized
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    wbP.SaveAs pot & zapNvSituacijeS & "_" & "situacija" & "_" & razsekanoImeDatoteke(2) & "_" & datumS
    
UserFormMain.LabelOdbelavaPodatkov = "shranjeno ... nadaljujem z izdelavo srajèke"
UserFormMain.ProgressBar1.Value = 10
DoEvents
    wbP.Sheets("sit").Activate
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
    '-------------------------------gremo na zavihek "sit"
    zacPolje = "A1"
    
    Cells(9, "T").Value = UserFormMain.TextBoxDatumSituacije.Value
    Cells(21, "E").Value = zapNvSituacije
    Cells(23, "H").Value = MonthName(Month(Date) - 1) & " " & Year(Date)
    
    '--------------pogledam katera je prva prosta celica pri vrednostih po mesecih v situaciji
    i = 73
    Do While Cells(i, "T").Value <> ""
        i = i + 1
    Loop
    
    Cells(i, "G").Value = zapNvSituacije & ". vmesna situacija"
    '----------------pogledam v kater stolpec piše v REK
    
    enacba = Range("T" & i - 1).Cells.Formula
    stolpec1 = ""
    If Len(enacba) > 8 Then
        stolpec = Mid(enacba, 6, 1)
        stolpec1 = Mid(enacba, 7, 1)
        
    Else
        stolpec = Mid(enacba, 6, 1)
    End If
    
    If stolpec1 = "" And Not stolpec = "Z" Then
        stolpec = Chr(Asc(stolpec) + 1)
    Else
        If stolpec = "Z" Then
            stolpec1 = "@"
            stolpec = "A"
        End If
        stolpec = stolpec & Chr(Asc(stolpec1) + 1)
    End If
    
    Dim vrsticaSume As Integer
    
    If podizvajalec = "steber" Then
        vrsticaSume = 56
        nrStolpec = Range(stolpec & 1).Column + 1
        stolpec = Split(Cells(1, nrStolpec).Address(), "$")(1) ' ------------- še dodatno pomaknem v desno, ker ima v rek pri stebru po dva stolpca za vsako situacijo
    ElseIf podizvajalec = "pokerznik" Then
        vrsticaSume = 37
    End If
    
    Cells(i - 1, "T").Copy
    Cells(i - 1, "T").PasteSpecial xlPasteValues
    Cells(i, "T").Formula = "=REK!" & stolpec & vrsticaSume
    Cells(92, "T").Formula = "=REK!" & stolpec & vrsticaSume
    
    '-----------------zaženem sub Rek
    Rek stolpec, podizvajalec
UserFormMain.ProgressBar1.Value = 50
UserFormMain.LabelOdbelavaPodatkov.Caption = "Obdelujem zavihek PRO ..."
    ProNepredvidena zapNvSituacije, stolpec, "Pro"
UserFormMain.ProgressBar1.Value = 70
UserFormMain.LabelOdbelavaPodatkov.Caption = "Obdelujem zavihek Nepredvidena ..."
    ProNepredvidena zapNvSituacije, stolpec, "Nepredvidena"
    
    'Enable
UserFormMain.LabelOdbelavaPodatkov.Caption = "srajèka narejena in arhivi nastavljeni za: " & podizvajalec
UserFormMain.ProgressBar1.Value = 100
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    zsSituacije = zapStSituacije
    Set ObdelajSrajcko = wbP
        
End Function
