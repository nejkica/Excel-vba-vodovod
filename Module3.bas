Attribute VB_Name = "Module3"
Public Sub Rek(ByVal stolpec As String, ByVal podizvajalec As String)
    
'    stolpec = "N"
    stolpec1 = ""
    prejsnjiStolpec = ""
UserFormMain.ProgressBar1.Value = 30
UserFormMain.LabelOdbelavaPodatkov.Caption = "obdelujem zavihek REK ..."
    Sheets("REK").Activate
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
'-------------------------- najprej popišem vse stolpce
    
    stStolpec = Range(stolpec & 1).Column
    stStolpec1 = stStolpec - 1
    stStolpec2 = stStolpec - 2
    stStolpec3 = stStolpec - 3
    
    tStolpec = Split(Cells(1, stStolpec).Address(), "$")(1)
    tStolpec1 = Split(Cells(1, stStolpec1).Address(), "$")(1)
    tStolpec2 = Split(Cells(1, stStolpec2).Address(), "$")(1)
    tStolpec3 = Split(Cells(1, stStolpec3).Address(), "$")(1)
    
'    If Len(stolpec) = 1 Then
'        prejsnjiStolpec = Chr(Asc(stolpec) - 1)
'    ElseIf Len(stolpec) = 2 Then
'        If stolpec = "AA" Then
'            prejsnjiStolpec = "Z"
'        Else
'            stolpec1 = Mid(stolpec, 2, 1)
'            prejsnjiStolpec = "A" & Chr(Asc(stolpec1) - 1)
'        End If
'
'    End If
'------------------------- konec - najprej popišem vse stolpce
    Debug.Print prejsnjiStolpec
    
    
    Dim vrsticeZaObdelavo() As String
    Dim vrsticeZaObdelavoSteber() As Variant
    
    '----------katere vrstice potegnemo v desno
    If podizvajalec = "pokerznik" Then
        Columns(tStolpec & ":" & tStolpec).ColumnWidth = 27
        ActiveSheet.PageSetup.PrintArea = "$A$1:$" & tStolpec & "$40"
        
        vrsticeZaObdelavo = Split("3,5,7,20,22,23,24,25,27,29,31,33,35,37", ",")
        PovleciDesnoPokerznik tStolpec, tStolpec1, vrsticeZaObdelavo
        
        Range(tStolpec1 & "8:" & tStolpec1 & "19").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        For i = 8 To 19
            Cells(i, tStolpec).Formula = "=D" & i & "-E" & i & "-" & tStolpec1 & i
        Next i
    ElseIf podizvajalec = "steber" Then
        Columns(tStolpec & ":" & tStolpec).ColumnWidth = 13
        Columns(tStolpec1 & ":" & tStolpec1).ColumnWidth = 8
        
        ActiveSheet.PageSetup.PrintArea = "$A$1:$" & tStolpec & "$56"
        
        vrsticeZaObdelavo = Split("4,5,6,12,14,26,28,36,38,40,42,46,48,50,51,53,54,56", ",")
        PovleciDesnoSteber tStolpec, tStolpec1, tStolpec2, tStolpec3, vrsticeZaObdelavo
        
        vrsticeZaObdelavoSteber = Array(7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 29, 30, 31, 32, 33, 34, 35, 36, 37, 39, 40, 41, 43, 46)
        
        For Each e In vrsticeZaObdelavoSteber
            Range(tStolpec3 & e & ":" & tStolpec2 & e).Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Cells(e, tStolpec).Formula = "=G" & e & "-I" & e & "-SUM(V" & e & ":" & tStolpec2 & e & ")"
        Next e
                
    End If
    
    
    
End Sub
Sub PovleciDesnoPokerznik(ByVal s As String, ByVal ps As String, ByRef vrsObd() As String)
    For Each e In vrsObd
        Range(ps & e).Select
        Selection.AutoFill Destination:=Range(ps & e & ":" & s & e), Type:=xlFillDefault
    
    Next e
    
End Sub

Sub PovleciDesnoSteber(ByVal ts As String, ByVal ts1 As String, ByVal ts2 As String, ByVal ts3 As String, ByRef vrsObd() As String)
    For Each e In vrsObd
        Range(ts3 & e & ":" & ts2 & e).Select
        Selection.AutoFill Destination:=Range(ts3 & e & ":" & ts & e), Type:=xlFillDefault
    
    Next e
    
End Sub

Sub test()
    colName = "AX"
    stolpec = Range(colName & 1).Column
Debug.Print stolpec
    
    crka = Split(Cells(1, stolpec).Address(), "$")(1)
Debug.Print crka

End Sub
