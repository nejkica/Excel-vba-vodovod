Attribute VB_Name = "Module4"
Public Sub ProNepredvidena(ByVal zsSituacije As Integer, ByVal stolpec As String, ByVal zavihek As String)
    'stolpec = "N"
    'zsSituacije = 9
    
    Sheets(zavihek).Activate
    
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect "mojdenar"
    End If
    
    Dim naslovSituacije As String
    
    naslovSituacije = zsSituacije & ". situacija"
    
    If (zavihek = "Nepredvidena") Then
        naslovSituacije = naslovSituacije & " - " & zavihek & " dela"
    End If
    
    Cells(8, "C").Value = naslovSituacije
    
    stolpecArhivK = Cells(12, Columns.Count).End(xlToLeft).Column + 1
    stolpecArhivV = stolpecArhivK + 1
    
    zadnjaVrsticaG = Range("G65536").End(xlUp).Row
    Range("G13:G" & zadnjaVrsticaG).Select
    Selection.Copy
    
    ciljniStolpecImeK = Split(Cells(1, stolpecArhivK).Address, "$")(1)
    Range(ciljniStolpecImeK & "13:" & ciljniStolpecImeK & zadnjaVrsticaG).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    zadnjaVrsticaI = Range("I65536").End(xlUp).Row
    Range("I13:I" & zadnjaVrsticaI).Select
    Selection.Copy
    
    ciljniStolpecImeV = Split(Cells(1, stolpecArhivV).Address, "$")(1)
    Range(ciljniStolpecImeV & "13:" & ciljniStolpecImeV & zadnjaVrsticaI).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '--------------zapiši naslov arhivov
    Cells(12, stolpecArhivK).Value = zsSituacije - 1 & ".sitK"
    Cells(12, stolpecArhivV).Value = zsSituacije - 1 & ".sitV"
    
    Columns(ciljniStolpecImeK & ":" & ciljniStolpecImeV).Select
    Selection.NumberFormat = "#,##0.00"
    
    
    Range("H14").Select
    ActiveCell.Formula = "=G14-" & ciljniStolpecImeK & "14"
    Selection.AutoFill Destination:=Range("H14:H" & zadnjaVrsticaG)
    'imePSituacijeAK = Split(Cells(, 18 + zsPSituacijeAK).Address, "$")(1)
    'ciljniStolpec = Range("Z1").Column
    'Debug.Print ciljniStolpecImeK
    
End Sub
