VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMain 
   Caption         =   "Obdelava situacije"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265
   OleObjectBlob   =   "UserFormMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonIzberiDatotekoPinSrajcka_Click()
    
    Dim strFileToOpen As String
    UserFormMain.ProgressBar1.Value = 10
    strFileToOpen = Application.GetOpenFilename _
                    (Title:="Odpri STARO datoteko ", _
                    FileFilter:="Excel Files *.xls* (*.xls*),")
    UserFormMain.ProgressBar1.Value = 100
    Set wbP = ObdelajSrajcko(strFileToOpen)
    
    
    UserFormMain.CommandButtonIzvediPreracun.Enabled = True
    
End Sub

Private Sub CommandButtonIzvediPreracun_Click()
    VrziVpodizvajalca
    Debug.Print wbP.Name
End Sub

Private Sub UserForm_Initialize()
    UserFormMain.ProgressBar1.Value = 0
    Application.WindowState = xlMinimized
    UserFormMain.TextBoxDatumSituacije = DateSerial(Year(Date), Month(Date), 0)
End Sub

Private Sub UserForm_Terminate()
    ThisWorkbook.Close True
    'Application.Quit
End Sub
