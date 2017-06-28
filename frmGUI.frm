VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGUI 
   Caption         =   "Mehrsprachigkeit"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2640
   OleObjectBlob   =   "frmGUI.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Beschreibung
'------------
    'Der Code f�r das Klick-Ereignis der Schaltfl�chen ist
    'weitgehend identisch und sollte per Refactoring optimiert
    'werden. F�r ein leichteres Verst�ndnis habe ich ihn in dieser
    'Demo redundant programmiert.
    'Da Excel die Reihenfolge der Steuerelemente auf dem UserForm
    'kennt, kann man sich darauf verlassen, dass die For-Each-Schleife
    'in der gleichen Reihenfolge vorgeht wie beim Auslesen und somit
    'die Beschriftungen den korrekten Elementen zuweist.
    'Werden die Steuerelemte auf dem UserForm ver�ndert, dann sollte
    'die Prozedur "auslesenSteuerelemente" erneut ausgef�hrt werden.

'Variablen (modulweit g�ltig)
Dim steuerelement As Control
Dim zeile As Integer

Private Sub CommandButton1_Click() 'Klick auf die Schaltfl�che "DE"
    
    'Z�hler f�r die Zeile auf dem Tabellenblatt auf 1 setzen
    zeile = 1
    
    'Spalte C durchlaufen und die Texte der "Caption"-Eigenschaft
    'den Steuerelementen auf dem UserForm zuweisen
    For Each steuerelement In frmGUI.Controls
        steuerelement.Caption = Worksheets(1).Cells(zeile, 3).Value 'Spalte 3 ("C") auslesen
        zeile = zeile + 1
    Next steuerelement
    
End Sub

Private Sub CommandButton2_Click() 'Klick auf die Schaltfl�che "EN"

    'Z�hler f�r die Zeile auf dem Tabellenblatt auf 1 setzen
    zeile = 1
    
    'Spalte D durchlaufen und die Texte der "Caption"-Eigenschaft
    'den Steuerelementen auf dem UserForm zuweisen
    For Each steuerelement In frmGUI.Controls
        steuerelement.Caption = Worksheets(1).Cells(zeile, 4).Value 'Spalte 4 ("D") auslesen
        zeile = zeile + 1
    Next steuerelement

End Sub

Private Sub CommandButton3_Click() 'Klick auf die Schaltfl�che "RUS"

    'Z�hler f�r die Zeile auf dem Tabellenblatt auf Zeile 1 setzen
    zeile = 1
    
    'Spalte E durchlaufen und die Texte der "Caption"-Eigenschaft
    'den Steuerelementen auf dem UserForm zuweisen
    For Each steuerelement In frmGUI.Controls
        steuerelement.Caption = Worksheets(1).Cells(zeile, 5).Value 'Spalte 5 ("E") auslesen
        zeile = zeile + 1
    Next steuerelement

End Sub
