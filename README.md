# DEMO_EXCEL_VBA_Mehrsprachige_GUI
## Demo-Code eines mehrsprachigen UserForms in Excel VBA 

#### Code des Moduls "modGUI"
````vba
Option Explicit

'Beschreibung
'------------
    'Diese Demo zeigt eine Möglichkeit, wie Steuerelemente
    'auf einem UserForm in verschiedenen Sprachen dargestellt
    'werden können.
    'Die Beschriftungen werden hier nicht im VBA-Code angelegt,
    'sondern aus einem Tabellenblatt zur Laufzeit ausgelesen.
    'Der größte Vorteil ist, dass so auch Texte mit zum Beispiel
    'kyrillischen Schriftzeichen leicht verwendet werden können,
    'was in VBA sonst extrem umständlich ist.
    'Weitere Vorteile sind ein schlankerer Code, eine leichte
    'Anpassung der Texte und die Möglichkeit schnell weitere Sprachen
    'zu implementieren.
    'Wenn der Code dann als Excel-Add-in (xlam-Datei) ausgeliefert wird,
    'dann werden die hier benutzten Tabellenblätter nur in der Add-in-Datei
    'verwendet und der Benutzer bekommt davon gar nichts mit, wenn er
    'das Add-in verwendet.

'Vorbereitung
'------------
    'Anlegen eines UserForm (hier "frmGUI") mit den Steuerelementen.
    'Einmaliges Auslesen der Steuerelemente mit einer For-Each-Schleife,
    'wobei die Namen und evtl. die ursprünglichen Beschriftungen der
    'Elemente in Spalten eines Tabellenblatts geschrieben werden (hier
    'mit der Methode "auslesenSteuerelemente").
    'Ausfüllen der Spalten auf dem Tabellenblatt mit den Texten für die
    'Steuerelemente in den gewünschten Sprachen (hier Deutsch in Spalte C,
    'Englisch in Spalte D und Russisch in Spalte E)

'Code
'----
'Einmaliges Auslesen der Steuerelemente während der Entwicklung
Private Sub auslesenSteuerelemente()
    
    'Variablen
    Dim steuerelement As Control
    Dim zeile As Integer
    
    'Alle Steuerelemente auf dem UserForm durchlaufen und die
    'Namen (also die ID) in Spalte A des ersten Tabellenblatts schreiben
    zeile = 1 'Zähler für die Zeile auf dem Tabellenblatt auf 1 setzen
    For Each steuerelement In frmGUI.Controls
        Worksheets(1).Cells(zeile, 1).Value = steuerelement.Name 'ID in Zelle schreiben
        zeile = zeile + 1
    Next steuerelement
    
    '(Optional:) Alle Steuerelemente auf dem UserForm durchlaufen und die
    'Beschriftungen in Spalte B des ersten Tabellenblatts schreiben
    zeile = 1 'Zähler für die Zeile auf dem Tabellenblatt auf 1 setzen
    For Each steuerelement In frmGUI.Controls
        Worksheets(1).Cells(zeile, 2).Value = steuerelement.Caption 'Beschriftung in Zelle schreiben
        zeile = zeile + 1
    Next steuerelement
    
End Sub

'Demo-GUI "Mehrsprachigkeit" aufrufen
Sub startenGUI()
    frmGUI.Show
End Sub
````

#### Code des UserForms "frmGUI"
````vba
Option Explicit

'Beschreibung
'------------
    'Der Code für das Klick-Ereignis der Schaltflächen ist
    'weitgehend identisch und sollte per Refactoring optimiert
    'werden. Für ein leichteres Verständnis habe ich ihn in dieser
    'Demo redundant programmiert.
    'Da Excel die Reihenfolge der Steuerelemente auf dem UserForm
    'kennt, kann man sich darauf verlassen, dass die For-Each-Schleife
    'in der gleichen Reihenfolge vorgeht wie beim Auslesen und somit
    'die Beschriftungen den korrekten Elementen zuweist.
    'Werden die Steuerelemte auf dem UserForm verändert, dann sollte
    'die Prozedur "auslesenSteuerelemente" erneut ausgeführt werden.

'Variablen (modulweit gültig)
Dim steuerelement As Control
Dim zeile As Integer

Private Sub CommandButton1_Click() 'Klick auf die Schaltfläche "DE"
    
    'Zähler für die Zeile auf dem Tabellenblatt auf 1 setzen
    zeile = 1
    
    'Spalte C durchlaufen und die Texte der "Caption"-Eigenschaft
    'den Steuerelementen auf dem UserForm zuweisen
    For Each steuerelement In frmGUI.Controls
        steuerelement.Caption = Worksheets(1).Cells(zeile, 3).Value 'Spalte 3 ("C") auslesen
        zeile = zeile + 1
    Next steuerelement
    
End Sub

Private Sub CommandButton2_Click() 'Klick auf die Schaltfläche "EN"

    'Zähler für die Zeile auf dem Tabellenblatt auf 1 setzen
    zeile = 1
    
    'Spalte D durchlaufen und die Texte der "Caption"-Eigenschaft
    'den Steuerelementen auf dem UserForm zuweisen
    For Each steuerelement In frmGUI.Controls
        steuerelement.Caption = Worksheets(1).Cells(zeile, 4).Value 'Spalte 4 ("D") auslesen
        zeile = zeile + 1
    Next steuerelement

End Sub

Private Sub CommandButton3_Click() 'Klick auf die Schaltfläche "RUS"

    'Zähler für die Zeile auf dem Tabellenblatt auf Zeile 1 setzen
    zeile = 1
    
    'Spalte E durchlaufen und die Texte der "Caption"-Eigenschaft
    'den Steuerelementen auf dem UserForm zuweisen
    For Each steuerelement In frmGUI.Controls
        steuerelement.Caption = Worksheets(1).Cells(zeile, 5).Value 'Spalte 5 ("E") auslesen
        zeile = zeile + 1
    Next steuerelement

End Sub
````
