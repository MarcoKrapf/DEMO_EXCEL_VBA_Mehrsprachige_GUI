Attribute VB_Name = "modGUI"
Option Explicit

'Beschreibung
'------------
    'Diese Demo zeigt eine M�glichkeit, wie Steuerelemente
    'auf einem UserForm in verschiedenen Sprachen dargestellt
    'werden k�nnen.
    'Die Beschriftungen werden hier nicht im VBA-Code angelegt,
    'sondern aus einem Tabellenblatt zur Laufzeit ausgelesen.
    'Der gr��te Vorteil ist, dass so auch Texte mit zum Beispiel
    'kyrillischen Schriftzeichen leicht verwendet werden k�nnen,
    'was in VBA sonst extrem umst�ndlich ist.
    'Weitere Vorteile sind ein schlankerer Code, eine leichte
    'Anpassung der Texte und die M�glichkeit schnell weitere Sprachen
    'zu implementieren.
    'Wenn der Code dann als Excel-Add-in (xlam-Datei) ausgeliefert wird,
    'dann werden die hier benutzten Tabellenbl�tter nur in der Add-in-Datei
    'verwendet und der Benutzer bekommt davon gar nichts mit, wenn er
    'das Add-in verwendet.

'Vorbereitung
'------------
    'Anlegen eines UserForm (hier "frmGUI") mit den Steuerelementen.
    'Einmaliges Auslesen der Steuerelemente mit einer For-Each-Schleife,
    'wobei die Namen und evtl. die urspr�nglichen Beschriftungen der
    'Elemente in Spalten eines Tabellenblatts geschrieben werden (hier
    'mit der Methode "auslesenSteuerelemente").
    'Ausf�llen der Spalten auf dem Tabellenblatt mit den Texten f�r die
    'Steuerelemente in den gew�nschten Sprachen (hier Deutsch in Spalte C,
    'Englisch in Spalte D und Russisch in Spalte E)

'Code
'----
'Einmaliges Auslesen der Steuerelemente w�hrend der Entwicklung
Private Sub auslesenSteuerelemente()
    
    'Variablen
    Dim steuerelement As Control
    Dim zeile As Integer
    
    'Alle Steuerelemente auf dem UserForm durchlaufen und die
    'Namen (also die ID) in Spalte A des ersten Tabellenblatts schreiben
    zeile = 1 'Z�hler f�r die Zeile auf dem Tabellenblatt auf 1 setzen
    For Each steuerelement In frmGUI.Controls
        Worksheets(1).Cells(zeile, 1).Value = steuerelement.Name 'ID in Zelle schreiben
        zeile = zeile + 1
    Next steuerelement
    
    '(Optional:) Alle Steuerelemente auf dem UserForm durchlaufen und die
    'Beschriftungen in Spalte B des ersten Tabellenblatts schreiben
    zeile = 1 'Z�hler f�r die Zeile auf dem Tabellenblatt auf 1 setzen
    For Each steuerelement In frmGUI.Controls
        Worksheets(1).Cells(zeile, 2).Value = steuerelement.Caption 'Beschriftung in Zelle schreiben
        zeile = zeile + 1
    Next steuerelement
    
End Sub

'Demo-GUI "Mehrsprachigkeit" aufrufen
Sub startenGUI()
    frmGUI.Show
End Sub
