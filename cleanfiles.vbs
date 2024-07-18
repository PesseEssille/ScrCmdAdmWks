' Supprime les fichiers et les dossiers âgés de plus de 7 jours dans le répertoire temporaire de l'utilisateur.
' Copyright : Public Domain
' Warranty  : None
'  

' Customer  : PCSI-L
' Filename  : cleanfiles.vbs
' Author    : Cédric Rathgeb
' Date      : 2005-08-08
' Version   : 1.1

' Modifs    : Pascal CROZET
'             prise en compte paramètres
'             vérification existence répertoire
'             journal d'opération facultatif
' Date      : 2011-06-15
' Version   : 1.3

' Modifs    : Pascal CROZET
'             traitement des erreurs en suppression de fichiers ou dossiers
' Date      : 2016-11-18
' Version   : 1.3.1

' Modifs    : Pascal CROZET
'             horodatage de l'affichage
' Date      : 2019-11-02
' Version   : 1.3.2

' Modifs    : Pascal CROZET
'             affichage du nom des fichiers et répertoires supprimés, en mode journal
' Date      : 2020-05-29
' Version   : 1.4.4

' Modifs    : Pascal CROZET
'             paramètre /S pour parcourir aussi les sous-répertoires
' Date      : jeudi 05 mai 2022
' Version   : 1.5

' Modifs    : Pascal CROZET
'             horodatage fichier et répertoire supprimé dans le journal
' Date      : Le dimanche 26 février 2023 à 01h 15
' Version   : 1.6

' Modifs    : Pascal CROZET
'             pas de log si pas de sous-répertoire
' Date      : Le mardi 13 juin 2023 à 08h 23
' Version   : 1.6.1

' Modifs    : Pascal CROZET
'             corrections orthographiques dans les messages et commentaires
' Date      : Le mardi 16 juillet 2024 à 12h 09
' Version   : 1.6.2


Const csVER = "1.6.2"   ' numéro de Version
Const ciDELAI_DEFO = 7  ' jours de rétention

' valeurs par défaut des paramètres
Dim bLog : bLog = False ' pas de journal
Dim bSub : bSub = False ' pas de sous-répertoire
Dim iOlderThanDays : iOlderThanDays = ciDELAI_DEFO
Dim oShell : Set oShell = WScript.CreateObject ("WScript.Shell")
Dim sPath : sPath = oShell.ExpandEnvironmentStrings ("%TEMP%")

Dim sMsg, sOpt, iRep
' examen des arguments en paramètres
Dim oArgs : Set oArgs = WScript.Arguments
If oArgs.Count > 0 Then
  For iRep = 1 To oArgs.Count
    sMsg = oArgs.Item ( iRep - 1 )
    sOpt = UCase ( Left ( sMsg , 3 ) )
    ' répertoire à nettoyer
    If ( sOpt = "/R:" Or sOpt = "-R:" ) And _
        Len ( sMsg ) > 4 Then
      sPath = Mid ( sMsg , 4 )
    ElseIf ( sOpt = "/D:" Or sOpt = "-D:" ) And _
            Len ( sMsg ) > 3 Then
      ' nombre de jours d'ancienneté du fichier
      On Error Resume Next   ' capture des erreurs de conversion
      iOlderThanDays = Abs ( Int ( Mid ( sMsg , 4 ) ) )
      If Err.Number = 0 Then
        If iOlderThanDays < 1 Then iOlderThanDays = ciDELAI_DEFO
      Else
        iOlderThanDays = ciDELAI_DEFO
      End If
      On Error Goto 0
    ElseIf ( sOpt = "/L" Or sOpt = "-L" ) Then
      bLog = True
    ElseIf ( sOpt = "/S" Or sOpt = "-S" ) Then
      bSub = True
    ElseIf ( sOpt = "/H" Or sOpt = "-H" Or _
            sOpt = "-?" Or sOpt = "/?" ) Then
      Help
    End If
  Next
'Else
'  Help
End If
Set oArgs = Nothing

' rapport des noms et nombre des fichiers et répertoires supprimés
Dim sReportFiles        : sReportFiles        = ""
Dim iReportFilesCount   : iReportFilesCount   = 0
Dim sReportFolders      : sReportFolders      = ""
Dim iReportFoldersCount : iReportFoldersCount = 0

' Compute old date
dOldDate = DateAdd ( "d" , 0 - iOlderThanDays, Now () )
Dim oFS : Set oFS = CreateObject ( "Scripting.FileSystemObject" )

' Call clean function
CleanFolder sPath
If bLog Then
  If iReportFilesCount > 0 Then
    sMsg = FormatDateTime ( now ) & " : Le"
    If iReportFilesCount > 1 Then sMsg = sMsg & "s " & iReportFilesCount
    sMsg = sMsg & " fichier"
    If iReportFilesCount > 1 Then sMsg = sMsg & "s"
    sMsg = sMsg & " suivant"
    If iReportFilesCount > 1 Then sMsg = sMsg & "s ont" Else sMsg = sMsg & " a"
    sMsg = sMsg & " été supprimé"
    If iReportFilesCount > 1 Then sMsg = sMsg & "s"
    WScript.Echo sMsg & " : " & vbCrLf & sReportFiles
  Else
    WScript.Echo "* Aucun fichier antérieur au " & _
      WeekdayName ( Weekday ( dOldDate ) ) & " " & FormatDateTime ( dOldDate ) & _
      " supprimé *"
  End If
  If iReportFoldersCount > 0 Then
    sMsg = FormatDateTime ( now ) & " : Le"
    If iReportFoldersCount > 1 Then sMsg = sMsg & "s " & iReportFoldersCount
    sMsg = sMsg & " répertoire"
    If iReportFoldersCount > 1 Then sMsg = sMsg & "s"
    sMsg = sMsg & " vide suivant"
    If iReportFoldersCount > 1 Then sMsg = sMsg & "s ont" Else sMsg = sMsg & " a"
    sMsg = sMsg & " été supprimé"
    If iReportFoldersCount > 1 Then sMsg = sMsg & "s"
    WScript.Echo sMsg & " : " & vbCrLf & sReportFolders
  Else
    If bSub Then WScript.Echo "* Aucun répertoire vide supprimé *"
  End If
End If

Sub CleanFolder ( sCurrentPath )
  If Not oFS.FolderExists ( sCurrentPath ) Then Exit Sub
  ' select current folder
  Set oFolder = oFS.GetFolder ( sCurrentPath )  
  
  If bSub Then
    ' Get sub-folders
    Set oSubFolders = oFolder.SubFolders
    
    ' Do a recursive call if it contains sub-folders
    For Each oCurrentFolder in oSubFolders
      CleanFolder oCurrentFolder.Path
    Next
  End If
  
  ' Get files in current folder
  Set oFiles = oFolder.Files
  
  ' Delete old Files
  For Each oCurrentFile in oFiles
    dFileCre = oCurrentFile.DateCreated : dFileMod = oCurrentFile.DateLastModified
    sFilePath = oCurrentFile.Path : sFileName = oCurrentFile.Name
    ' WScript.Echo iReportFilesCount & "=" & sFilePath & "\" & sFileName & _
      ' " (fic crée le " & _
      ' WeekdayName ( Weekday ( dFileCre ) ) & " " & FormatDateTime ( dFileCre ) & _
      ' " - modifié le " & _
      ' WeekdayName ( Weekday ( dFileMod ) ) & " " & FormatDateTime ( dFileMod ) & _
      ' " - référence le " & _
      ' WeekdayName ( Weekday ( dOldDate ) ) & " " & FormatDateTime ( dOldDate ) & ")"
    
    If dFileCre < dOldDate AND _
       dFileMod < dOldDate Then
'       oCurrentFile.DateLastAccessed < dOldDate AND _
      ' récupération de propriétés du fichier courant

      On Error Resume Next
      oCurrentFile.Delete True
      If Err.Number = 0 Then
        sReportFiles = sReportFiles & vbCrLf & sFilePath & _
          " (crée le " & _
          WeekdayName ( Weekday ( dFileCre ) ) & " " & FormatDateTime ( dFileCre ) & _
          " - modifié le " & _
          WeekdayName ( Weekday ( dFileMod ) ) & " " & FormatDateTime ( dFileMod ) & _
          " - référence le " & _
          WeekdayName ( Weekday ( dOldDate ) ) & " " & FormatDateTime ( dOldDate ) & ")"
        ' WScript.Echo sReportFiles
        iReportFilesCount = iReportFilesCount + 1
      End If
      On Error Goto 0
    End If
  Next
  
  Set oFiles = oFolder.Files
  If oFiles.Count = 0 AND oFolder.Path <> sPath Then
    sFolderPath = oFolder.Path
    dFoldCre = oFolder.DateCreated : dFoldMod = oFolder.DateLastModified
    ' WScript.Echo sFolderPath & "\" & _
      ' " (rép crée le " & _
      ' WeekdayName ( Weekday ( dFoldCre ) ) & " " & FormatDateTime ( dFoldCre ) & _
      ' " - modifié le " & _
      ' WeekdayName ( Weekday ( dFoldMod ) ) & " " & FormatDateTime ( dFoldMod ) & _
      ' " - référence le " & _
      ' WeekdayName ( Weekday ( dOldDate ) ) & " " & FormatDateTime ( dOldDate ) & ")"

    If dFoldCre < dOldDate AND _
       dFoldMod < dOldDate Then
        On Error Resume Next
        oFolder.Delete True
        If Err.Number = 0 Then
          sReportFolders = sReportFolders & vbCrLf & sFolderPath & "\" & _
            " (crée le " & _
            WeekdayName ( Weekday ( dFoldCre ) ) & " " & FormatDateTime ( dFoldCre ) & _
            " - modifié le " & _
            WeekdayName ( Weekday ( dFoldMod ) ) & " " & FormatDateTime ( dFoldMod ) & _
            " - référence le " & _
            WeekdayName ( Weekday ( dOldDate ) ) & " " & FormatDateTime ( dOldDate ) & ")"
          ' WScript.Echo sReportFolders
          iReportFoldersCount = iReportFoldersCount + 1
        End If
        On Error Goto 0
    End if
  End If
End Sub

Sub Help
  Dim sNomScr : sNomScr = WScript.ScriptFullName
  If InStr ( sNomScr , " " ) > 0 Then sNomScr = """" & sNomScr & """"
  sMsg = sNomScr & " - v°" & csVER & vbNewLine & vbNewLine & _
         "Supprime les fichiers et les dossiers âgés de plus de " & ciDELAI_DEFO & _
         " jours dans le répertoire temporaire de l'utilisateur." & vbNewLine & _
         "Options :" & vbNewLine & _
         "      -R:""Nom de répertoire"" ou /R:""Nom de répertoire"" ou -R:NomRepertoire ou /R:NomRepertoire" & vbNewLine & _
         "         spécifie un autre répertoire que le répertoire par défaut [" & oShell.ExpandEnvironmentStrings ("%TEMP%") & "]" & vbNewLine & _
         "         Les unités locales, réseau ou chemins UNC sont acceptées." & vbNewLine & _
         "         Si le nom du répertoire contient des espaces, il doit être entouré de guillemets double." & vbNewLine & _
         "      -D:j ou /D:j" & vbNewLine & _
         "         j Indique le nombre de jours d'ancienneté du fichier, que le script doit supprimer" & vbNewLine & _
         "         " & ciDELAI_DEFO & " par défaut. Si ce nombre est inférieur à 1, il est mis à " & ciDELAI_DEFO & vbNewLine & _
         "      -S ou /S" & vbNewLine & _
         "         Parcours aussi les sous-répertoires" & vbNewLine & _
         "      -L ou /L" & vbNewLine & _
         "         Affiche un journal des opérations après l’exécution" & vbNewLine & _
         "         Aucun journal par défaut, le script est conçu pour être automatisé." & vbNewLine & _
         "      -H ou /H ou -? ou /?" & vbNewLine &_
         "         Affiche cette page d'aide" & vbNewLine & _
         "Exemple :" & vbNewLine & _
         "      " & sNomScr & " -r:c:\logs /D:15 -s" & vbNewLine & _
         "      va supprimer les fichiers et dossiers de plus de 15 jours du répertoire C:\Logs et de ses sous-répertoires"
  WScript.Echo sMsg
  WScript.Quit
End Sub