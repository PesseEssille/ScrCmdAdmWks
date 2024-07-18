@ECHO OFF
SET RepLog=%USERPROFILE%\journaux\
IF NOT EXIST "%RepLog%." MKDIR "%RepLog%"
SET RepLog=%RepLog%\demar\
IF NOT EXIST "%RepLog%." MKDIR "%RepLog%"
SET FicLog="%RepLog%%~n0-%date:~6,4%%date:~3,2%%date:~,2%.log"

CALL :log %~0 === DEBUT ===

SET cscpt=%windir%\System32\cscript.exe //nologo
SET clnfl="%~dp0cleanfiles.vbs"
CALL :log %clnfl% purge fichiers temporaires
CALL :log suppression des fichiers temporaires de plus de 7j
%cscpt% %clnfl% -L /S>> %FicLog%

CALL :log suppression des journaux de d‚marrage de plus de 7j
%cscpt% %clnfl% -r:"%RepLog%" -L>> %FicLog%

CALL :log suppression des fichiers d'Outlook de plus de 15j
%cscpt% %clnfl% -r:"C:\Users\pcsi-\AppData\Local\Packages\microsoft.windowscommunicationsapps_8wekyb3d8bbwe\LocalState\Files\S0" -L -d:15>> %FicLog%

SET cscpt=
SET clnfl=

REM fin du nettoyage - lancement des applications
REM ==== M  A  I  N  ===
SET AppNam=KeePass
CALL :LogApp %AppNam% Password Safe
START /d "%ProgramFiles%\%AppNam%\" %AppNam%.exe
CALL :attend_lance %AppNam% %AppNam%

SET AppNam=Thunderbird
CALL :LogApp Mozilla %AppNam%
START /d "%ProgramFiles%\Mozilla %AppNam%\" %AppNam%.exe
CALL :attend_lance %AppNam% Tous les messages - *
IF %CodRet% == "KO" CALL :attend_lance %AppNam% Courrier entrant - *
REM CALL :sleep 5

SET AppNam=Firefox
CALL :LogApp Mozilla %AppNam%
START /d "%ProgramFiles%\Mozilla %AppNam%\" %AppNam%.exe
CALL :attend_lance %AppNam% Mozilla %AppNam%
REM CALL :sleep 5

SET prgm=msedge
CALL :LogApp Microsoft Edge - Outlook ITS
START /d "%ProgramFiles(x86)%\Microsoft\Edge\Application\" %prgm%.exe https://outlook.office.com/mail/
CALL :attend_lance %prgm% Courrier*

:fin
CALL :log %~0 === F I N ===
GOTO :eof

REM *** sous-programmes ***

REM horodatage dans le journal d'un message
:log
ECHO @ %date% %time% = "%*"
ECHO @ %date% %time% = "%*">> %FicLog%
GOTO :eof

REM Titre de l'application
:LogApp
ECHO.
ECHO.>> %FicLog%
CALL :log _____-----{{{{{%*}}}}}-----_____
GOTO :eof

REM pause d'un nombre de secondes
:sleep
SET /a nbsec=%1+1
CALL :log %1 secondes d'attente. %2 %3 %4 %5 %6 %7 %8 %9
ping -n %nbsec% localhost > nul
GOTO :eof

(
Attente qu'un ex‚cutable soit en mode running
PARAM 1 : nom de l'ex‚cutable. ajout de .exe si absent … la fin
paramŠtres suivants : nom de la fenˆtre, si besoin
) 2>nul
:attend_lance
SET FicExe=%1   &:: Le 1er paramŠtre contient le nom de l'ex‚cutable
IF NOT "%FicExe:~-4%" == ".exe" SET FicExe=%FicExe%.exe
REM les paramŠtres suivants, facultatifs, contiennent le nom de la fenˆtre
SET AllPrm=%*
FOR /F "tokens=1*" %%a in ("%AllPrm%") DO SET NomFen=%%b

SET essai=1 &:: compteur du nombre d'essais
SET CodRet=KO &:: code de retour de la fonction

REM tƒche avec le nom de l'application et du statut
SET cmd2run=tasklist /nh /v /fi ^"IMAGENAME eq %FicExe%^" /fi ^"STATUS eq RUNNING^" &:: commande … ex‚cuter
REM si le nom de la fenˆtre est fournis, on le concatŠne avec la commande
IF "%NomFen%" NEQ "" SET cmd2run=%cmd2run% /fi ^"WINDOWTITLE eq %NomFen%^"
REM filtre du retour de la commande pour avoir un code d'erreur
SET cmd2fnd=%cmd2run% ^| findstr /i /r ^"^^%FicExe% ^"

:boucle
ECHO ______________ essai num‚ro {%essai%} ______________
ECHO ______________ essai num‚ro {%essai%} ______________>>%FicLog%

REM liste des tƒches avec le nom de l'application, le statut et en option le nom de la fenˆtre
REM sortie terminal
SET cmd2run
%cmd2run%
REM sortie journal
SET cmd2run>>%FicLog%
%cmd2run%>>%FicLog% 2>&1

REM filtr‚ sur le nom, pour avoir un code erreur
REM sortie terminal
SET cmd2fnd
%cmd2fnd%
REM sortie journal
SET cmd2fnd>>%FicLog%
%cmd2fnd%>>%FicLog% 2>&1

IF ERRORLEVEL 1 (SET CodRet=KO) ELSE (SET CodRet=OK)
CALL :log CodRet=[%CodRet%]
CALL :log CodRet=[%CodRet%]>>%FicLog%
IF "%CodRet%" == "OK" GOTO :eof

SET /A essai=essai+1
IF %essai% GTR 30 GOTO :eof &:: sortie de la fct si plus de 30 essais

CALL :sleep 2 que [%FicExe%] ait le statut 'RUNNING' (%essai%e essai)

GOTO :boucle
