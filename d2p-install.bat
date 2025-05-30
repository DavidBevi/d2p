@echo off
setlocal enabledelayedexpansion

echo ,----------------------------------,
echo ^| d2p - doc 2 pdf - bulk converter ^|
echo '----------------------------------'

:: CHECK FILEPATH OF THIS INSTANCE ################################################################################
:: If this file is not launched from its final location it will try to fix it
:: --> Admin: you'll be guided in installing
:: --> User: you'll be asked to relaunch as admin
set "finalFilePath=%SYSTEMROOT%\System32\d2p.bat"
set "thisFilePath=%~f0"

if /I not "!thisFilePath!"=="!finalFilePath!" (
    >nul 2>&1 "%SystemRoot%\System32\cacls.exe" "%SystemRoot%\System32\config\system"
    if errorlevel 1 (
        echo Hi, reopen this as admin to manage d2p install
        echo.
        pause
        exit /b 1
    ) else (
        echo Hi admin. You are running this script from:
        echo     !thisFilePath!
        echo.

        :: FILE CHECK ###########################################################
        if exist "!finalFilePath!" (
            echo --^> d2p.bat installed in %SystemRoot%\System32
            fc /b "!thisFilePath!" "!finalFilePath!" >nul 2>&1
            if errorlevel 1 (
                echo     Installed version is DIFFERENT from this file, reinstall recommended.
            ) else (
                echo     Installed version is EQUAL to this file, reinstall is possible, but not recommended.
            )
        )

        set "answer="
        set /p answer="--> (Re)install? (y/n): "

        if /I "!answer!"=="y" (
            copy /Y "!thisFilePath!" "%SystemRoot%\System32\d2p.bat" >nul
            if errorlevel 1 (
                echo     ERROR ^(not installed^)
            ) else (
                echo     DONE ^(installed^)
            )
        ) else (
            echo     Installation skipped
        )
        echo.

        :: COMMAND CHECK ########################################################
        where d2p.bat >nul 2>&1
        if errorlevel 1 (
            echo --^> 'd2p' command is NOT installed.
        ) else (
            echo --^> 'd2p' command IS installed. Reinstalling is possible, but not recommended.
        )

        set "answer="
        set /p answer="--> (Re)install? (y/n): "

        if /I "!answer!"=="y" (
            reg add "HKCU\Environment" /v PATH /t REG_EXPAND_SZ /d "%PATH%;%SystemRoot%\System32" /f >nul 2>&1

            if !errorlevel!==0 (
                echo     User PATH updated with %SystemRoot%\System32
                echo     REBOOT to use 'd2p' command
            ) else (
                echo     ERROR: cannot update user PATH ^(and enable 'd2p' command^)
            )
        ) else (
            echo     Installation skipped
        )
        echo.

        pause
        exit /b 1
    )
)


:: D2P ACTUAL CODE #####################################################################################################
:: This will only run if the script is launched from <YourDriveLetter>:\Windows\System32\d2p.bat


:: --- Calculate target folder ---
if "%~1"=="" (
    set "TARGET=%cd%"
) else (
    set "TARGET=%~1"
)

:: Verify existance
if not exist "%TARGET%" (
    echo Folder "%TARGET%" does not exist.
    exit /b 1
)

:: Prepare temp file %VBS%
set "VBS=%temp%\d2p-temp.vbs"
del "%VBS%" >nul 2>&1

:: Extract lines that start with ##> and write in temp file %VBS%
for /f "usebackq tokens=1* delims=>" %%A in (`findstr /b "##>" "%thisFilePath%"`) do (
    >>"%VBS%" echo(%%B
)

:: Execute VBS script
cscript //nologo "%VBS%" "%TARGET%"

:: Cleanup and exit
del "%VBS%" >nul 2>&1
exit /b 0

:: Body of VBS script
##>Option Explicit
##>Dim fso, folderPath
##>folderPath = WScript.Arguments(0)
##>Set fso = CreateObject("Scripting.FileSystemObject")
##>
##>Dim totalDocs, alreadyDone
##>totalDocs = 0
##>alreadyDone = 0
##>Call CountDocs(fso.GetFolder(folderPath))
##>Dim msg
##>If totalDocs = 0 Then
##>    msg = "DONE!"
##>Else
##>    msg = totalDocs & " doc(x) to convert"
##>End If
##>
##>If alreadyDone > 0 Then msg = msg & " -- " & alreadyDone & " pdf already done"
##>WScript.Echo msg
##>
##>If Not fso.FolderExists(folderPath) Then
##>    WScript.Echo "Folder does not exist: " & folderPath & vbCrLf & vbCrLf & "DONE (failed)"
##>    WScript.Sleep 3000
##>    WScript.Quit 1
##>End If
##>Call ConvertFolder(fso.GetFolder(folderPath))
##>WScript.Echo vbCrLf & "DONE!" & vbCrLf
##>WScript.Sleep 3000
##>WScript.Quit 0
##>
##>Sub CountDocs(folder)
##>    Dim file, subFolder, ext, pdfPath
##>    For Each file In folder.Files
##>        ext = LCase(fso.GetExtensionName(file.Name))
##>        If ext = "doc" Or ext = "docx" Then
##>            pdfPath = fso.GetParentFolderName(file.Path) & "\" & fso.GetBaseName(file.Name) & ".pdf"
##>            If fso.FileExists(pdfPath) Then
##>                alreadyDone = alreadyDone + 1
##>            Else
##>                totalDocs = totalDocs + 1
##>            End If
##>        End If
##>    Next
##>    For Each subFolder In folder.SubFolders
##>        CountDocs subFolder
##>    Next
##>End Sub
##>
##>Sub ConvertFolder(folder)
##>    Dim file, subFolder
##>    For Each file In folder.Files
##>        Dim ext : ext = LCase(fso.GetExtensionName(file.Name))
##>        If ext = "doc" Or ext = "docx" Then ConvertDocToPDF file.Path
##>    Next
##>    For Each subFolder In folder.SubFolders
##>        Call ConvertFolder(subFolder)
##>    Next
##>End Sub
##>
##>Sub ConvertDocToPDF(docPath)
##>    On Error Resume Next
##>    Dim word, doc, pdfPath, relPDF
##>    ' Build absolute PDF path
##>    pdfPath = fso.GetParentFolderName(docPath) & "\" & fso.GetBaseName(docPath) & ".pdf"
##>    ' Compute relative path
##>    relPDF = Mid(pdfPath, Len(folderPath & "\") + 1)
##>    ' Skip existing files
##>    If fso.FileExists(pdfPath) Then
##>        WScript.Echo "  Skipped:  " & relPDF
##>        Exit Sub
##>    End If
##>    ' Open Word in background
##>    Set word = CreateObject("Word.Application")
##>    word.Visible = False
##>    word.DisplayAlerts = 0
##>    Set doc = word.Documents.Open(docPath, False, True)
##>    If Err.Number <> 0 Then
##>        WScript.Echo "  Failed to open: " & docPath
##>        Err.Clear
##>        word.Quit
##>        Exit Sub
##>    End If
##>    ' Export in PDF
##>    doc.ExportAsFixedFormat pdfPath, 17
##>    doc.Close False
##>    word.Quit
##>    ' Print relative path
##>    WScript.Echo "  Created:  " & relPDF
##>    Set doc = Nothing
##>    Set word = Nothing
##>End Sub
