function d2p {

<# 

.SYNOPSIS 
  d2p converts DOC and DOCX files to PDF files.

.DESCRIPTION 
  The cmdlet queries the given source folder including sub-folders to find *.docx and *.doc files, 
  converts all found files and saves them as pdf in their original folder. 
  A progress bar will show the progress and the name of the file being converted.

.PARAMETER SourceFolder
  Mandatory. Enter the source folder of your Microsoft Word documents.

.EXAMPLE 
  d2p C:\temp
  d2p -SourceFolder C:\Temp

.NOTES 
  Based on ContertWordTo-PDF by Patrick Gruenauer
  Original author: Patrick Gruenauer | https://sid-500.com
  Modder: David Bevi | https://github.com/DavidBevi
  
#>

[CmdletBinding()]

param
(
 
[Parameter (Mandatory=$true,Position=0)]
[String]
$SourceFolder

)

    ''
    Write-Host " d2p - Bulk conversion of DOC(X) to PDF " -BackgroundColor White -ForegroundColor DarkCyan
    Write-Host "    https://github.com/DavidBevi/d2p    " -ForegroundColor DarkCyan

    $i = 0

    $word = New-Object -ComObject word.application 
    $FormatPDF = 17
    $word.visible = $false 
    $types = '*.docx','*.doc'

    If ((Test-Path $SourceFolder) -eq $false) {
        throw "Error. Source Folder $SourceFolder not found." }
    
    $files = Get-ChildItem -Path $SourceFolder -Include $Types -Recurse -ErrorAction Stop

    foreach ($f in $files) {
        $activity = 'Converting files to PDF: ' + ($i+1) + '/' + $files.Length
        Write-Progress -Activity $activity -Status $f.Name -PercentComplete (($i / $files.Length) * 100)

        $path = $f.FullName.Substring(0,($f.FullName.LastIndexOf('.')))
        $nam = $f.Name.Substring(0,($f.Name.LastIndexOf('.')))
        Write-Host ($nam + ".pdf ") -NoNewLine

        if (Test-Path -Path ($path + ".pdf")) {
          Write-Host ("already exists") -ForegroundColor DarkCyan
          $i++
          
        } else {
            try { 
              $doc = $word.documents.open($f.FullName)
              $doc.SaveAs($path,$FormatPDF)
              Write-Host ("created") -ForegroundColor DarkGreen
              $doc.Close()
              $i++
            } catch { 
              Write-Error ("Can't convert" + $f.Name)
            }
        }
    }
    $word.Quit()

    ''
    if ($i -eq $files.Length) {
        Write-Host " All files succesfully exported in PDF " -BackgroundColor DarkGreen
    } else {
        Write-Host (" Warning: " + ($files - $i) + " files NOT exported in PDF ") -BackgroundColor DarkYellow
    }
    ''
}
