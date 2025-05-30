## d2p - Bulk-convert Doc(x) 2 Pdf, from a folder and its subfolders.
Results are saved in their original locations. *Microsoft Word is required.*

-----

### In the beginning there was the Powershell script.
But trying to install it on a new machine was painful, so I made a Cmd version that should be more compatible

| (v1) Powershell - [d2p.psm1](https://github.com/DavidBevi/d2p/releases/download/v1/d2p.psm1) | (v2) Cmd - d2p.bat |
|-|-|
|✔️ Has colors (prettier)|❌ Monochrome|
|❌ Requires Powershell|✔️ Doesn't require Powershell|
|❌ I can't remember or reconstruct how I got it installed|✔️ Just double-click on the script, it will guide you in installing|
|![image](https://github.com/DavidBevi/d2p/blob/main/2dp_demo.gif?raw=true)| gif coming soon |

-----

# (v1) Powershell (v2 coming soon)

## INSTALLATION
- Download [`d2p.psm1`](https://github.com/DavidBevi/d2p/releases/download/v1/d2p.psm1) and save it into `C:\Program Files\WindowsPowerShell\Modules\d2p`.

## USAGE
- Open Powershell and type `d2p ` followed by the folder you want to process.
  - EXAMPLE: `d2p C:\MyDocuments`

### CREDITS 
- `d2p by DavidBevi` is based on [`ConvertWordTo-PDF`](https://sid-500.com/2020/10/20/powershell-convert-word-documentes-to-pdf-documents/) by Patrick Gruenauer | https://sid-500.com
