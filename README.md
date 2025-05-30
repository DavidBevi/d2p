# d2p - doc 2 pdf
### Bulk-convert .doc(x) to .pdf files, from a folder and its subfolders.
Results are saved in their original locations. *Microsoft Word is required.*

-----

### In the beginning there was the Powershell script.
But trying to install it on a new machine was painful, so I made a Cmd version that should be more compatible

# v2 - d2p-install.bat
Use this for maximum compatibility. It's less pretty (no colors) but better in any other way.
1. Download (link coming soon)
2. Open it as Admin (it will guide you)

Alternatively, dowload the script, put it in `<YourDriveLetter>:\Windows\System32\d2p.bat`, ensure you have %SystemRoot%\System32 in your system-path or user-path (if you need to add it reboot, or log-off-and-on, to enable "d2p" command)


| (v1) Powershell - [d2p.psm1](https://github.com/DavidBevi/d2p/releases/download/v1/d2p.psm1) | (v2) Cmd - d2p.bat |
|-|-|
| Has colors (prettier)| Monochrome|
| Requires Powershell| Doesn't require Powershell|
| I can't remember or reconstruct how I got it installed| Just double-click on the script, it will guide you in installing|
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
