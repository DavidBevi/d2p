# d2p - doc 2 pdf
### Bulk-convert .doc(x) to .pdf files, from a folder and its subfolders.
Results are saved in their original locations. *Microsoft Word is required.*

-----

# [v2] [Download `d2p-install.bat`](https://github.com/DavidBevi/d2p/releases/download/v2/d2p-install.bat)
**Just open it as Admin & follow instructions**. Use this for maximum compatibility. It's less pretty (no colors) but better in any other way.

![image](https://github.com/DavidBevi/d2p/blob/main/d2p-v2-demo.gif?raw=true)

<details>
<summary>Alternatively, install manually:</summary>
 
> 1. download the script
> 2. rename it `d2p.bat`
> 3. either:
> - put it into a folder [listed into user-path or system-path](https://www.thewindowsclub.com/system-user-environment-variables-windows) (like C:\Windows\System32)
> - insert the folder with d2p into the user-path or system-path, and reboot (or relog)
</details>

-----

# [v1] Powershell script

![image](https://github.com/DavidBevi/d2p/blob/main/d2p-v1-demo.gif?raw=true)

**This is the original d2p, for Powershell**, but the install instructions I originally wrote are incomplete, and I discovered that Powershell is not always included in Windows, so I made the **Cmd version**. 

I recommend **against** using the Powershell version, because I abandoned it.

<details>
<summary>Oiginal install instructions:</summary>
  
> **INSTALLATION**
> - Download [`d2p.psm1`](https://github.com/DavidBevi/d2p/releases/download/v1/d2p.psm1) and save it into `C:\Program Files\WindowsPowerShell\Modules\d2p`.
> 
> **USAGE**
> - Open Powershell and type `d2p ` followed by the folder you want to process.
>   - EXAMPLE: `d2p C:\MyDocuments`
</details>

-----

# Credits
- `d2p.psm1` (v1) is based on [`ConvertWordTo-PDF`](https://sid-500.com/2020/10/20/powershell-convert-word-documentes-to-pdf-documents/) by Patrick Gruenauer | https://sid-500.com
- `d2p.bat` (v2) is a remake, with original code
