# d2p - doc 2 pdf
### Bulk-convert .doc(x) to .pdf files, from a folder and its subfolders.
Results are saved in their original locations. *Microsoft Word is required.*

-----

# [v2] [Download `d2p-install.bat`](https://github.com/DavidBevi/d2p/releases/download/v2/d2p-install.bat)
**Just open it as Admin & follow instructions**. Use this for maximum compatibility. It's less pretty (no colors) but better in any other way.

![image](https://github.com/DavidBevi/d2p/blob/main/d2p-v2-demo.gif?raw=true)

<details>
<summary>Alternatively, install manually:</summary>
 
> - dowload the script
> - put it in `<YourDriveLetter>:\Windows\System32\d2p.bat`
> - ensure you have %SystemRoot%\System32 in your system-path or user-path
>   - if you need to add it then ensure to reboot (or log-off-and-on) to enable "d2p" command
</details>

-----

# [v1] Powershell script

![image](https://github.com/DavidBevi/d2p/blob/main/d2p-v1-demo.gif?raw=true)

**This is the original d2p, for Powershell**, but trying to install it on a new machine was so clunky that I gave up.

I recommend the **Cmd version**. I recommend **against** using the Powershell version, because I abandoned it.

<details>
<summary>Here's the original instructions:</summary>
  
> **INSTALLATION**
> - Download [`d2p.psm1`](https://github.com/DavidBevi/d2p/releases/download/v1/d2p.psm1) and save it into `C:\Program Files\WindowsPowerShell\Modules\d2p`.
> 
> **USAGE**
> - Open Powershell and type `d2p ` followed by the folder you want to process.
>   - EXAMPLE: `d2p C:\MyDocuments`
</details>

-----

# Credits
`d2p by DavidBevi` is based on [`ConvertWordTo-PDF`](https://sid-500.com/2020/10/20/powershell-convert-word-documentes-to-pdf-documents/) by Patrick Gruenauer | https://sid-500.com
