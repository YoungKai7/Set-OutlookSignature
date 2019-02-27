# Set-OutlookSignature
Obtains user's profile from Active Directory or Excel, and configures user's Outlook signature using a Word document template.

<p align="center">
  <a href="https://raw.githubusercontent.com/YoungKai7/Set-OutlookSignature/assets/demo-template.png">
    <img src="https://raw.githubusercontent.com/YoungKai7/Set-OutlookSignature/assets/demo-template.png" alt="Set-OutlookSignature Demo">
  </a>
</p>
<p align="center">
  <a href="https://raw.githubusercontent.com/YoungKai7/Set-OutlookSignature/assets/demo-gif.gif">
    <img src="https://raw.githubusercontent.com/YoungKai7/Set-OutlookSignature/assets/demo-gif.gif" alt="Set-OutlookSignature Demo">
  </a>
</p>

### Features
- [x] Obtain user profile from AD or Excel.
- [x] Use MS Word document to create signature templates.
- [x] Customizable variables (e.g. `[[FirstLastName]]`, `[[myVariable]]`) can be used as text and part of hyperlink on the Word signature template(s).
- [x] Supports multiple templates.
- [x] Silent execution option for automated deployment.
- [x] Executable for manual deployment, so users don't have to work with script execution commands.
- [x] Shorten execution time to skip signature update if signature template has not been changed, or use the `-forceupdate` switch to ensure signature standard.
- [x] Automatically selects signature for new messages/replies/forwards.


### User Friendly
| **END-USERS** |
| :--- |
| Automated process makes it transparent to the end-users. |
| Alternatively, if process automation the process through GPO is not an option, user can simply double-click on an executable to set their Outlook signature. |

| **MARKETING** |
| :--- |
| Use of MS Word document to easily design signature templates.  Use different templates for different companies or departments to deliver audience focused messaging. |
| Ability to use Excel for user profiles grants more control without having to rely on I.T. to keep personnel information updated. |

| **I.T.** |
| :--- |
| Flexible deployment. Automate through GPO/log-on script to force the signature standard, and/or simply share the executable on a shared folder for users to manually execute. |
| All options can be easily configured through a .config file. |
| Easily add new variables to use on the MS Word signature template(s). |
| Detail commented script for easy customizations. |

## Instructions
`[Automated Deployment]` is meant to run `Set-OutlookSignature` automatically, so the process to set Outlook signature is transparent to the users.

`[Manual Deployment]` requires each user to manually run `Set-OutlookSignature` each time the user wishes to update their Outlook signature.

### Standard Usage
1. Download the [latest release](https://github.com/YoungKai7/Set-OutlookSignature/releases/latest/download/Set-OutlookSignature.zip)
2. Save extracted files:

    `[Automated Deployment]` Save to NETLOGON share.
    
    `[Manual Deployment]` Save to a network shared folder.
    
3. Run `Set-OutlookSignature.exe` once and following on-screen instructions. This should guide you through editing the .config file.
4. Edit the signature template (`Unified-Signature.docx`)
5. To run:
    
    `[Automated Deployment]` Configure GPO and/or user logon script to run `Set-OutlookSignature.exe -silent`
    
    `[Manual Deployment]` Instruct users to run `Set-OutlookSignature.exe`
    
### Advanced Usage
Run `Set-OutlookSignature.exe -help` to learn more.

Run `Set-OutlookSignature.exe -help -detailed` to learn even more.

### Customizations
1. Download the source code.
2. Make necessary changes.
3. Compile the script using [PS2EXE-GUI](https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-GUI-Convert-e7cb69d5)

## Notes
- The .config file will automatically generate with default values if no existing .config is found.
- The .config file must match the name of the executable (yes, you can rename `Set-OutlookSignature.exe`).
- You can combine `-Silent` and `-ForceUpdate` switches for automated enforced signature standard.  E.g. `Set-OutlookSignature.exe -silent -forceupdate`

## Components
This is a Powershell script (.ps1) written based on [`Set-OutlookSingature.ps1 v1.2`](https://gallery.technet.microsoft.com/office/Outlook-signature-based-on-8178d376) authored by Jan Egil Ring, Darren
Kattan, and Michael West.

<details><summary>...</summary>
<p>
Unfortunately I didn't discover [Jan's repo](https://github.com/janegilring/PSCommunity/blob/master/Microsoft%20Office/Set-OutlookSignature.ps1) until I had finished with my changes against Michael's v1.2 and as I'm writing these last words in README.md.  Else I could've branched off Jan's latest version instead of creating a new repo, saved some hassle, and gain couple more enhancements in the script.  This will do for now.
</p>
</details>
<br/>

It incorporates [Import-Xls](http://gallery.technet.microsoft.com/scriptcenter/17bcabe7-322a-43d3-9a27-f3f96618c74b) function by Francis de la Cerna, and was compiled using [PS2EXE-GUI](https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-GUI-Convert-e7cb69d5) by Markus Scholtes.
