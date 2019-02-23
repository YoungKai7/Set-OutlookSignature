#region FUNCTIONS ##########
# Setup the .config file for configurable options
function Set-ConfigFile {
    # Determine configuration file name (same file name with .exe but with .config file extension)
    if ($script:MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
        $FileName=$script:MyInvocation.MyCommand.Definition
        $ConfigFileName=[io.path]::ChangeExtension($FileName, ".config")
    } else {
        $FileName=([Environment]::GetCommandLineArgs()[0])
        $ConfigFileName=[io.path]::ChangeExtension($FileName, ".config")
    }
    # Create the configuration file if it doesn't exist
    If(!(Test-Path $ConfigFileName -PathType Leaf)) {
        $ConfigData = @'
<config>
    <main>
        <UserName>$env:username</UserName>
        <SigSource>.\signature_data</SigSource><!-- path to folder containing the signature template(s) and Excel user profiles -->
        <UserSource>Excel</UserSource><!-- "Excel" or "AD" -->
        <UserSourceFile>$SigSource\UserDirectory.xlsx</UserSourceFile><!-- only for Excel user source -->
    </main>
    <profile>
        <!--
            Values can reference user profile as $($User) if needed. Ex. $($User.company)

            Different templates can be used based on Company and/or Department.
            Save template file into subfolders and name the subfolders according to the Company
            or Department names.

            Template priority: Department > Company > Default
        -->
        <Company>$env:userdomain</Company>
        <Department>$($User.Department)</Department>
        <regKey comment="Windows registry key where update history will be saved to"
            >$env:userdomain</regKey>
        <signatureName comment="Outlook signature name and user signation file names"
            >$env:userdomain (AUTO-SIG)</signatureName>
        <TemplateName comment="name of Word doc signature template"
            >Unified-Signature.docx</TemplateName>
        <ForceSignatureNew comment="Set as default signature for new messages. 0 = Not Set, 1 = Set"
            >1</ForceSignatureNew>
        <ForceSignatureReplyForward comment="Set as default signature for reply and forward messages. 0 = Not Set, 1 = Set"
            >1</ForceSignatureReplyForward>
    </profile>
    <customProfile>
        <!--
            Add custom nodes here to use in signature template.
            Values can be static or it can reference user profile as $($User) if needed. Ex. $($User.Company)
            
            Name of each node here should match the user property name.
            i.e. User Property name in AD, or column heading in Excel.
            The values used from default config matches standard AD user properties.
        -->
        <CompanyName>$($User.Company)</CompanyName>
        <FirstName>$($User.FirstName)</FirstName>
        <LastName>$($User.LastName)</LastName>
        <FirstLastName>$($User.FirstName) $($User.LastName)</FirstLastName>
        <FullName>$($User.DisplayName)</FullName>
        <Title>$($User.Title)</Title>
        <Telephone>$($User.TelephoneNumber)</Telephone>
        <Mobile>$($User.Mobile)</Mobile>
        <Email>$($User.Mail)</Email>
        <DepartmentName>$($User.Department)</DepartmentName>
    </customProfile>
</config>
'@
        $ConfigData | out-file $ConfigFileName
        attrib +s +h $ConfigFileName

        Write-Host @"
THANK YOU FOR USING SET-OUTLOOKSIGNATURE

It seems like this is your first time using Set-OutlookSignature.  Instructions are available using the "-help" switch:
        $FileName -Help

A config file has been created for you:
        $ConfigFileName

The config file will now open.  Please make necessary changes then run Set-OutlookSignature again.
"@
        Invoke-Item $ConfigFileName
        exit
    }
     $script:ConfigFileName = $ConfigFileName
}

# Helper: Import-Xls function makes it easy to work with Excel
function Import-Xls {
    <#
    .SYNOPSIS
    Import an Excel file.
    
    .DESCRIPTION
    Import an excel file. Since Excel files can have multiple worksheets, you can specify the worksheet you want to import. You can specify it by number (1, 2, 3) or by name (Sheet1, Sheet2, Sheet3). Imports Worksheet 1 by default.
    
    .PARAMETER Path
    Specifies the path to the Excel file to import. You can also pipe a path to Import-Xls.
    
    .PARAMETER Worksheet
    Specifies the worksheet to import in the Excel file. You can specify it by name or by number. The default is 1.
    Note: Charts don't count as worksheets, so they don't affect the Worksheet numbers.
    
    .INPUTS
    System.String
    
    .OUTPUTS
    Object
    
    .EXAMPLE
    ".\employees.xlsx" | Import-Xls -Worksheet 1
    Import Worksheet 1 from employees.xlsx
    
    .EXAMPLE
    ".\employees.xlsx" | Import-Xls -Worksheet "Sheet2"
    Import Worksheet "Sheet2" from employees.xlsx
    
    .EXAMPLE
    ".\deptA.xslx", ".\deptB.xlsx" | Import-Xls -Worksheet 3
    Import Worksheet 3 from deptA.xlsx and deptB.xlsx.
    Make sure that the worksheets have the same headers, or have some headers in common, or that it works the way you expect.
    
    .EXAMPLE
    Get-ChildItem *.xlsx | Import-Xls -Worksheet "Employees"
    Import Worksheet "Employees" from all .xlsx files in the current directory.
    Make sure that the worksheets have the same headers, or have some headers in common, or that it works the way you expect.
    
    .LINK
    Import-Xls
    http://gallery.technet.microsoft.com/scriptcenter/17bcabe7-322a-43d3-9a27-f3f96618c74b
    Export-Xls
    http://gallery.technet.microsoft.com/scriptcenter/d41565f1-37ef-43cb-9462-a08cd5a610e2
    Import-Csv
    Export-Csv
    
    .NOTES
    Author: Francis de la Cerna
    Created: 2011-03-27
    Modified: 2011-04-09
    #Requires �Version 2.0
    #>
    
    [CmdletBinding(SupportsShouldProcess=$true)]

    Param(
        [parameter(
            mandatory=$true,
            position=1,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $Path,

        [parameter(mandatory=$false)]
        $Worksheet = 1,

        [parameter(mandatory=$false)] 
        [switch]
        $Force
    )

    Begin
    {
        function GetTempFileName($extension)
        {
            $temp = [io.path]::GetTempFileName();
            $params = @{
                Path = $temp;
                Destination = $temp + $extension;
                Confirm = $false;
                Verbose = $VerbosePreference;
            }
            Move-Item @params;
            $temp += $extension;
            return $temp;
        }

        # since an extension like .xls can have multiple formats, this
        # will need to be changed
        #
        $xlFileFormats = @{
            # single worksheet formats
            '.csv'  = 6;        # 6, 22, 23, 24
            '.dbf'  = 11;       # 7, 8, 11
            '.dif'  = 9;        # 
            '.prn'  = 36;       # 
            '.slk'  = 2;        # 2, 10
            '.wk1'  = 31;       # 5, 30, 31
            '.wk3'  = 32;       # 15, 32
            '.wk4'  = 38;       # 
            '.wks'  = 4;        # 
            '.xlw'  = 35;       # 
                
            # multiple worksheet formats
            '.xls'  = -4143;    # -4143, 1, 16, 18, 29, 33, 39, 43
            '.xlsb' = 50;       # 
            '.xlsm' = 52;       # 
            '.xlsx' = 51;       # 
            '.xml'  = 46;       # 
            '.ods'  = 60;       # 
        }

        $xl = New-Object -ComObject Excel.Application;
        $xl.DisplayAlerts = $false;
        $xl.Visible = $false;
    }

    Process
    {
        $Path | ForEach-Object {

            if ($Force -or $psCmdlet.ShouldProcess($_)) {

                $fileExist = Test-Path $_
    
                if (-not $fileExist) {
                    Write-Error "Error: $_ does not exist" -Category ResourceUnavailable;
                } else {
                    # create temporary .csv file from excel file and import .csv
                    # 
                    $_ = (Resolve-Path $_).toString();
                    $wb = $xl.Workbooks.Add($_);
                    if ($?) {
                        $csvTemp = GetTempFileName(".csv");
                        $ws = $wb.Worksheets.Item($Worksheet);
                        $ws.SaveAs($csvTemp, $xlFileFormats[".csv"]);

                        $wb.Close($false);
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | out-null
                        Remove-Variable -Name ('ws', 'wb') -Confirm:$false;

                        Import-Csv $csvTemp;
                        Remove-Item $csvTemp -Confirm:$false -Verbose:$VerbosePreference;
                    }
                }
            }
        }
    }

    End
    {
        $xl.Quit();

        [gc]::Collect();
        [System.GC]::WaitForPendingFinalizers()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | out-null

        Remove-Variable -name xl -Confirm:$false;
    }
}

# Main function
function Main {
    # Load .config file
    Write-Progress -Activity 'Loading config file'
    Set-ConfigFile
    $ConfigFile=[xml](Get-Content $ConfigFileName)

    # Determine source of user profile from config
    $ConfigFile.SelectNodes('//config/main/*') | ForEach-Object {
        Write-Progress -Activity 'Loading config file' -Status "//config/main/$($_.Name)"
        Set-Variable -Name $_.Name -Value $ExecutionContext.InvokeCommand.ExpandString($_.'#text')
    }

    # Get user profile information for current user
    if($UserSource -eq "AD"){
        Write-Progress -Activity 'Loading User' -Status "Locating user from $UserSource"
        $Filter="(&(objectCategory=User)(samAccountName=$UserName))"
        $Searcher=New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.Filter=$Filter
        $ADUserPath=$Searcher.FindOne()
        $User=$ADUserPath.GetDirectoryEntry()
    } elseif($UserSource -eq "Excel") {
        Write-Progress -Activity 'Loading User' -Status "Locating user from $UserSourceFile"
        $UserDirectory=Import-XLS $UserSourceFile
        $User=$UserDirectory | Where-Object {($_.Username -eq $UserName)}
    } else {
        Throw "Invalid UserSource specified."
    }
    if (!$User) {
        Throw "Unable to find user '$UserName' from $UserSource."
    }
    
    # Retrieve rest of variables from .config file.
    # This section follows setting up user profile, so that these varaiables can reference the user profile if needed.
    $ConfigFile.SelectNodes('//config/profile/*') | ForEach-Object {
        Write-Progress -Activity 'Loading config file' -Status "//config/profile/$($_.Name)"
        Set-Variable -Name $_.Name -Value $ExecutionContext.InvokeCommand.ExpandString($_.'#text')
    }

    # Environment variables 
    $AppData=(Get-Item env:appdata).value 
    $SigPath="\Microsoft\Signatures"
    $LocalSignaturePath=$AppData+$SigPath 
    
    Write-Progress -Activity 'Locating signature template file'
    If(Test-Path ($SigSource+"\"+$Department+"\"+$TemplateName) -PathType Leaf) {
        $SigSource=$SigSource+"\"+$Department
    } elseif (Test-Path ($SigSource+"\"+$Company+"\"+$TemplateName) -PathType Leaf) {
        $SigSource=$SigSource+"\"+$Company
    } elseif (Test-Path ($SigSource+"\"+$TemplateName) -PathType Leaf) {
        $SigSource=$SigSource
    } else {
        Throw "'$TemplateName' not found."
    }

    $RemoteSignaturePathFull=$SigSource+"\"+$TemplateName
    $LocalSignaturePathFull=$LocalSignaturePath+'\'+$signatureName+'.docx'

    # Setting registry information for the current user
    Write-Progress -Activity 'Configuring Windows registry'
    $CompanyRegPath="HKCU:\Software\"+$regKey 

    if(Test-Path $CompanyRegPath){} 
    else
    {New-Item -path "HKCU:\Software" -name $regKey >$null}

    if(Test-Path $CompanyRegPath'\Outlook Signature Settings'){}
    else
    {New-Item -path $CompanyRegPath -name "Outlook Signature Settings" >$null}

    $SigVersion=(Get-ChildItem $RemoteSignaturePathFull).LastWriteTime.ToString("yyyyMMddTHHmmssffff") #Signature template file's last modified date.
    $ForcedSignatureNew=(Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').ForcedSignatureNew
    $ForcedSignatureReplyForward=(Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').ForcedSignatureReplyForward
    $rSigVersion=(Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').rSigVersion  #Signature version recorded from the last time this script was run.
    
    if($UserSource -eq "Excel"){
        $UserSourceFileVersion=(Get-ChildItem $UserSourceFile).LastWriteTime.ToString("yyyyMMddTHHmmssffff")
        $rUserSourceFileVersion=(Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').rUserSourceFileVersion
    }

    # Forcing signature for new messages if enabled
    if(([string]::IsNullOrEmpty($ForcedSignatureNew)) -or ($ForceSignatureNew -eq 1)){
        Write-Progress -Activity 'Updating Outlook settings' -Status 'Assign signature for new emails'
        $MSWord=New-Object -com word.application
        $EmailOptions=$MSWord.EmailOptions
        $EmailSignature=$EmailOptions.EmailSignature
        $EmailSignature.NewMessageSignature=$signatureName
        $MSWord.Quit()
        Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name ForcedSignatureNew -Value $ForceSignatureNew
    }
    
    # Forcing signature for reply/forward messages if enabled
    if(([string]::IsNullOrEmpty($ForcedSignatureReplyForward)) -or ($ForceSignatureReplyForward -eq 1)){
        Write-Progress -Activity 'Updating Outlook settings' -Status 'Assign signature for reply and forward emails'
        $MSWord=New-Object -com word.application
        $EmailOptions=$MSWord.EmailOptions
        $EmailSignature=$EmailOptions.EmailSignature
        $EmailSignature.ReplyMessageSignature=$signatureName
        $MSWord.Quit()
        Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name ForcedSignatureReplyForward -Value $ForceSignatureReplyForward
    }

    # Copying signature sourcefiles and creating signature if signature-version are different from local version
    Write-Progress -Activity 'Determining if updated signature is available'
    if ($forceupdate) {
        $UpdateCondition=$True
    } elseif ($UserSource -eq "Excel") {
        $UpdateCondition = (!(Test-Path $LocalSignaturePathFull)) -or (($rSigVersion -ne $SigVersion) -or ($rUserSourceFileVersion -ne $UserSourceFileVersion))
    } elseif ($UserSource -eq "AD") {
        $UpdateCondition = (!(Test-Path $LocalSignaturePathFull)) -or ($rSigVersion -ne $SigVersion)
    }
    if($UpdateCondition){
        # Copy signature templates to local signature folder
        Write-Progress -Activity 'Configuring signature' -Status 'Copying from template'
        New-Item -ItemType Directory -Force -Path "$LocalSignaturePath" | Out-Null
        Copy-Item "$RemoteSignaturePathFull" "$LocalSignaturePathFull" -Recurse -Force
    
        $ReplaceAll=2 
        $FindContinue=1 
        $MatchCase=$False 
        $MatchWholeWord=$True 
        $MatchWildcards=$False 
        $MatchSoundsLike=$False 
        $MatchAllWordForms=$False 
        $Forward=$True 
        $Wrap=$FindContinue 
        $Format=$False 
    
        # Replace customProfile variables from local signature template
        Write-Progress -Activity 'Configuring signature' -Status 'Opening Word template'
        $MSWord=New-Object -com word.application 
        $MSWord.Documents.Open("$LocalSignaturePathFull") | Out-Null
        
        $ConfigFile.SelectNodes('//config/customProfile/*') | ForEach-Object {
            Write-Progress -Activity 'Configuring signature' -Status "Replacing variables in Word template: $($_.Name)"
            $FindText='[['+$_.Name+']]'
            Set-Variable -Name ReplaceText -Value $ExecutionContext.InvokeCommand.ExpandString($_.'#text')

            $MSWord.Selection.Find.Execute($FindText,$MatchCase,$MatchWholeWord,$MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,$Wrap,$Format,$ReplaceText,$ReplaceAll) | Out-Null

            $MSWord.ActiveDocument.Hyperlinks | ForEach-Object {
                if ($_.Address -like "*$FindText*") {
                    $_.Address = $_.Address -replace [regex]::Escape($FindText),$ReplaceText
                }
            }
            
        }

        Write-Progress -Activity 'Configuring signature' -Status 'Save Word template'
        $MSWord.ActiveDocument.Save()
        
        # Fixes Enumeration Problems 
        $wdTypes = Add-Type -AssemblyName 'Microsoft.Office.Interop.Word' -Passthru 
        $wdSaveFormat = $wdTypes | Where-Object {$_.Name -eq "wdSaveFormat"} 
        
        # Save HTML 
        Write-Progress -Activity 'Configuring signature' -Status 'Save as HTML'
        $MSWord.ActiveDocument.saveas([ref]($LocalSignaturePath+'\'+$signatureName+'.htm'), [ref]$wdSaveFormat::wdFormatHTML);
        $MSWord.ActiveDocument.saveas(($LocalSignaturePath+'\'+$signatureName+'.htm'), $wdSaveFormat::wdFormatHTML);
        
        # Save RTF  
        Write-Progress -Activity 'Configuring signature' -Status 'Save as RTF'
        $MSWord.ActiveDocument.SaveAs([ref]($LocalSignaturePath+'\'+$signatureName+'.rtf'), [ref]$wdSaveFormat::wdFormatRTF);
        $MSWord.ActiveDocument.SaveAs(($LocalSignaturePath+'\'+$signatureName+'.rtf'), $wdSaveFormat::wdFormatRTF);
        
        # Save TXT     
        Write-Progress -Activity 'Configuring signature' -Status 'Save as TXT'
        $MSWord.ActiveDocument.SaveAs([ref]($LocalSignaturePath+'\'+$signatureName+'.txt'), [ref]$wdSaveFormat::wdFormatText);
        $MSWord.ActiveDocument.SaveAs(($LocalSignaturePath+'\'+$signatureName+'.txt'), $wdSaveFormat::wdFormatText);
        
        Write-Progress -Activity 'Configuring signature' -Status 'Closing Microsoft Word'
        $MSWord.ActiveDocument.Close();
        $MSWord.Quit();

        Write-Progress -Activity 'Configuring signature' -Status 'Record updates signature version to Windows registry'
        Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name SignatureSourceFile -Value $RemoteSignaturePathFull
        Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name rSigVersion -Value $SigVersion
        Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name rUserSourceFileVersion -Value $UserSourceFileVersion

        # Complete and exit
        Write-Progress -Activity 'Set-OutlookSignature' -Status 'Signature updated.' -Completed
        if (!$silent) { Write-Host "Signature updated." }
        exit
    } else {
        # Complete and exit
        Write-Progress -Activity 'Set-OutlookSignature' -Status 'No updates.' -Completed
        if (!$silent) { Write-Host "Already using latest signature. Signature update is not needed." }
        exit
    }    
}

# Parameter handling and Comment-Based Help
function Set-OutlookSignature {
<#
    .SYNOPSIS
        Configures user's Outlook signature using a Word document
        template.

    .DESCRIPTION
        Create/update Outlook signature based on user information from
        Active Directory or Excel file, using a Word document as template.
        
        Supported options are editable in a hidden .config system file.
        Use the -config parameter to open config file.
        
        Word document signature template will support variables defined
        under <customProfile> node on the config file.
        Enclose the variable with double brackets to on the Word signature
        template file.

        Examples:
            [[FullName]], [[Title]], [[Telephone]], [[Email]]

        Multiple templates are supported by placing the template file in
        subfolders under SigSource. Name the subfolders using Company
        and/or Department.
        Priority: Department > Company > Defalt
        Note: Define "Company" and "Department" in config file.

    .PARAMETER Help
        Outputs help text.

    .PARAMETER Config
        Open config file.

    .PARAMETER Silent
        Surpresses completion messages, however does not eliminate error
        messages. 

    .PARAMETER ForceUpdate
        Skips version difference checksum and forces signature files to be
        updated. 

    .EXAMPLE
        C:> Set-OutlookSignature.exe -help -detailed
        C:> Set-OutlookSignature.exe -config
        C:> Set-OutlookSignature.exe -forceupdate -silent
        C:> Set-OutlookSignature.exe -extract:C:\Output.ps1

    .COMPONENT
        This is a Powershell script (.ps1) written based on
        Set-OutlookSingature.ps1 v1.2 authored by: Jan Egil Ring, Darren
        Kattan, and Michael West.
            http://blog.powershell.no/2010/01/09/outlook-signature-based-on-user-information-from-active-directory  
            http://www.immense.net/deploying-unified-email-signature-template-outlook

        It incorporates Import-Xls function by Francis de la Cerna, and
        is compiled with PS2EXE-GUI by Markus Scholtes which can be
        decompiled with the "-extract:'<filename>'" switch.
    
    .NOTES
        3.0.0 02.23.2019 - Modified by YoungKai7
        - Repackaged for Github upload
#>

    param([switch]$help,[switch]$config,[switch]$silent,[switch]$forceupdate)

    if ($help) {
        # Ensure the .config file is available for user to follow help instructions
        Set-ConfigFile
    
        # Hyphens in parameters passed through Cmd causes Cmd to insert a 'True' element for each hyphen. These elements needs to be eliminated.
        # Ex. Input paramenter: "-help -examples"
        #     POSH outputs: "-help -examples"
        #     Cmd  outputs: "-help True -examples True"
        $prm=$args | Where-Object { $_ -ne "True" }
    
        $cmd = "& Get-Help Set-OutlookSignature $prm"
        $HelpOutput = Invoke-Expression $cmd | Out-String
        Write-Output $HelpOutput
    } elseif ($config) {
        # Setup the .config file and open it
        Set-ConfigFile
        Invoke-Item $ConfigFileName
    } else {
        if ($silent) {
            $ProgressPreference = "SilentlyContinue"
        }
        Write-Progress -Activity 'Loading Set-OutlookSignature'
        Main
    }
}

#endregion ##########

#region STARTUP ##########
Set-OutlookSignature @args
#endregion ##########