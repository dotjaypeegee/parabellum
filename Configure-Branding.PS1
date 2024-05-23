    # =============================================================================================================================
#
# Script Name:     Configure-Branding.PS1
# Description:     This script creates a scheduled task, a Powershell script, a VBS file, copies modules(s) and file(s)
#                  The outputted script will manage wallpaper and lockscreen media as needed
#                  
#=============================================================================================================================


#################### App setup Here 
    #CustomerName
    $Cust = $env:USERDOMAIN
    #Vendorname
    $Vendor = #Support Team name here
    #Purpose - e.g: Set to Win32 App name
    $Purpose = 'WallpaperRefresh'
    $Version = '2.6'
    $Function = 'Install'
    #Name of the App being installed
    $Foldername = 'ScheduledTasks'
    $DestFolder = "$env:windir\$Cust\files\$Foldername\$Purpose"
    $NewOwner = "BUILTIN\Administrators"
    $SystemContext = "NT AUTHORITY\SYSTEM"
    $DefaultWallpaperImagePath = Join-Path -Path $env:WINDIR -ChildPath "WEB\Wallpaper\Windows\img0.jpg"
    $MSImages = Get-ChildItem -Path "C:\Windows\Web" -Filter "*.jpg" -Recurse
    $RootImages = Get-ChildItem -Path $DestFolder -Filter "*.jpg" -Recurse | Where-Object { $PSItem.Name -like "*.jpg" }   
    $LocalAdministratorsPrincipal = "BUILTIN\Administrators"
    $LocalUsersPrincipal = "BUILTIN\Users"
    $LocalSystemPrincipal = "NT AUTHORITY\SYSTEM"
    $TrustedInstallerPrincipal = "NT SERVICE\TrustedInstaller"
    $RestrictedApplicationPackagesPrincipal = "ALL RESTRICTED APPLICATION PACKAGES"
    $ApplicationPackagesPrincipal = "ALL APPLICATION PACKAGES"
    #Change the log name to match the Application being installed
    $logFilename = $Function + "-" + $Purpose + '.log'
    #Declare the Source Directory. Dont Dot Source.
    $SourceFolder = "$PSScriptRoot\"
    #TestLineonLocalPCbelow
    #$SourceFolder = "C:\temp\"
    $DetectionMethod = $DestFolder + "\" + $Function  + "-" + $Purpose + "_" + $version + ".ps1_tag"
    $SaveExecutionPolicy = Get-ExecutionPolicy
    Set-ExecutionPolicy bypass -Scope Process -Force -Confirm:$false




#################### Determine if a 32-bit process is used on a x64 system
    if ("$env:PROCESSOR_ARCHITEW6432" -ne "ARM64"){
        if (Test-Path "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe")
            { & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath"
                Exit $lastexitcode
        }
    }

#####################################################################################################

    function Write-LogEntry {
            param (
                [parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
                [ValidateNotNullOrEmpty()]
                [string]$Value,
    
                [parameter(Mandatory = $false, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
                [ValidateNotNullOrEmpty()]
                [ValidateSet("1", "2", "3")]
                [string]$Severity = "1",
    
                [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
                [ValidateNotNullOrEmpty()]
                [string]$FileName =  $logFilename
            )
            # Determine log file location
            #$LogFilePath = Join-Path -Path (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Temp") -ChildPath $FileName 
            $LogFilePath = Join-Path -Path (Join-Path -Path $env:windir -ChildPath "debug") -ChildPath $FileName
        
            # Construct time stamp for log entry
            $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), "+", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
        
            # Construct date for log entry
            $Date = (Get-Date -Format "MM-dd-yyyy")
        
            # Construct context for log entry
            $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
        
            # Construct final log entry
            $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($logFilename)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
        
            # Add value to log file
            try {
                Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -erroraction SilentlyContinue
            }
            catch [System.Exception] {
                Write-Warning -Message "Unable to append log entry to $($logFilename).log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
            }
        }

#####################################################################################################

#################### Create Folder for app

    if ( !(Test-Path "$DestFolder") ) {
        
        try {
            New-Item -Path "$DestFolder" -ItemType Directory -Force  -ErrorAction Stop 
            Write-LogEntry -Value "Created folder for $Purpose  " -Severity 1
        } 
        
        catch [exception]
            {     
            Write-LogEntry -Value "Failed to Create folder for   $Purpose  . Error is: $($_.Exception.Message))" -Severity 3
        } 
     }

##################### Copy Files for Wallpaper and LockScreen
# Currently used as WIn32 app with Webmod folder (Contains image media) and required modules see below
# Offload the media and initial module requirements to Blob storage in future iterations

   try {
        Copy-Item -Path "$($SourceFolder)WebMod" -Recurse  -Destination $DestFolder  -Force -ErrorAction SilentlyContinue
        Write-LogEntry -Value "Copied WebMod to $DestFolder  " -Severity 1
   }

   catch [System.Exception] {
        Write-LogEntry -Value "Failed to Copy WebMod to sys32 modules  . Error is: $($_.Exception.Message))" -Severity 3
   }

#####################################################################################################


    # Stage NTFSSecurity module

    try {
        $NTFSSecurityModulePath = Join-Path -Path $env:WINDIR -ChildPath "System32\WindowsPowerShell\v1.0\Modules\NTFSSecurity"
        
        if (-not(Test-Path -Path $NTFSSecurityModulePath -PathType Container)) {
        
            Copy-Item -Path (Join-Path -Path $SourceFolder -ChildPath "Modules") -Destination $NTFSSecurityModulePath -Recurse -ErrorAction Stop
            Write-LogEntry -Value "Copied NTFSSecurityModule to sys32 modules " -Severity 1
        }
    }
    catch [System.Exception] {
        Write-LogEntry -Value "Failed to Copy NTFSSecurityModule to sys32 modules  . Error is: $($_.Exception.Message))" -Severity 3
    }


#####################################################################################################


    # Import NTFSSecurity module
    try {
        Import-Module -Name NTFSSecurity -ErrorAction Stop
        Write-LogEntry -Value "imported NTFSSecurityModule " -Severity 1
    }
    catch [System.Exception] {
       Write-LogEntry -Value "Failed to import NTFSSecurityModule   . Error is: $($_.Exception.Message))" -Severity 3 
    }


#####################################################################################################



    Write-LogEntry -Value "working on  the Rootimage permission step" -Severity 1
    $RootImages = Get-ChildItem -Path $DestFolder -Filter "*.jpg" -Recurse | Where-Object { $PSItem.Name -like "*.jpg" } 
    $RootImagescount = ($RootImages | Measure-Object).Count
    Write-LogEntry -Value "working on  the $RootImagescount Root Images" -Severity 1
  
     foreach ($RootImage in $RootImages){    
 
       Write-LogEntry -Value "Amending the $RootImage Root Images" -Severity 1
       $MSImage = $MSImages | Where-Object { $PSItem.Name -like $RootImage.Name } 


             # if ((Get-FileHash $RootImage.fullname).hash  -ne (Get-FileHash $MSImage.fullname).hash){
                    try {
                            # Grant non-inherited permissions for wallpaper item
                            Write-LogEntry -Value "Granting '$($LocalSystemPrincipal)' Read and Execute on: '$($rootimage.name)'" -Severity 1
                            Add-NTFSAccess -Path $rootimage.fullname -Account $LocalSystemPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($LocalAdministratorsPrincipal)' Read and Execute on: $($rootimage.name)" -Severity 1
                            Add-NTFSAccess -Path $rootimage.fullname -Account $LocalAdministratorsPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($LocalUsersPrincipal)' Read and Execute on: $($rootimage.name)" -Severity 1
                            Add-NTFSAccess -Path $rootimage.fullname -Account $LocalUsersPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($ApplicationPackagesPrincipal)' Read and Execute on: $($rootimage.name)" -Severity 1
                            Add-NTFSAccess -Path $rootimage.fullname -Account $ApplicationPackagesPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($RestrictedApplicationPackagesPrincipal)' Read and Execute on: $($rootimage.name)" -Severity 1
                            Add-NTFSAccess -Path $rootimage.fullname -Account $RestrictedApplicationPackagesPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                           Write-LogEntry -Value "Granting '$($TrustedInstallerPrincipal)' Full Control on: $($rootimage.name)" -Severity 1
                            Add-NTFSAccess -Path $rootimage.fullname -Account $TrustedInstallerPrincipal -AccessRights "FullControl" -erroraction SilentlyContinue
                           Write-LogEntry -Value "Disabling inheritance on: $($RootImage.name)" -Severity 1
                            Disable-NTFSAccessInheritance -Path $RootImage.fullname -RemoveInheritedAccessRules -erroraction SilentlyContinue

                            try {
                                # Set owner to trusted installer for new wallpaper file
                                Write-LogEntry -Value "Setting ownership for '$($TrustedInstallerPrincipal)' on wallpaper image file: $($rootimage.fullname)" -Severity 1
                                Set-NTFSOwner -Path $rootimage.fullname -Account $TrustedInstallerPrincipal -erroraction SilentlyContinue
                            }
                            catch [System.Exception] {
                                Write-LogEntry -Value "Failed to set ownership for '$($TrustedInstallerPrincipal)' on wallpaper image file: $($rootimage.fullname). Error message: $($_.Exception.Message)" -Severity 3
                            }
                        }
                    catch [System.Exception] {
                            Write-LogEntry -Value "Failed to revert permissions for wallpaper image file. Error message: $($_.Exception.Message)" -Severity 3}
                   
                    try {
                        $CurrentOwner = Get-Item -Path $MSImage.fullname | Get-NTFSOwner
                        $CurrentOwnerAccountName = $CurrentOwner.Owner.AccountName
                        Write-LogEntry -Value "Owner of $MSImage.fullname is $CurrentOwnerAccountName "
    
                if ($CurrentOwner.Owner -notlike $NewOwner) {
                    Write-LogEntry -Value "Setting new owner of '$($NewOwner)' on: $($MSImage.fullname)"
                    Set-NTFSOwner -Path $MSImage.fullname -Account $NewOwner -erroraction SilentlyContinue }

            # Grant NT AUTHORITY\SYSTEM and Local Administrator Full Control access
                    try {
                       Write-LogEntry -Value "Granting '$($SystemContext)' Full Control on: $($MSImage.fullname)"
                        Add-NTFSAccess -Path $MSImage.fullname -Account $SystemContext -AccessRights FullControl -AccessType Allow -erroraction SilentlyContinue
                        Write-LogEntry -Value "Granting '$($NewOwner)' Full Control on: $($MSImage.fullname)"
                        Add-NTFSAccess -Path $MSImage.fullname -Account $NewOwner -AccessRights FullControl -AccessType Allow -erroraction SilentlyContinue
                        }
                    catch [System.Exception] {
                        Write-LogEntry -Value "Unable to grant required Full Control permissions on: $($MSImage.fullname). Error message: $($_.Exception.Message)" 
                        }

            # Replace wallpaper
                    try {
                        Write-LogEntry -Value "Replacing default media object in: $($MSImage.fullname)"
                        Remove-Item -Path $MSImage.fullname -Force -erroraction SilentlyContinue
                        Copy-Item -Path $RootImage.fullname -Destination $MSImage.fullname -Force -erroraction SilentlyContinue
                          }
                    catch [System.Exception] {
                        Write-LogEntry -Value "Unable to replace image: $($MSImage.fullname). Error message: $($_.Exception.Message)"  -Severity 3
                    }
                            }
                    catch [System.Exception] {
                            Write-LogEntry -Value "Unable to take ownership of: $($MSImage.fullname). Error message: $($_.Exception.Message)"  -Severity 3
                }
                }
                

#####################################################################################################


##########################################Define Script

$PowerSHellscript = @'
    #################### App setup Here 
    #CustomerName
    $Cust = #''
    #Vendorname
    $Vendor = #Support Team name here
    #Purpose - e.g: Set to Win32 App name
    $Version = '2.6'
    $Function = 'Manage'
    #Purpose - e.g: Set to Win32 App name
    $Purpose = 'WallpaperRefresh'
    $Foldername = 'ScheduledTasks'
    $DestFolder = "$env:windir\$cust\files\$Foldername\$Purpose"
    $NewOwner = "BUILTIN\Administrators"
    $SystemContext = "NT AUTHORITY\SYSTEM"
    #Initial declaration of this variable - Will change further down the page condition dependent...
    $DefaultWallpaperImagePath = Join-Path -Path $env:WINDIR -ChildPath "WEB\Wallpaper\Windows\img0.jpg"
    $DefaultScreenImagePath = Join-Path -Path $env:WINDIR -ChildPath "Web\Screen\img100.jpg"
    $MSImages = Get-ChildItem -Path "C:\Windows\Web" -Filter "*.jpg" -Recurse
    $RootImages = Get-ChildItem -Path $DestFolder -Filter "*.jpg" -Recurse | Where-Object { $PSItem.Name -like "*.jpg" }       
    $TranscodedWallpaperpath = "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Themes"
    $TranscodedWallpaperImage = "TranscodedWallpaper"
    $TranscodedWallpaperImagepath = $TranscodedWallpaperpath + "\" + $TranscodedWallpaperImage
    $LocalAdministratorsPrincipal = "BUILTIN\Administrators"
    $LocalUsersPrincipal = "BUILTIN\Users"
    $LocalSystemPrincipal = "NT AUTHORITY\SYSTEM"
    $TrustedInstallerPrincipal = "NT SERVICE\TrustedInstaller"
    $RestrictedApplicationPackagesPrincipal = "ALL RESTRICTED APPLICATION PACKAGES"
    $ApplicationPackagesPrincipal = "ALL APPLICATION PACKAGES"
    $User = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI).LastLoggedOnUser.Split(‘\’)[1] 
    $UserSid = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI).LastLoggedOnUserSID
    $RKP1 = "Microsoft.PowerShell.Core\Registry::HKEY_USERS\" + $Usersid + "\" + "Control Panel\Desktop"
    $RKP2 = "Microsoft.PowerShell.Core\Registry::HKEY_USERS\" + $Usersid + "\" + "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    $RKP3 = "Microsoft.PowerShell.Core\Registry::HKEY_USERS\" + $Usersid + "\" + "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop" 
    $RKP4 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
    $RKP5 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Creative"
    $RKP6 = "Microsoft.PowerShell.Core\Registry::HKEY_USERS\" + $Usersid + "\" + "Software\Microsoft\Windows\CurrentVersion\Lock Screen\Creative" 
    $RKP7 = "Microsoft.PowerShell.Core\Registry::HKEY_USERS\" + $Usersid + "\" + "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    $RKP8 = "Microsoft.PowerShell.Core\Registry::HKEY_USERS\" + $Usersid + "\" + "Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers"
    $logFilename = $Function + "-" + $Purpose + '.log'

############################################################ End set script variables


############################################################ LogFile Function
    function Write-LogEntry {
            param (
                [parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
                [ValidateNotNullOrEmpty()]
                [string]$Value,
    
                [parameter(Mandatory = $false, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
                [ValidateNotNullOrEmpty()]
                [ValidateSet("1", "2", "3")]
                [string]$Severity = "1",
    
                [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
                [ValidateNotNullOrEmpty()]
                [string]$FileName = $logFilename
            )
            # Determine log file location
            $LogFilePath = Join-Path -Path (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Temp") -ChildPath $FileName 
            #$LogFilePath = Join-Path -Path (Join-Path -Path $env:windir -ChildPath "debug") -ChildPath $FileName
        
            # Construct time stamp for log entry
            $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), "+", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
        
            # Construct date for log entry
            $Date = (Get-Date -Format "MM-dd-yyyy")
        
            # Construct context for log entry
            $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
        
            # Construct final log entry
            $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($logFilename)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
        
            # Add value to log file
            try {
                Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -erroraction SilentlyContinue
            }
            catch [System.Exception] {
                Write-Warning -Message "Unable to append log entry to $($logFilename) file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
            }
        }

############################################################ End LogFile Function
$SaveExecutionPolicy = Get-ExecutionPolicy
Set-ExecutionPolicy RemoteSigned -Scope Process -Force -Confirm:$false

############################################################ Determine run context 

function Test-IsSystem {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        return $id.Name -like "NT AUTHORITY*" -or $id.IsSystem
    }       

############################################################ End Determine run context 

############################################################ Get Latest NTFSSecurity module

   #TLS Protocol set
      [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
      $TLSSecurityProtocol = [Net.ServicePointManager]::SecurityProtocol
      Write-LogEntry -Value  "The TLS SecurityProtocol is set to $TLSSecurityProtocol"

      $Package = "NuGet"
      Write-LogEntry -Value  "The package being validated is $Package"

       try {
            $CurrentPackage = Get-PackageProvider -Name $Package -Force -erroraction SilentlyContinue -Verbose:$false
            $CurrentPackagename = $CurrentPackage.name
            $CurrentPackageVersion =   $CurrentPackage.Version 

               if ( $null -ne $CurrentPackage) {
                    Write-LogEntry -Value  "The $CurrentPackagename Package exists and uses version $CurrentPackageVersion  "                        
                    $LatestPackageVersion = (Find-PackageProvider -Name $Package -erroraction SilentlyContinue -Verbose:$false).Version
                            
                     if ($LatestPackageVersion -eq $CurrentPackage.Version){  
                         Write-LogEntry -Value  "The $CurrentPackagename Package uses the latest version." 
                     }          
                                              
                     elseif ($LatestPackageVersion -gt $CurrentPackage.Version) {
                         Write-LogEntry -Value  "The $CurrentPackagename Package exists but is being upgraded as version $LatestPackageVersion is available "
                 
                         try {
                             INstall-PackageProvider -Name $Package -Force -erroraction SilentlyContinue -Confirm:$false -Verbose:$false
                             import-PackageProvider  -Name $Package -Force -erroraction SilentlyContinue  -Verbose:$false
                         }                           
                         catch [System.Exception]{        
                             Write-LogEntry -Value  "An error occurred while attempting to install $($Package) Package. Error message: $($_.Exception.Message)" -Severity 3 
                         }                            
                     }
                            
               }
                            
               else{ Write-LogEntry -Value  "The $($Package) Package is not installed. Going to try to install it" -Severity 2
                     
                     try {   
                         INstall-PackageProvider -Name $Package -Force -erroraction SilentlyContinue -Confirm:$false -Verbose:$false
                          import-PackageProvider  -Name $Package -Force -erroraction SilentlyContinue  -Verbose:$false
                     }      
                     catch [System.Exception]{  
                         Write-LogEntry -Value  "An error occurred while attempting to install $($Package) Package. Error message: $($_.Exception.Message)" -Severity 3
                     }
               }
       }
       catch [System.Exception]{
              Write-LogEntry -Value  "The $($Package) package is not installed" -Severity 3
              Write-LogEntry -Value  "An error occurred while attempting to install $($Package) Package. Error message: $($_.Exception.Message)" -Severity 3
        }
    
 
      Write-LogEntry -Value  "The TLS SecurityProtocol is set to $TLSSecurityProtocol"
      $module = "NTFSSecurity"
      Write-LogEntry -Value  "The module being validated is $module"

       try {
            $Currentmodule = Get-InstalledModule -Name $Module -erroraction SilentlyContinue -Verbose:$false
            $Currentmodulename = $Currentmodule.name
            $CurrentmoduleVersion =   $Currentmodule.Version 

               if ( $null -ne $Currentmodule) {
                    Write-LogEntry -Value  "The $Currentmodulename module exists and uses version $CurrentmoduleVersion  "                        
                    $LatestmoduleVersion = (Find-Module -Name $Module -erroraction SilentlyContinue -Verbose:$false).Version
                            
                     if ($LatestmoduleVersion -eq $Currentmodule.Version){  
                         Write-LogEntry -Value  "The $Currentmodulename module uses the latest version." 
                     }          
                                              
                     elseif ($LatestmoduleVersion -gt $Currentmodule.Version) {
                         Write-LogEntry -Value  "The $Currentmodulename module exists but is being upgraded as version $LatestmoduleVersion is available "
                 
                         try {
                             INstall-Module -Name $Module -Force -erroraction SilentlyContinue -Confirm:$false -Verbose:$false
                             Import-Module -Name $module -Force -Scope Local -erroraction SilentlyContinue
                         }                           
                         catch [System.Exception]{        
                             Write-LogEntry -Value  "An error occurred while attempting to install $($module) module. Error message: $($_.Exception.Message)" -Severity 3 
                         }                            
                     }
                            
               }
                            
               else{ Write-LogEntry -Value  "The $($module) module is not installed. Going to try to install it" -Severity 2
                     
                     try {   
                          INstall-Module -Name $Module -Force -erroraction SilentlyContinue -Confirm:$false -Verbose:$false
                           Import-Module -Name $module -Force -Scope Local -erroraction SilentlyContinue      
                     }
                     catch [System.Exception]{  
                         Write-LogEntry -Value  "An error occurred while attempting to install $($module) module. Error message: $($_.Exception.Message)" -Severity 3 
                     }
               }
       }
       catch [System.Exception]{
              Write-LogEntry -Value  "The $($module) module is not installed" -Severity 3
              Write-LogEntry -Value  "An error occurred while attempting to install $($module) module. Error message: $($_.Exception.Message)" -Severity 3
       }            




############################################################ End get Latest NTFSSecurity module

############################################################ Process source and destination wallpaper permissions : Grant non-inherited permissions for the "new" wallpaper in its source location

    foreach ($RootImage in $RootImages){            
             $MSImage = $MSImages | Where-Object { $PSItem.Name -like $RootImage.Name }

              if ((Get-FileHash $RootImage.fullname).hash  -ne (Get-FileHash $MSImage.fullname).hash){
                    
                    try {
                            # Grant non-inherited permissions for wallpaper item
                            Write-LogEntry -Value "Granting '$($LocalSystemPrincipal)' Read and Execute on: $($RootImage)" -Severity 1
                            Add-NTFSAccess -Path $RootImage -Account $LocalSystemPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($LocalAdministratorsPrincipal)' Read and Execute on: $($RootImage)" -Severity 1
                            Add-NTFSAccess -Path $RootImage -Account $LocalAdministratorsPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($LocalUsersPrincipal)' Read and Execute on: $($RootImage)" -Severity 1
                            Add-NTFSAccess -Path $RootImage -Account $LocalUsersPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($ApplicationPackagesPrincipal)' Read and Execute on: $($RootImage)" -Severity 1
                            Add-NTFSAccess -Path $RootImage -Account $ApplicationPackagesPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($RestrictedApplicationPackagesPrincipal)' Read and Execute on: $($RootImage)" -Severity 1
                            Add-NTFSAccess -Path $RootImage -Account $RestrictedApplicationPackagesPrincipal -AccessRights "ReadAndExecute" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($TrustedInstallerPrincipal)' Full Control on: $($RootImage)" -Severity 1
                            Add-NTFSAccess -Path $RootImage -Account $TrustedInstallerPrincipal -AccessRights "FullControl" -erroraction SilentlyContinue
                            Write-LogEntry -Value "Disabling inheritance on: $($RootImage)" -Severity 1
                            Disable-NTFSAccessInheritance -Path $RootImage -RemoveInheritedAccessRules -erroraction SilentlyContinue

                            try {
                                # Set owner to trusted installer for new wallpaper file
                                Write-LogEntry -Value "Setting ownership for '$($TrustedInstallerPrincipal)' on wallpaper image file: $($RootImage)" -Severity 1
                                Set-NTFSOwner -Path $RootImage -Account $TrustedInstallerPrincipal -erroraction SilentlyContinue
                            }
                            catch [System.Exception] {
                                Write-LogEntry -Value "Failed to set ownership for '$($TrustedInstallerPrincipal)' on wallpaper image file: $($RootImage). Error message: $($_.Exception.Message)" -Severity 3
                            }
                    }
                    catch [System.Exception] {
                            Write-LogEntry -Value "Failed to revert permissions for wallpaper image file. Error message: $($_.Exception.Message)" -Severity 3
                    }

                    try {
                        $CurrentOwner = Get-Item -Path $MSImage.fullname | Get-NTFSOwner
                        $CurrentOwnerAccountName = $CurrentOwner.Owner.AccountName
                        Write-LogEntry -Value "Owner of $MSImage.fullname is $CurrentOwnerAccountName "
    
                        if ($CurrentOwner.Owner -notlike $NewOwner) {
                            Write-LogEntry -Value "Setting new owner of '$($NewOwner)' on: $($MSImage.fullname)"
                            Set-NTFSOwner -Path $MSImage.fullname -Account $NewOwner -erroraction SilentlyContinue 
                        }

                        # Grant NT AUTHORITY\SYSTEM and Local Administrator Full Control access
                        try {
                            Write-LogEntry -Value "Granting '$($SystemContext)' Full Control on: $($MSImage.fullname)"
                            Add-NTFSAccess -Path $MSImage.fullname -Account $SystemContext -AccessRights FullControl -AccessType Allow -erroraction SilentlyContinue
                            Write-LogEntry -Value "Granting '$($NewOwner)' Full Control on: $($MSImage.fullname)"
                            Add-NTFSAccess -Path $MSImage.fullname -Account $NewOwner -AccessRights FullControl -AccessType Allow -erroraction SilentlyContinue
                        }
                        catch [System.Exception] {
                            Write-LogEntry -Value "Unable to grant required Full Control permissions on: $($MSImage.fullname). Error message: $($_.Exception.Message)" 
                        }

                        # Replace wallpaper
                        try {
                            Write-LogEntry -Value "Replacing default wallpaper in: $($MSImage.fullname)"
                            Remove-Item -Path $MSImage.fullname -Force -erroraction SilentlyContinue
                            Copy-Item -Path $RootImage.fullname -Destination $MSImage.fullname -Force -erroraction SilentlyContinue
                        }
                        catch [System.Exception] {
                            Write-LogEntry -Value "Unable to replace image: $($MSImage.fullname). Error message: $($_.Exception.Message)"  -Severity 3
                        }
                    }
                    catch [System.Exception] {
                        Write-LogEntry -Value "Unable to take ownership of: $($MSImage.fullname). Error message: $($_.Exception.Message)"  -Severity 3
                    }
              }

              else{ 
                Write-LogEntry -Value "THE '$($RootImage.fullname)' and '$($MSImage.fullname)'hashes match " 
              }

    }        
  
############################################################ Transcoded file management: This replaces / manages images from the registry and userdata theme directory

        # Find Transcoded Wallpaper registry source to use as a reference 
        if (test-path -path $RKP1 ){
        
            Try {  
                   Get-ItemProperty $RKP1  TranscodedImageCache -erroraction SilentlyContinue         
                   Write-LogEntry -Value "found the Transcoded image property value"        
                   $TIC=(Get-ItemProperty $RKP1  TranscodedImageCache -erroraction SilentlyContinue).TranscodedImageCache         
                   $TICSourcePath = [System.Text.Encoding]::Unicode.GetString($TIC) -replace '(.+)([A-Z]:[0-9a-zA-Z\\])+','$2'         
                   Write-LogEntry -Value "The Transcoded image source path is: '$($TICSourcePath)'. "
            }
                   
            catch [System.Exception]  {
                   Write-LogEntry -Value "Could not find the Transcoded image source path. Error message: $($_.Exception.Message)" -Severity 3
                   Write-LogEntry -Value "This probably means that the script is being run under system context. Weird huh?." -Severity 3}
            }

        else {
            Write-LogEntry -Value "Could not find the $RKP1short  Registry key - Not Good!!!!. Error message: $($_.Exception.Message)" -Severity 3
        }

#Remove Transcoded Wallpaper jpg cached files

        $ThemePath = "C:\Users\$user\AppData\roaming" + "\Microsoft\Windows\Themes\"
        $ThemeFileimages = (Get-ChildItem -Path $ThemePath  -Recurse | Where-Object { $_.Extension -eq '.jpg' } |Select-Object fullname).fullname

        if  (!($ThemeFileimages)){
                     Write-LogEntry -Value "There is no cached $ThemeFileimage file to be removed. Copying a new  one"

                 try {
                      Copy-Item -Path $DefaultWallpaperImagePath  -Destination $ThemeFileimage -Force -erroraction SilentlyContinue
                      Write-LogEntry -Value "Created a new cached file"}

                 catch [System.Exception]{
                      Write-LogEntry -Value "Unable to create  a new cached file . Error message: $($_.Exception.Message) "  -Severity 3 } 
        }
        else {
        
            foreach ($ThemeFileimage in $ThemeFileimages ) {

                if ((Get-FileHash $ThemeFileimage).hash  -ne (Get-FileHash $DefaultWallpaperImagePath).hash){
                     Write-LogEntry -Value "Removing the cached wallpaper file $ThemeFileimage from the Themes cache as it doesnt match the Default Wallpaper hash value."
                         
                    try {
                         Write-LogEntry -Value "Removing the cached wallpaper file $ThemeFileimage from the Themes cache" 
                         Remove-Item -Path $ThemeFileimage -Recurse  -Force -erroraction SilentlyContinue
                         Write-LogEntry -Value "Removed the cached wallpaper file $ThemeFileimage from the Themes cache"
                         Copy-Item -Path $DefaultWallpaperImagePath  -Destination $ThemeFileimage -Force -erroraction SilentlyContinue
                         Write-LogEntry -Value "Recreated the cached wallpaper file $ThemeFileimage in the Themes cache"
                    }
                    catch [System.Exception]{
                         Write-LogEntry -Value "Unable to remove the cached wallpaper file $ThemeFileimage . Error message: $($_.Exception.Message)"  -Severity 3 
                    }
                } 
                else {
                     Write-LogEntry -Value "No need to remove the cached wallpaper file $ThemeFileimage from the Themes cache. It's hash value matches the Default Wallpaper hash value." } 
            }
        }
        
# Remove and Replace TranscodedWallpaperImage file (The file with no extension) 

        $TranscodedWallpaperImage = "TranscodedWallpaper"
        $TranscodedWallpaperImagepath = $ThemePath  + $TranscodedWallpaperImage

        if (!(test-path $TranscodedWallpaperImagepath)) {
                Write-LogEntry -Value "There is no cached $TranscodedWallpaperImage file to be removed. Creating a new blank one"
   
             try {
                 #New-Item -ItemType "file" -Path $TranscodedWallpaperpath -Name $TranscodedWallpaperImage | Out-Null
                 Copy-Item -Path $DefaultWallpaperImagePath  -Destination $TranscodedWallpaperImagepath -Force -erroraction SilentlyContinue
                 Write-LogEntry -Value "Created a new  $TranscodedWallpaperImage file"
                 Start-Sleep -Seconds 7
             }
             catch [System.Exception] {
                 Write-LogEntry -Value "Unable to create a new blank $TranscodedWallpaperImage file . Error message: $($_.Exception.Message) "  -Severity 3
             }
        }

        elseif ((test-path $TranscodedWallpaperImagepath) -and (Get-FileHash $TranscodedWallpaperImagepath).hash  -ne (Get-FileHash $DefaultWallpaperImagePath).hash) {
                  Write-LogEntry -Value "The $TranscodedWallpaperImage file will be removed."
        
             try{
           
                 Remove-Item -Path $TranscodedWallpaperImagepath -Force -erroraction SilentlyContinue
                 Write-LogEntry -Value "The $TranscodedWallpaperImage file was removed."
                 Copy-Item -Path $DefaultWallpaperImagePath  -Destination $TranscodedWallpaperImagepath -Force -erroraction SilentlyContinue
                 Write-LogEntry -Value "The $TranscodedWallpaperImage file was replaced wwith the $DefaultWallpaperImagePath."
                 Start-Sleep -Seconds 7 
             }
  
             catch [System.Exception] {
                 Write-LogEntry -Value "Unable to remove the $TranscodedWallpaperImage file. Error message: $($_.Exception.Message)"  -Severity 3
             }
        }

        else { Write-LogEntry -Value "The $TranscodedWallpaperImage file did not need to be replaced wwith the $DefaultWallpaperImagePath as the file hashes are the same."
         }
       

############################################################ End transcoded file management: This replaces manages images from the registry and userdata theme directory
   
   #Define the correct Wallpaper to use

        $DefaultWallpaperImagePath = Join-Path -Path $env:WINDIR -ChildPath "WEB\Wallpaper\Windows\img0.jpg"
        $LastWallpaperUsed = Get-ItemPropertyValue -path $RKP8 -name BackgroundHistoryPath0

        if (( $DefaultWallpaperImagePath -eq  $TICSourcePath) -and ( $LastWallpaperUsed -eq  $DefaultWallpaperImagePath)){
            $DefaultWallpaperImagePath = $DefaultWallpaperImagePath
            Write-LogEntry -Value "The Default WallpaperImage is this $($DefaultWallpaperImagePath)."
        }

        else{
            $DefaultWallpaperImagePath = $LastWallpaperUsed 
            Write-LogEntry -Value "The Default WallpaperImage is this $($DefaultWallpaperImagePath)."
        }

         
############################################################ Registry - define keys

# All keys that manage the Wallpaper and Screen settings are touched to allow Varioous OS editions to behave correctly
# Kludgy and may change with conditional OS edition determination in future

$RegKeys=@(   
            @{  Key=$RKP1 
                Properties=@(
                        @{  Name="Wallpaper"; PropertyType="String"; Value=$DefaultWallpaperImagePath}

                        @{  Name="WallpaperStyle";PropertyType="String";Value=4}

                        @{  Name="TileWallpaper"; PropertyType="String";Value= 0}

                        @{   Name="JPEGImportQuality"; PropertyType="Dword";Value=100} 
                            )
            }

             @{  Key=$RKP2 
                Properties=@(
                        @{  Name="Wallpaper"; PropertyType="String"; Value=$DefaultWallpaperImagePath}

                        @{  Name="WallpaperStyle";PropertyType="String";Value=4} 
                            )
            }
    
             @{ Key=$RKP3
                Properties=@(
                        @{  Name="NoChangingWallPaper";PropertyType="String";Value=0 }
                            )
            }
    
            @{  Key=$RKP4
                Properties=@(
                        @{  Name="DesktopImageStatus";PropertyType="Dword";Value=1}

                        @{  Name="LockScreenImage";PropertyType="String";Value=$DefaultScreenImagePath}

                        @{  Name="DesktopImagePath";PropertyType="String";Value= $DefaultWallpaperImagePath }

                        @{  Name="LockScreenImageStatus";PropertyType="Dword";Value=1} 

                        @{  Name="LockScreenImageUrl";PropertyType="String";Value=$DefaultScreenImagePath }
                             )
            }

             @{  Key=$RKP6 
                Properties=@(
                        @{  Name="PortraitAssetPath";PropertyType="String";Value=$DefaultWallpaperImagePath}

                        @{  Name="LandscapeAssetPath"; PropertyType="String";Value=$DefaultWallpaperImagePath} 

                        @{  Name="HotspotImageFolderPath";PropertyType="String";Value=$DefaultWallpaperImagePath} 
                            )
            }
              @{  Key=$RKP7 
                    Properties=@(
                            @{  Name="RotatingLockScreenOverlayEnabled";PropertyType="DWORD";Value=0}

                            @{  Name="RotatingLockScreenEnabled";PropertyType="DWORD";Value=0}  
                                )
                }

)

############################################################ End Registry - define keys  

        ForEach ($item in $RegKeys.GetEnumerator()) {

            if (!(Test-Path $item.key)){ 
                Write-LogEntry -Value "Creating registry path $item.key."
                New-Item  -Path $item.key -Force | Out-Null
            }             
              
            else{
                Write-LogEntry -Value "Working on $($item.key) ."
                $RegProperties = $item.Properties
                
                ForEach ($Arg in $RegProperties) {
                    
                    #Check for the Keys existence
                    if ([bool](Get-Item -path $item.key | Select $($Arg.name))){ 
                        $ItemProperty =  Get-Item $item.key  -erroraction SilentlyContinue 
                        $ItemValue = $ItemProperty.GetValue($($Arg.name))
                        $ItemType = $ItemProperty.GetValueKind($($Arg.name))
                        $OriginalValue =  $ItemValue
                        Write-LogEntry -Value "The $($Arg.name) key value at $($item.key) is $ItemValue."

                            #Try and change the key to the required value
                        If ($OriginalValue  -ne $($Arg.value)){
                              
                            try {
                                Write-LogEntry -Value "This loop is changing the $($Arg.name) key value to $($Arg.value) ."
                                New-ItemProperty -Path $item.key -Name $($Arg.name) -Value $($Arg.value) -PropertyType $($Arg.PropertyType) -Force -ErrorAction Stop
                            }
                            catch [System.Exception] { 
                                Write-LogEntry -Value "Unable to modify the $($Arg.name) key value to $($Arg.value) . Error message: $($_.Exception.Message)."  -Severity 3
                            }
                        }
                        else{
                             Write-LogEntry -Value "No need to modify the $($Arg.name) key value to $($Arg.value)."
                        }
                    }
                   
                    else{  
                        Write-host "The $($Arg.name)  key value  we are looking for at $($item.key) doesnt exist.."                          
                        Write-LogEntry -Value "This loop will change the $($Arg.name) key value to $($Arg.value) for the first time ."
                        New-ItemProperty -Path $item.key -Name $($Arg.name) -Value $($Arg.value) -PropertyType $($Arg.PropertyType) -Force -ErrorAction SilentlyContinue 
                    }
                }
            }
        }
        


############################################################ End registry Key amendment. 

############################################################ Refresh the user's images

# Add type to do the wallpaper refresh 

Write-LogEntry -Value "Creating the Wallpaper refresh object."
Add-Type @"
  using System.Runtime.InteropServices;
  public class Wallpaper {
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    public static void Refresh(string path) {
      SystemParametersInfo(20, 0, path, 0x01|0x02); 
    }
  }
"@
Write-LogEntry -Value "About to run the Wallpaper refresh object."
# Refresh wallpaper
[Wallpaper]::Refresh(($DefaultWallpaperImagePath ))

Write-LogEntry -Value "Ran the Wallpaper refresh object. End of script"

############################################################ End refresh the user's images


Set-ExecutionPolicy $SaveExecutionPolicy -Scope Process -Force -Confirm:$false

'@

##########################################End Powershell Define script


$PSScriptnameandpath = $DestFolder + "\" + $Purpose + "_" + $version
$PSScriptPath = "$PSScriptnameandpath.ps1"
$PowerSHellscript | Out-File -FilePath $PSScriptPath -Force

##########################################Define VBS
#VB Script for now - CHange this soonish 
$VBScript = @'
Set objShell = CreateObject("Wscript.Shell")
Set args = Wscript.Arguments
Dim PsPath

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists("C:\Program Files\PowerShell\7\pwsh.exe") Then
    PsPath="""C:\Program Files\PowerShell\7\pwsh.exe"""
ElseIf fso.FileExists("C:\Program Files\PowerShell\7-preview\pwsh.exe") Then
    PsPath="""C:\Program Files\PowerShell\7-preview\pwsh.exe"""
ElseIf fso.FileExists("C:\Program Files\PowerShell\6\pwsh.exe") Then
    PsPath="""C:\Program Files\PowerShell\6\pwsh.exe"""
Else
    PsPath="""C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
End If

For Each arg In args
    Dim PSRun
    PSRun = PsPath & " -WindowStyle hidden -ExecutionPolicy bypass -NonInteractive -File " & arg
    objShell.Run(PSRun), 0
Next

'@

##########################################End Define VBS


$VBSPath = "$DestFolder\runps1.vbs"
$VBScript | Out-File -FilePath $VBSPath -force




########################################## Define SCHTASK

function Test-RunningAsSystem {
	[CmdletBinding()]
	param()
	process {
		return [bool]($(whoami -user) -match "S-1-5-18")
	}
}

$proc = Get-CimInstance Win32_Process -Filter "name = 'explorer.exe'"
$user = Invoke-CimMethod -InputObject $proc[0] -MethodName GetOwner | select-object -ExpandProperty user
$SID = (gwmi win32_userprofile | Where-Object {$_.localpath -eq "C:\Users\$User"} | select sid).sid

#Determine what the username will be
if (!(Test-RunningAsSystem)) {$Username = $env:USERNAME}else{$Username = $user}


#Assemble the Scheduled task variables
$schtaskName = $Purpose + "_" + $version
$schtaskDescription = $Purpose

$trigger = New-ScheduledTaskTrigger -AtLogOn

$trigger = New-ScheduledTaskTrigger -AtLogOn

$class = cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler
$trigger2 = $class | New-CimInstance -ClientOnly
$trigger2.Enabled = $True
$trigger2.Subscription = '<QueryList><Query Id="0" Path="System"><Select Path="System">*[System[Provider[@Name=''Microsoft-Windows-Kernel-Power''] and EventID=105]]</Select></Query></QueryList>'

$trigger3 = $class | New-CimInstance -ClientOnly
$trigger3.Enabled = $True
$trigger3.Subscription = '<QueryList><Query Id="0" Path="Security"><Select Path="Security">*[System[Provider[@Name=''Microsoft-Windows-Security-Auditing''] and EventID=4801]]</Select></Query></QueryList>'


$principal= New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -Id "$Vendor"
#$principal= New-ScheduledTaskPrincipal -LogonType Interactive -RunLevel Limited -UserId $User
$action = New-ScheduledTaskAction -Execute C:\Windows\System32\wscript.exe -Argument "`"$VBSPath`" `"$PSScriptPath`""

#Delete older tasks
$OldTasks = (Get-ScheduledTask | Where-Object {$_.taskname -like "*$Purpose*" -and $_.taskname -ne $schtaskName  } | select taskname).taskname
            
            foreach($o in $OldTasks){
                Try{
                    Get-ScheduledTask -Taskname $o -ErrorAction SilentlyContinue | Unregister-ScheduledTask -TaskName $o -Confirm:$False
                 }
                 catch [exception]{
                    #Write-LogEntry -Value "$_.exception" -Severity 3  
                 }
            }

#Register current Schtask
$null=Register-ScheduledTask -TaskName $schtaskName -TaskPath \Corrections\ -Trigger $trigger , $Trigger2 , $trigger3  -Action $action -Principal $principal -Settings $settings -Description $schtaskDescription -Force
#start current tasdk
Get-ScheduledTask -Taskname $schtaskName | Start-ScheduledTask

########################################## End Define SCHTASK


$DetectionMethod = $DestFolder + "\" + $Function  + "-" + $Purpose + "_" + $version + ".ps1_tag"




# Add type to do the wallpaper refresh 

Write-LogEntry -Value "Creating the Wallpaper refresh object."
Add-Type @"
  using System.Runtime.InteropServices;
  public class Wallpaper {
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    public static void Refresh(string path) {
      SystemParametersInfo(20, 0, path, 0x01|0x02); 
    }
  }
"@
Write-LogEntry -Value "About to run the Wallpaper refresh object."
# Refresh wallpaper
[Wallpaper]::Refresh(($DefaultWallpaperImagePath ))
Set-Content -Path $DetectionMethod -Value "Installed" 
Write-LogEntry -Value "Ran the Wallpaper refresh object. End of script"   

Set-ExecutionPolicy $SaveExecutionPolicy
