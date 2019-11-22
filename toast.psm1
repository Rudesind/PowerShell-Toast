<#
    Module : toast.psm1
    Updated: 11/22/2019
    Author : Rudesind <rudesind76@gmail.com>
    Version: 1.0

    Summary:
    This module is used to create custom Windows 10 Toast notificaitons.

#>

# Status codes
#
enum StatusCode {

    FAILED_INITIALIZATION   = 4000
    FAILED_MODULE_LOAD      = 4001
    FAILED_LOG_LOAD         = 4002
    FAILED_ASSEMBLY_LOAD    = 4003
    FAILED_ASSIGNING_APPID  = 4004
    FAILED_LOADING_XML      = 4005
    FAILED_CREATING_TOAST   = 4006
    FAILED_EXIT             = 4007
    SUCCESS                 = 0
    FATAL_ERROR             = -1

}

Function New-Toast { 
    <# 
    .Synopsis 
        Creates a Windows 10 Toast notification from a custom template. 
     
    .Description 
        Creates a Windows 10 Toast notification from a supplied custom 
        XML template.

    .Parameter XmlPath
        The XML file to be used for creating the toast notification
    
    .Parameter AppID
        The app to run the notification under. Default is "Windows PowerShell"
        
    .Parameter Debugging
        (Optional) Turn on local console debugging when running this script.
    
    .Notes

        Module : toast.psm1
        Updated: 11/22/2019
        Author : Rudesind <rudesind76@gmail.com>
        Version: 1.0

    .link
        For information on creating a toast XML document, see:
            https://docs.microsoft.com/en-us/windows/uwp/design/shell/tiles-and-notifications/adaptive-interactive-toasts
    .Link
        For information on advanced parameters, see: 
            https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters?view=powershell-6

    .Link
        More information on variable types at: 
            https://docs.microsoft.com/en-us/powershell/developer/cmdlet/strongly-encouraged-development-guidelines

    .Link
        Details on debug options: 
            https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_preference_variables?view=powershell-6
    
    .Inputs
        System.String

    .Outputs
        System.Int32

    .Example
        New-Toast .\template.xml

        This creates a toast notification using the specified XML file as a 
        template.
    
    .Example
        New-Toast .\template.xml "Command Prompt"

        This creates a toast notification using the specified XML file as a 
        template, and runs the notification under the program "Command Prompt."
#>

    # Allows the function to operate like a complied cmdlet
    #
    [CmdletBinding()]

    Param (

        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Path $_})]
        [Parameter(ValueFromPipeline=$True, Mandatory=$True)]
        [string] $XmlPath,

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline=$True)]
        [string] $AppID = [string]::Empty,

        [ValidateNotNullOrEmpty()]
        [Parameter(DontShow)]
        [System.Management.Automation.SwitchParameter] $Debugging
    )
    
    # Debugging
    #
    try {

        if ($Debugging) {

            $DebugPreference = "Continue"

            Write-Debug "Debug logging has been enabled"

        }

    } catch {
        return [int][StatusCode]::FATAL_ERROR
    }

    # Initialize Variables
    #
    try {

        Write-Debug "Initializing Variables"

        # Friendly error message for the function
        #
        [string] $errorMsg = [string]::Empty

        # Holds the raw xml from the file
        #
        [Object[]] $toastXmlRaw = @() 

        # XML object representing the template
        #
        [Windows.Data.XML.Dom.XmlDocument] $toastXml = $null

        # Object that represents the toast notificaiton
        #
        [Windows.UI.Notifications.ToastNotification] $toast = $null

    } catch {
        Write-Debug "Error, could not initialize variables: " + $Error[0]
        return [int][StatusCode]::FAILED_INITIALIZATION
    }
    
    # Load the necessary assemblies for toast notifications
    #
    try {

        Write-Debug "Loading Toast assemblies"

        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
        [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null

    } catch {
        return [int][StatusCode]::FAILED_ASSEMBLY_LOAD
    }

    # Load default PowerShell ID if variable is empty
    #
    try {

        Write-Debug "Checking for app ID"

        if ($AppID -eq [string]::Empty) {

            Write-Debug "No app ID found, assigning to 'Windows PowerShell'"

            $AppID = (Get-StartApps | Where-Object Name -eq "Windows PowerShell").AppID

        } else {

            $AppID = (Get-StartApps | Where-Object Name -eq $AppID).AppID

        }
        
    } catch {
        $errorMsg = "Error, could not validate app ID: " + $Error[0]
        Write-Debug $errorMsg
        return [int][StatusCode]::FAILED_ASSIGNING_APPID
    }

    try {

        Write-Debug "Loading XML into object"

        $toastXmlRaw = (Get-Content $XmlPath)
        $toastXml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $toastXml.LoadXml($toastXmlRaw)

    } catch {
        $errorMsg = "Error, couldn't load XML object: " + $Error[0]
        Write-Debug $errorMsg
        return [int][StatusCode]::FAILED_LOADING_XML
    }

    # Create the new toast notification
    #
    try {

        Write-Debug "Create the toast notification"

        $toast = New-Object Windows.UI.Notifications.ToastNotification $toastXML
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppID).Show($toast)

    } catch {
        $errorMsg = "Error, couldn't create toast notificaiton: " + $Error[0]
        Write-Debug $errorMsg
        return [int][StatusCode]::FAILED_LOADING_XML
    }


    Write-Debug "Toast notification created with no issues"

    return [int][StatusCode]::SUCCESS

}

Function New-SimpleToast { 
    <# 
    .Synopsis 
        Creates a Windows 10 Toast notification.
     
    .Description 
        Creates a simple Windows 10 Toast notification. This notification is 
        limitied to adding the title, body, images and some additional text.
    
    .Parameter AppID
        The app to run the notification under. Default is "Windows PowerShell"
        
    .Parameter Debugging
        (Optional) Turn on local console debugging when running this script.
    
    .Notes
        Module : toast.psm1
        Updated: 11/22/2019
        Author : Rudesind <rudesind76@gmail.com>
        Version: 1.0

    .link
        For information on creating a toast XML document, see:
            https://docs.microsoft.com/en-us/windows/uwp/design/shell/tiles-and-notifications/adaptive-interactive-toasts
    .Link
        For information on advanced parameters, see: 
            https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters?view=powershell-6

    .Link
        More information on variable types at: 
            https://docs.microsoft.com/en-us/powershell/developer/cmdlet/strongly-encouraged-development-guidelines

    .Link
        Details on debug options: 
            https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_preference_variables?view=powershell-6
    
    .Inputs
        System.String

    .Outputs
        System.Int32

    .Example
        New-SimpleToast -ToastTitle "Hello" -ToastBody "This is a toast notificaiton" -ToastAttribution "Via IT" -ToastLogo "C:\Logo.png"

        This creates a toast notification with specified title, body and logo.
#>

    # Allows the function to operate like a complied cmdlet
    #
    [CmdletBinding()]

    Param (

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline=$True, Mandatory=$True)]
        [string] $ToastTitle,

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline=$True, Mandatory=$True)]
        [string] $ToastBody,

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline=$True)]
        [string] $ToastAttribution,

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline=$True)]
        [string] $ToastLogo,

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline=$True)]
        [string] $ToastHero,

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline=$True)]
        [string] $ToastInLine,

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline=$True)]
        [string] $AppID = [string]::Empty,

        [ValidateNotNullOrEmpty()]
        [Parameter(DontShow)]
        [System.Management.Automation.SwitchParameter] $Debugging
    )
    
    # Debugging
    #
    try {

        if ($Debugging) {

            $DebugPreference = "Continue"

            Write-Debug "Debug logging has been enabled"

        }

    } catch {
        return [int][StatusCode]::FATAL_ERROR
    }

    # Initialize Variables
    #
    try {

        Write-Debug "Initializing Variables"

        # Friendly error message for the function
        #
        [string] $errorMsg = [string]::Empty

        # The xml template used for the notification
        #
        [Object[]] $toastTemplate = @("
            <toast>
                <visual>
                    <binding template=""ToastGeneric"">
                        <text>$ToastTitle</text>
                        <text>$ToastBody</text>
                        <text placement=""attribution"">$ToastAttribution</text>
                        <image placement=""appLogoOverride"" hint-crop=""circle"" src=""$ToastLogo""/>
                        <image placement=""hero"" src=""$ToastHero""/>
                        <image src=""$ToastInLine""/>
                    </binding>
                </visual>
            </toast>
        ")

        # XML object representing the template
        #
        [Windows.Data.XML.Dom.XmlDocument] $toastXml = $null

        # Object that represents the toast notificaiton
        #
        [Windows.UI.Notifications.ToastNotification] $toast = $null

        # Object that represents the toast notificaiton
        #
        [Windows.UI.Notifications.ToastNotification] $toast = $null

    } catch {
        Write-Debug "Error, could not initialize variables: " + $Error[0]
        return [int][StatusCode]::FAILED_INITIALIZATION
    }
    
    # Load the necessary assemblies for toast notifications
    #
    try {

        Write-Debug "Loading Toast assemblies"

        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
        [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null

    } catch {
        return [int][StatusCode]::FAILED_ASSEMBLY_LOAD
    }

    # Load default PowerShell ID if variable is empty
    #
    try {

        Write-Debug "Checking for app ID"

        if ($AppID -eq [string]::Empty) {

            Write-Debug "No app ID found, assigning to 'Windows PowerShell'"

            $AppID = (Get-StartApps | Where-Object Name -eq "Windows PowerShell").AppID

        } else {

            $AppID = (Get-StartApps | Where-Object Name -eq $AppID).AppID

        }
        
    } catch {
        $errorMsg = "Error, could not validate app ID: " + $Error[0]
        Write-Debug $errorMsg
        return [int][StatusCode]::FAILED_ASSIGNING_APPID
    }

    try {

        Write-Debug "Loading XML into object"

        $toastXml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $toastXml.LoadXml($toastTemplate)

    } catch {
        $errorMsg = "Error, couldn't load XML object: " + $Error[0]
        Write-Debug $errorMsg
        return [int][StatusCode]::FAILED_LOADING_XML
    }

    # Create the new toast notification
    #
    try {

        Write-Debug "Create the toast notification"

        $toast = New-Object Windows.UI.Notifications.ToastNotification $toastXML
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppID).Show($toast)

    } catch {
        $errorMsg = "Error, couldn't create toast notificaiton: " + $Error[0]
        Write-Debug $errorMsg
        return [int][StatusCode]::FAILED_LOADING_XML
    }


    Write-Debug "Toast notification created with no issues"

    return [int][StatusCode]::SUCCESS

}