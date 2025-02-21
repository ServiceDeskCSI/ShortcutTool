<# 
Use this func to create a new shortcut on a public desktop 

New-DesktopShortcut -Name "Google" -Icon ".\googleicon.ico" -Target "https://google.com" -Public
#>
function New-DesktopShortcut {
    [CmdletBinding(DefaultParameterSetName = 'Public')]
    param (
        [parameter(Mandatory = $false)]
        [string]
        $Icon,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [parameter(Mandatory = $true)]
        [string]
        $Target,

        [parameter(Mandatory = $false, ParameterSetName = 'Public')]
        [switch]
        $Public,

        [parameter(Mandatory = $false, ParameterSetName = 'User')]
        [switch]
        $User
    )

    switch ($PSCmdlet.ParameterSetName) {
        'Public' {
            $Path = [System.Environment]::GetFolderPath('CommonDesktop')
        }
        'User' {
            $Path = [System.Environment]::GetFolderPath('Desktop')
        }
    }

    $ShortcutPath = Join-Path -Path $Path -ChildPath "$Name.lnk"
    
    # Create the shortcut
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)

    # Set shortcut properties
    $Shortcut.TargetPath = $Target
    if ($Icon) {
        $Shortcut.IconLocation = $Icon
    }

    # Save the shortcut
    $Shortcut.Save()
}