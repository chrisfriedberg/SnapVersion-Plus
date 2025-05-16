<#
.SYNOPSIS
Creates a shortcut (.lnk file) on the user's Desktop to run a specified Python script (.py or .pyw)
with a custom icon (.ico).

.DESCRIPTION
This script uses graphical file dialogs to prompt the user for:
1. The target Python script file (.py or .pyw).
2. The icon file (.ico) to use for the shortcut.
It then creates a .lnk shortcut file on the user's Desktop.
The shortcut is configured to run the target script using the Python executable found in the system's PATH.
It sets the working directory to the directory containing the target script.

.NOTES
Author: Gemini
Requires: Windows PowerShell or PowerShell 7+, .NET Framework (for dialogs)
Execution Policy: May require adjusting PowerShell execution policy to run.
                 (e.g., `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` or running from an admin prompt)

.EXAMPLE
.\CreatePythonShortcut.ps1
Follow the graphical prompts to select the script and icon files.
#>

# --- Configuration ---
$ErrorActionPreference = "Stop" # Exit script on error

# --- Function to Show File Dialog ---
function Show-FileDialog {
    param(
        [string]$Title,
        [string]$Filter, # Example: "Python Scripts (*.py,*.pyw)|*.py;*.pyw|All files (*.*)|*.*"
        [string]$InitialDirectory = ([Environment]::GetFolderPath("MyDocuments")) # Default start location
    )

    try {
        # Load necessary assembly for OpenFileDialog
        Add-Type -AssemblyName System.Windows.Forms

        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = $Title
        $openFileDialog.Filter = $Filter
        $openFileDialog.InitialDirectory = $InitialDirectory
        $openFileDialog.Multiselect = $false # Only allow selecting one file

        # Show the dialog and check if the user clicked OK
        $result = $openFileDialog.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Write-Verbose "File selected: $($openFileDialog.FileName)"
            return $openFileDialog.FileName
        } else {
            Write-Warning "File selection cancelled by user."
            return $null # Return null if cancelled
        }
    } catch {
        Write-Error "Damn, couldn't show the file dialog. Error: $($_.Exception.Message)"
        # Attempt to show a basic message box as fallback
        try { Add-Type -AssemblyName Microsoft.VisualBasic; [Microsoft.VisualBasic.Interaction]::MsgBox("Error showing file dialog. Check PowerShell console for details.", "OKOnly,SystemModal,Critical", "Dialog Error") } catch {}
        return $null
    } finally {
        # Clean up the dialog object if it exists
        if ($openFileDialog -ne $null) {
            $openFileDialog.Dispose()
        }
    }
}

# --- Function to Show Message Box ---
function Show-MessageBox {
    param(
        [string]$Message,
        [string]$Title = "Shortcut Creator",
        [string]$Type = "Information" # Information, Warning, Error
    )
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $buttonType = [System.Windows.Forms.MessageBoxButtons]::OK
        $iconType = switch ($Type.ToLower()) {
            "warning"     { [System.Windows.Forms.MessageBoxIcon]::Warning }
            "error"       { [System.Windows.Forms.MessageBoxIcon]::Error }
            default       { [System.Windows.Forms.MessageBoxIcon]::Information }
        }
        [System.Windows.Forms.MessageBox]::Show($Message, $Title, $buttonType, $iconType) | Out-Null
    } catch {
        Write-Warning "Couldn't show the fancy message box. Fallback to console."
        Write-Host "[$Title] $Message"
        # Attempt basic VB message box as another fallback
        try { Add-Type -AssemblyName Microsoft.VisualBasic; [Microsoft.VisualBasic.Interaction]::MsgBox($Message, "OKOnly,SystemModal", $Title) } catch {}
    }
}


# --- Main Script Logic ---
Write-Host "Starting Python shortcut creator..."

# 1. Get the target Python script
Write-Host "Prompting for Python script..."
$scriptFilter = "Python Scripts (*.py, *.pyw)|*.py;*.pyw|All files (*.*)|*.*"
$targetScriptPath = Show-FileDialog -Title "Alright, pick the Python script (.py or .pyw) you wanna run" -Filter $scriptFilter

if (-not $targetScriptPath) {
    Show-MessageBox -Message "No script selected. Quitting." -Title "Cancelled" -Type Warning
    Write-Warning "User cancelled script selection. Exiting."
    exit
}
Write-Host "Selected script: $targetScriptPath"

# 2. Get the icon file
Write-Host "Prompting for Icon file..."
$iconFilter = "Icon Files (*.ico)|*.ico|All files (*.*)|*.*"
$iconPath = Show-FileDialog -Title "Now, pick the pretty little icon (.ico) for the shortcut" -Filter $iconFilter -InitialDirectory (Split-Path $targetScriptPath -Parent) # Start near the script

if (-not $iconPath) {
    Show-MessageBox -Message "No icon selected. Quitting." -Title "Cancelled" -Type Warning
    Write-Warning "User cancelled icon selection. Exiting."
    exit
}
Write-Host "Selected icon: $iconPath"

# 3. Determine paths and names
try {
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $scriptFileName = Split-Path $targetScriptPath -Leaf # Just the filename.ext
    $shortcutNameBase = [System.IO.Path]::GetFileNameWithoutExtension($scriptFileName)
    $shortcutFileName = "$($shortcutNameBase).lnk"
    $shortcutFullPath = Join-Path -Path $desktopPath -ChildPath $shortcutFileName

    # Determine which Python executable to use (.pyw usually runs without console)
    $pythonExecutable = if ($targetScriptPath -like '*.pyw') { "pythonw.exe" } else { "python.exe" }
    # Check if the chosen executable is actually in the PATH
    $pythonFullPath = (Get-Command $pythonExecutable -ErrorAction SilentlyContinue).Source
    if (-not $pythonFullPath) {
        Show-MessageBox -Message "Shit, couldn't find '$pythonExecutable' in your system PATH. Make sure Python is installed correctly and added to PATH. Cannot create shortcut." -Title "Python Not Found" -Type Error
        Write-Error "Could not find '$pythonExecutable'. Ensure Python is installed and in PATH."
        exit
    }

    $workingDirectory = Split-Path $targetScriptPath -Parent # The folder containing the script

    Write-Host "Desktop path: $desktopPath"
    Write-Host "Shortcut name: $shortcutFileName"
    Write-Host "Target executable: $pythonFullPath"
    Write-Host "Target arguments: `"$targetScriptPath`"" # Arguments need quoting
    Write-Host "Working directory: $workingDirectory"
    Write-Host "Icon location: $iconPath"

} catch {
    Show-MessageBox -Message "Hell, couldn't figure out the paths. Error: $($_.Exception.Message)" -Title "Path Error" -Type Error
    Write-Error "Error determining paths: $($_.Exception.Message)"
    exit
}

# 4. Create the shortcut using WScript.Shell COM object
try {
    Write-Host "Creating shortcut object..."
    # Create the WScript.Shell COM object (the classic way to make shortcuts)
    $WshShell = New-Object -ComObject WScript.Shell

    # Create the shortcut object linked to the file path
    $Shortcut = $WshShell.CreateShortcut($shortcutFullPath)

    # --- Set Shortcut Properties ---
    # Target: The full path to python.exe or pythonw.exe
    $Shortcut.TargetPath = $pythonFullPath

    # Arguments: The full path to *your* script, enclosed in quotes
    $Shortcut.Arguments = """$targetScriptPath""" # Double quotes inside single quotes for literal quotes

    # WorkingDirectory: Where the script should run *from* (important for relative paths in your script)
    $Shortcut.WorkingDirectory = $workingDirectory

    # IconLocation: Path to the .ico file, index 0
    $Shortcut.IconLocation = "$iconPath,0"

    # Description (optional, shows in properties/tooltip)
    $Shortcut.Description = "Run $($scriptFileName)"

    # --- Save the shortcut ---
    Write-Host "Saving shortcut to $shortcutFullPath ..."
    $Shortcut.Save()

    # Release the COM object (good practice)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WshShell) | Out-Null
    Remove-Variable WshShell # Clean up variable

    Write-Host "Shortcut created successfully."
    Show-MessageBox -Message "Boom! Shortcut '$shortcutFileName' created on your Desktop." -Title "Success!" -Type Information

} catch {
    $errMsg = $_.Exception.Message
    Show-MessageBox -Message "God damn it, failed to create the shortcut. Error was: $errMsg" -Title "Shortcut Creation Failed" -Type Error
    Write-Error "Failed to create shortcut: $errMsg"
    # Clean up COM object if it exists and an error occurred
    if (Test-Path variable:WshShell) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WshShell) | Out-Null } catch {}
        Remove-Variable WshShell
    }
    exit
}

Write-Host "Script finished."
