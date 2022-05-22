$REQUIREMENTS = @(
    @{
        name    = "ImportExcel" 
        version = "7.4.1" 
    },
    @{
        name    = "PSWriteWord" 
        version = "1.1.4" 
    }
)
function verify-requirements() {

    $execution_policy = $(Get-ExecutionPolicy -Scope CurrentUser)

    if ($execution_policy -ne "Unrestricted") {
        write-host "Changement des droits d'exécution en cours" -ForegroundColor Green

        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

    }

    $REQUIREMENTS | ForEach-Object {
        if (Get-InstalledModule -Name $_.name) {
            Write-Host ===> $_.name déja installé -ForegroundColor Green 
        }
        else {
            write-host "Téléchargement du Module $_.name assurez-vous d'avoir une connexion internet stable" 
            Install-Module -Name ===> $_.name -RequiredVersion $_.version -Scope CurrentUser
        }
    }  
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
  
}

Start-sleep -s 5
verify-requirements
Start-sleep -s 5
