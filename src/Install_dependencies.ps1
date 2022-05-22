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
        # write-host "Changement des droits d'exécution en cours" -ForegroundColor Green

        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

    }

    $REQUIREMENTS | ForEach-Object {
        if (Get-InstalledModule -Name $_.name -ErrorAction SilentlyContinue) {
            Write-Host ===> $_.name déja installé -ForegroundColor Green 
        }
        else {
            Write-Host ===> $_.name manquant -ForegroundColor Red 

            write-host "Téléchargement du Module $($_.name) assurez-vous d'avoir une connexion internet stable" 
            Install-Module -Name $($_.name) -RequiredVersion $($_.version) -Scope CurrentUser -Confirm:$False -Force
        }
    }  
    try {

        Set-ExecutionPolicy -ExecutionPolicy $execution_policy -Scope CurrentUser | Out-Null
    }
    catch {
        write-host "Fin d'installation"
    }
  
}

verify-requirements
write-host Vous pouvez lancer la création de scripts 

Start-sleep -s 5
