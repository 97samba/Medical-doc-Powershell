$ENTETE = @"
DOCTEUR PAPA BIRANE BOYE
N° 0499 DE L’ORDRE NATIONAL DES MEDECINS DU SENEGAL 
MEDECINE DU TRAVAIL
MEDECINE GENERALE
Rue des CHAIS BEL AIR
TEL 774503252
E- Mail : pbboye@hotmail.com
BP 3750
"@

$BODY = @"
Je soussigné Docteur d’Etat de médecine PAPA BIRANE BOYE, certifie avoir consulté dans le cadre de la visite médicale annuelle $(Get-Date -Format "yyyy"), 
"@

$APTE = @"
Par conséquent certifions qu’il n’est atteint d’aucune maladie cliniquement ni radiologiquement décelable
"@

$EXCEL_FILE = "People.xlsx"
$WORD_FILE = "Certificats.docx"
$LIST_ITEMS = @(
  "Examen clinique normal", "Examen biologique normal", "Examen radiologique pulmonaire normal"
)
$DATAS = @()


function write-document($firstName, $lastName) {
  if (Test-Path -Path $EXCEL_FILE) {
    
    write-host ==> Lecture du fichier Excel`n 
    $DATAS = import-excel -path $EXCEL_FILE -AsDate 'birthday'
  }
  else {
    Write-Host "-**-ERREUR Le fichier People.xlsx est introuvable"
    Start-sleep -s 10
    exit
  }

  $WordDocument = New-WordDocument $WORD_FILE 
  
  $DATAS | ForEach-Object {
    $oneData = $_
    
    $FONT_FAMILY = "Cambria"
    
    ## add 3 paragraphs
    Add-WordText -WordDocument $WordDocument -Text "$ENTETE `n`n" -FontSize 11 -FontFamily $FONT_FAMILY  > $null
    Add-WordText -WordDocument $WordDocument -Text "`t`tOLEOSEN" -FontSize 22 -FontFamily $FONT_FAMILY -bold $true -Color DarkRed  > $null
    Add-WordText -WordDocument $WordDocument -Text "Dakar le $(Get-Date -Format "dd/MM/yyyy")`n`n`n" -FontSize 11 -FontFamily $FONT_FAMILY  -Alignment right > $null
  
  
    Add-WordText -WordDocument $WordDocument -Text "CERTIFICAT MEDICAL`n`n" -FontSize 14 -FontFamily $FONT_FAMILY -Bold $true -Color DarkBlue -Alignment center > $null
  
    Add-WordText -WordDocument $WordDocument -Text $BODY -FontSize 11 -FontFamily $FONT_FAMILY > $null
    Add-WordText -WordDocument $WordDocument -Text "Mr. $($oneData.firstName) $($oneData.lastName)" -FontSize 11 -FontFamily $FONT_FAMILY -Bold $true > $null
    Add-WordText -WordDocument $WordDocument -Text "Né (e) le $(Get-Date $oneData.birthday -Format "dd/MM/yyyy")`n" -FontSize 11 -FontFamily $FONT_FAMILY -Bold $true > $null
  
    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $LIST_ITEMS   > $null
  
    Add-WordText -WordDocument $WordDocument -Text "`n$APTE, en conclusion l’estimons :`n" -FontSize 11 -FontFamily $FONT_FAMILY  > $null
    Add-WordText -WordDocument $WordDocument -Text "APTE pour le travail. `n" -FontSize 11 -FontFamily $FONT_FAMILY -bold $true > $null
  
    Add-WordText -WordDocument $WordDocument -Text "`n" -FontSize 11 -FontFamily $FONT_FAMILY -bold $true > $null
  
    Add-WordText -WordDocument $WordDocument -Text "Le médecin d'entreprise" -FontSize 11 -FontFamily $FONT_FAMILY -UnderlineStyle singleLine -Alignment center > $null
  
    Add-WordPicture -WordDocument $WordDocument -ImagePath "./src/Cachet.jpg" -ImageWidth 200 -ImageHeight 110 -Alignment center > $null
  
    Add-WordPageBreak -WordDocument $WordDocument  > $null

    write-host "  --OK--> Certificat pour $($oneData.firstName) $($oneData.lastName)`n"
    Start-sleep -Milliseconds 200
  
  }
  ### Save document
  write-host ==> Sauvegarde du fichier People.docx`n 

  Save-WordDocument $WordDocument -OpenDocument
}

write-host @"
`n#-------------------------------------#
   Création des certificats médicaux
#-------------------------------------#`n
"@


# write-document
write-document


write-host ==> Fin`n
