# Module ImportExcel
Install-Module -Name ImportExcel -Scope CurrentUser

# ----- Nom de l'utilisateur -----
$UserInfo = Get-CimInstance Win32_Process -Filter 'name="explorer.exe"' | Invoke-CimMethod -MethodName GetOwner
$UserName = $UserInfo.User

# ------ BIOS -----
$Bios = Get-CimInstance -ClassName Win32_BIOS
$BiosCaption = $Bios.Caption
$BiosVersion = $Bios.Version
$BiosSMBIOS = $Bios.SMBIOSMajorVersion
$BiosSystemBiosVersion = $Bios.SystemBiosMajorVersion
$BiosReleaseDate = $Bios.ReleaseDate
$BiosSerialNumber = $Bios.SerialNumber

# ----- Processeur -----
$Processor = Get-CimInstance -ClassName Win32_Processor
$ProcessorName = $Processor.Name
$ProcessorCaption = $Processor.Caption
$ProcessorMaxClockSpeed = $Processor.MaxClockSpeed
$ProcessorNumberOfCores = $Processor.NumberOfCores
$ProcessorSocketDesignation = $Processor.SocketDesignation

# ----- Carte mère -----
$MotherBoard = Get-CimInstance -ClassName Win32_BaseBoard
$MotherBoardManufacturer = $MotherBoard.Manufacturer
$MotherBoardProduct = $MotherBoard.Product
$MotherBoardVersion = $MotherBoard.Version

# ----- RAM -----
function Convert-SMBIOSMemoryType {
    param (
        [int]$SMBIOSMemoryType
    )
    switch ($SMBIOSMemoryType) {
        20 { return "DDR" }
        21 { return "DDR2" }
        22 { return "DDR2 FB-DIMM" }
        24 { return "DDR3" }
        26 { return "DDR4" }
        34 { return "DDR5" }
        default { return [int]$SMBIOSMemoryType }
    }
}

$MemoryInstance = Get-CimInstance -ClassName CIM_PhysicalMemory
$MemorySpeed = $MemoryInstance.Speed
$MemoryDeviceLocator = $MemoryInstance.DeviceLocator
$MemoryCapacity = ($MemoryInstance.Capacity / 1GB)
$MemoryType = $MemoryInstance | Select-Object @{
    label = 'MemoryType';
    expression = { Convert-SMBIOSMemoryType -SMBIOSMemoryType $_.SMBIOSMemoryType }
}

# ----- Carte Graphique -----
$GraphicsInstance = Get-CimInstance -ClassName CIM_VideoController
$GraphicsName = $GraphicsInstance.Name
$GraphicsDriverVersion = $GraphicsInstance.DriverVersion
$GraphicsVRAM = ($GraphicsInstance.AdapterRAM / 1GB).ToString('F2')

# ----- Disque dur -----
$DiskInstance = Get-CimInstance -Class Win32_LogicalDisk -Filter "DriveType=3"
$DiskDeviceID = $DiskInstance.DeviceID
$DiskVolumeName = $DiskInstance.VolumeName
$DiskFreeSpace = ($DiskInstance.FreeSpace / 1GB).ToString('F2')
$DiskUsedSpace = ((($DiskInstance.Size - $DiskInstance.FreeSpace) / 1GB).ToString('F2'))
$DiskTotalSpace = ($DiskInstance.Size / 1GB).ToString('F2')

# ----- OS / système -----
$OSInstance = Get-CimInstance -ClassName CIM_OperatingSystem
$OSCaption = $OSInstance.Caption
$OSVersion = $OSInstance.Version
$OSInstallDate = $OSInstance.InstallDate
$OSArchitecture = $OSInstance.OSArchitecture
$OSWindowsDirectory = $OSInstance.WindowsDirectory
$OSNumberOfUsers = $OSInstance.NumberOfUsers
$OSBootDevice = $OSInstance.BootDevice

$SystemInstance = Get-CimInstance -ClassName Win32_ComputerSystem
$SystemName = $SystemInstance.Name
$SystemDomain = $SystemInstance.Domain
$SystemModel = $SystemInstance.Model
$SystemManufacturer = $SystemInstance.Manufacturer

# ----- Etat de santé des disques -----
$disksHealth = Get-PhysicalDisk | Select-Object MediaType, OperationalStatus, HealthStatus

# ----- Réseaux ------
$networkConfig = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled = true"
$networkConnection = Get-CimInstance Win32_NetworkConnection

# ----- Mise à jour OS -----
$Maj = Get-CimInstance -ClassName Win32_QuickFixEngineering

# ----- Ecrans d'ordinateur -----

# Fonction de décodage des infos moniteur
function Decode {
    If ($args[0] -is [System.Array]) {
        [System.Text.Encoding]::ASCII.GetString($args[0])
    }
    Else {
        "Not Found"
    }
}

# Récupération des informations du contrôleur vidéo /!\
$videoController = Get-WmiObject Win32_VideoController

#Info moniteur
ForEach ($Monitor in Get-WmiObject WmiMonitorID -Namespace root\wmi) {
    $Manufacturer = Decode $Monitor.ManufacturerName
    $Name = Decode $Monitor.UserFriendlyName
    $Serial = Decode $Monitor.SerialNumberID
    $ManufactureYear = $Monitor.YearOfManufacture
    $HorizontalResolution = $videoController.CurrentHorizontalResolution
    $VerticalResolution = $videoController.CurrentVerticalResolution
    $RefreshRate = $videoController.CurrentRefreshRate
    $DriverVersion = $videoController.DriverVersion
    $DriverDate = $videoController.DriverDate
}


# ----- Logiciels ------

#pointe vers l'emplacement des logiciels installés dans windows (64 et 32 bits)
$keys = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
          "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")

#Pour chaque clé de registre récupère les propriétés uniquement si displayname n'est pas nulle.
$installedSoftware = $keys | ForEach-Object { Get-ItemProperty $_ } |
    Where-Object { $_.DisplayName -ne $null } |
    Select-Object DisplayName, DisplayVersion



# --------------------------------------------------- Tableau Excel -----------------------------------------------

# ----- Données système -----
$CombinedData = @(
    [PSCustomObject]@{ Property = "Utilisateur"; Value = $UserName }
    [PSCustomObject]@{ Property = "BIOS Caption"; Value = $BiosCaption }
    [PSCustomObject]@{ Property = "BIOS Version"; Value = $BiosVersion }
    [PSCustomObject]@{ Property = "BIOS SMBIOS Version"; Value = $BiosSMBIOS }
    [PSCustomObject]@{ Property = "BIOS System BIOS Version"; Value = $BiosSystemBiosVersion }
    [PSCustomObject]@{ Property = "BIOS Release Date"; Value = $BiosReleaseDate }
    [PSCustomObject]@{ Property = "Numéro de série"; Value = $BiosSerialNumber }
    [PSCustomObject]@{ Property = "Processeur Name"; Value = $ProcessorName }
    [PSCustomObject]@{ Property = "Processeur Caption"; Value = $ProcessorCaption }
    [PSCustomObject]@{ Property = "Processeur Max Clock Speed"; Value = $ProcessorMaxClockSpeed }
    [PSCustomObject]@{ Property = "Processeur Number of Cores"; Value = $ProcessorNumberOfCores }
    [PSCustomObject]@{ Property = "Processeur Socket"; Value = $ProcessorSocketDesignation }
    [PSCustomObject]@{ Property = "Carte Mère Manufacturer"; Value = $MotherBoardManufacturer }
    [PSCustomObject]@{ Property = "Carte Mère Product"; Value = $MotherBoardProduct }
    [PSCustomObject]@{ Property = "Carte Mère Version"; Value = $MotherBoardVersion }
    [PSCustomObject]@{ Property = "RAM Speed (MHz)"; Value = $MemorySpeed }
    [PSCustomObject]@{ Property = "RAM device locator"; Value = $MemoryDeviceLocator }
    [PSCustomObject]@{ Property = "RAM Capacity (GB)"; Value = $MemoryCapacity }
    [PSCustomObject]@{ Property = "RAM Type"; Value = $MemoryType.MemoryType }
    [PSCustomObject]@{ Property = "Carte Graphique"; Value = $GraphicsName }
    [PSCustomObject]@{ Property = "Version Pilote Graphique"; Value = $GraphicsDriverVersion }
    [PSCustomObject]@{ Property = "VRAM (GB)"; Value = $GraphicsVRAM }
    [PSCustomObject]@{ Property = "Disque ID"; Value = $DiskDeviceID }
    [PSCustomObject]@{ Property = "Nom du Volume"; Value = $DiskVolumeName }
    [PSCustomObject]@{ Property = "Espace Total (GB)"; Value = $DiskTotalSpace }
    [PSCustomObject]@{ Property = "Espace Libre (GB)"; Value = $DiskFreeSpace }
    [PSCustomObject]@{ Property = "Espace Utilisé (GB)"; Value = $DiskUsedSpace }
    [PSCustomObject]@{ Property = "Nom de l'OS"; Value = $OSCaption }
    [PSCustomObject]@{ Property = "Version de l'OS"; Value = $OSVersion }
    [PSCustomObject]@{ Property = "Date d'Installation"; Value = $OSInstallDate }
    [PSCustomObject]@{ Property = "Architecture de l'OS"; Value = $OSArchitecture }
    [PSCustomObject]@{ Property = "Dossier Windows"; Value = $OSWindowsDirectory }
    [PSCustomObject]@{ Property = "Nombre d'Utilisateurs"; Value = $OSNumberOfUsers }
    [PSCustomObject]@{ Property = "Périphérique de Démarrage"; Value = $OSBootDevice }
    [PSCustomObject]@{ Property = "Nom d'instance"; Value = $SystemName }
    [PSCustomObject]@{ Property = "Nom de domaine"; Value = $SystemDomain }
    [PSCustomObject]@{ Property = "Modèle Système"; Value = $SystemModel }
    [PSCustomObject]@{ Property = "Fabricant Système"; Value = $SystemManufacturer }
)

# ----- Données Réseaux -----

# ----- Données écrans -----
$CombinedDataMonitor = @(
    [PSCustomObject]@{ Property = "Fabricant"; Value = $Manufacturer }
    [PSCustomObject]@{ Property = "Nom"; Value = $Name }
    [PSCustomObject]@{ Property = "Numéro de série"; Value = $Serial }
    [PSCustomObject]@{ Property = "Année de fabrication"; Value = $ManufactureYear }
    [PSCustomObject]@{ Property = "Résolution horizontal"; Value = $HorizontalResolution }
    [PSCustomObject]@{ Property = "Résolution vertical"; Value = $VerticalResolution }
    [PSCustomObject]@{ Property = "Taux de rafraichissement"; Value = $RefreshRate }
    [PSCustomObject]@{ Property = "Version du driver"; Value = $DriverVersion }
    [PSCustomObject]@{ Property = "Date dernière mise à jour driver"; Value = $DriverDate }
)

# ----- Chemin du fichier Excel -----
$excelFilePath = "Y:\FRESNAIS\Fiche utilisateurs\Fiche_$UserName.xlsx"

# ----- Exportation des données -----
$CombinedData | Export-Excel -Path $excelFilePath -WorksheetName "Données Système" -AutoSize
$disksHealth | Export-Excel -Path $excelFilePath -WorksheetName "Etat de santé des disques" -AutoSize -Append
$networkConfig | Export-Excel -Path $excelFilePath -WorksheetName "Configuration accès réseaux" -AutoSize -Append
$networkConnection | Export-Excel -Path $excelFilePath -WorksheetName "Accès serveur" -AutoSize -Append
$Maj | Export-Excel -Path $excelFilePath -WorksheetName "Mise à jour système" -AutoSize -Append
$CombinedDataMonitor | Export-Excel -Path $excelFilePath -WorksheetName "Ecrans" -AutoSize -Append
$installedSoftware | Export-Excel -Path $excelFilePath -WorksheetName "Données logiciel" -AutoSize -Append


# ----------- Modification de l'alignement des cellules avec COM objets -----------

# Ouvrir Excel via COM objet
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False  # False = invisible / true = Visible

# Ouvrir le fichier Excel
$Workbook = $Excel.Workbooks.Open($excelFilePath)

# ---- Appliquer l'alignement à gauche sur la feuille "Données Système" ----
# Sélectionner la feuille "Données Système"
$Worksheet1 = $Workbook.Sheets.Item("Données Système")

# Sélectionner toute la feuille
$UsedRange1 = $Worksheet1.UsedRange

# Appliquer un alignement à gauche
$UsedRange1.HorizontalAlignment = -4131 #(Alignement à gauche)

# ---- Appliquer l'alignement à gauche sur la feuille "Écrans" ----
# Sélectionner la feuille "Écrans"
$Worksheet2 = $Workbook.Sheets.Item("Ecrans")

# Sélectionner toute la feuille
$UsedRange2 = $Worksheet2.UsedRange

# Appliquer un alignement à gauche
$UsedRange2.HorizontalAlignment = -4131 #(Alignement à gauche)

# Sauvegarder et fermer le fichier Excel
$Workbook.Save()
$Workbook.Close()

# Fermer l'application Excel
$Excel.Quit()

# Libérer la mémoire des objets COM
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($UsedRange1) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($UsedRange2) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet1) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet2) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
