function Set-Shortcut
{
    param ( [string]$Source, [string]$DestinationPath )

    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($DestinationPath)
    $Shortcut.TargetPath = $Source
    $Shortcut.Save()
}

$location = split-path -parent $MyInvocation.MyCommand.Definition
$favrootname = "Arcadis Links"

#get csv file
$FavsImport = Import-Csv -Path (Join-Path -Path $location -ChildPath 'ManagedFavorites.csv')

#setup folder location
$rootfolder = Join-Path -Path $location -ChildPath $favrootname
#remove existing folder, if it exists, to start fresh
if (Test-Path $rootfolder) {Remove-Item -Path $rootfolder -Recurse}
#not create a new folder!
if (!(Test-Path -Path $rootfolder)) {New-Item -Path $rootfolder -ItemType Directory}

#Create IE shortcuts (which Edge will also use)
#start processing the csv import
foreach ($fav in $FavsImport)
    {
    if ($fav.Folder -ne 'root')
        {
        $destfolder = Join-Path -Path $rootfolder -ChildPath $fav.Folder
        if (!(Test-Path $destfolder)) {New-Item -Path $destfolder -ItemType Directory}
        }
        else
        {
        $destfolder = $rootfolder
        }
    $destfile = Join-Path -Path $destfolder -ChildPath $fav.Name
    $destfile = $destfile + ".url"
    Set-Shortcut -Source $fav.URL -DestinationPath $destfile
    }

#Create JSON file for Edge Chromium and Chrome
#Get list of unique folder names from csv import
$folders = $FavsImport | Select-Object "Folder" -Unique
#Setup overall array for JSON
$jsonBase = New-Object System.Collections.ArrayList
#Set title of managed favourites
$jsonBase.Add(@{"toplevel_name"=$favrootname;})
#process through each folder name
foreach ($foldername in $folders)
    {
    $list = New-Object System.Collections.ArrayList
    foreach ($fav in $FavsImport)
        {
        if ($fav.Folder -eq $foldername.Folder)
            {
            if ($fav.Folder -eq 'root')
                {
                $jsonBase.Add(@{"name"=$($fav.Name);"url"=$($fav.URL);})
                }
            else
                {
                #childfolder process
                $list.Add(@{"name"=$($fav.Name);"url"=$($fav.URL);})
                }
            }
        }
    if ($foldername.Folder -eq 'root')
        {
        #$jsonBase.Add($list)
        }
    else
        {
        $folder=@{"name"=$foldername.Folder;"children"=$list;}
        $jsonBase.Add($folder)
        }
    }
$favourites = $jsonBase | ConvertTo-Json -Depth 100

Set-Content -Path (Join-Path -Path $rootfolder -ChildPath 'ManagedFavourites.json') -Value $favourites

#Set-ItemProperty -Path HKCU:\Software\Policies\Microsoft\Edge -Name ManagedFavorites -Value ($jsonBase | ConvertTo-Json -Depth 100)
#Set-ItemProperty -Path HKCU:\Software\Policies\Google\Chrome -Name ManagedBookmarks -Value ($jsonBase | ConvertTo-Json -Depth 100) 
#Set-ItemProperty -Path HKCU:\Software\Policies\Microsoft\Edge -Name ManagedFavorites -Value (Get-Content -Path (Join-Path -Path $rootfolder -ChildPath 'ManagedFavourites.json'))