# ManagedFavorites
Create the JSON file for Google Chrome and Microsoft Edge Managed Bookmarks/Favorites from a CSV file
Also creates a set of normal shortcuts (LNK) files, for import into Internet Explorer, legacy Edge, etc.

How to use
Need to create a CSV file in the same folder that the script is in, called "ManagedFavorites.csv".
This will contain four columns; Parent,Folder,Name,URL
Parent = root, if the top level - NB: this does not work - there for the future to allow for nested folders.
Folder = the name of the folder, if it's not in the root.
Name = displayed name of the shortcut
URL = URL of the shortcut
(In future, the location and name of the CSV file could be a parameter if the script is made into a function.)

Need to edit the script and enter the name of the folder in which the JSON file (and the IE shortcuts) will reside. This will also be the name of the managed favorites as seen in Chrome and Edge.
  $favrootname
(In future, this could be a parameter if script is made into a function.)

How to import into the registry
Sample commands;
Set-ItemProperty -Path HKCU:\Software\Policies\Microsoft\Edge -Name ManagedFavorites -Value ($jsonBase | ConvertTo-Json -Depth 100)
Set-ItemProperty -Path HKCU:\Software\Policies\Google\Chrome -Name ManagedBookmarks -Value ($jsonBase | ConvertTo-Json -Depth 100) 
Set-ItemProperty -Path HKCU:\Software\Policies\Microsoft\Edge -Name ManagedFavorites -Value (Get-Content -Path (Join-Path -Path $rootfolder -ChildPath 'ManagedFavourites.json'))

References
http://www.chromium.org/administrators/policy-list-3#ManagedBookmarks and https://cloud.google.com/docs/chrome-enterprise/policies/?policy=ManagedBookmarks
https://docs.microsoft.com/en-us/deployedge/microsoft-edge-policies#managedfavorites
