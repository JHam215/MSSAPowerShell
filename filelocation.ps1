#create a shell application
$Shell = New-Object -ComObject Shell.application

#Get file information
$filepath = "C:\photoedit\DSC_3848.JPG"
$folder = $shell.namespace((get-item $filepath).directoryname)
$file = $folder.parseName((get-item $filepath).name)

#$folder.GetDetailsOf($file, 21)   # Title
#$folder.GetDetailsOf($file, 13)   # Contributing artists
#$folder.GetDetailsOf($file, 15)   # Year

#$folder.GetDetailsOf($file, f*)  

# Define the path to the photo and the new author name
$newAuthor = "Jeff Hamer"

# Update the author name metadata
Set-ItemProperty -Path $filepath -Name "Authors" -Value $newAuthor

# Verify the update
Get-ItemProperty -Path $filePath | Get-Member

Get-ItemProperty -Path $filePath | FL -property * -Force


