#create a shell application
$Shell = New-Object -ComObject Shell.application

#Get file information
$filepath = "C:\photoedit\DSC_3848.JPG"
$folder = $shell.NameSpace((Split-Path $filePath -Parent))
# $folder = $shell.namespace((get-item $filepath).directoryname)
$file = $folder.ParseName((Split-Path $filePath -Leaf))
# $file = $folder.parseName((get-item $filepath).name)

#$folder.GetDetailsOf($file, 21)   # Title
#$folder.GetDetailsOf($file, 13)   # Contributing artists
#$folder.GetDetailsOf($file, 15)   # Year

#$folder.GetDetailsOf($file, f*)  

# Define the path to the photo and the new author name
$newAuthor = "Jeff Hamer"

# Update the author name metadata
# does not work: Set-ItemProperty -Path $filepath -Name "Authors" -Value $newAuthor

# Set an extended file attribute
$file.ExtendedProperty("System.Author") = "John Doe"

# Verify the update
Get-ItemProperty -Path $filePath | Get-Member

Get-ItemProperty -Path $filePath | FL -property * -Force

# Read an extended file attribute
$author = $file.ExtendedProperty("System.Author")
Write-Output $author 

Get-ChildItem $filepath  | ForEach{
    $_.Attributes = [System.IO.FileAttributes]($_.Attributes.value__ + 8192)
    Write-Output($_.Attributes.value)
}

