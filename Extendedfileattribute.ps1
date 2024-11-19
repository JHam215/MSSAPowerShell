#Creating a shell.application COM object to reference a files extended attributes
# one way: $picture = (New-Object -ComObject Shell.Application).NameSpace("$pwd")

$path = 'C:\photoedit\DSC_3848.JPG'
$shell = New-Object -COMObject Shell.Application
$folder = Split-Path $path
$file = Split-Path $path -Leaf
$shellfolder = $shell.Namespace($folder)
$shellfile = $shellfolder.ParseName($file)

# Access extended atributes like this: $shellfolder.GetDetailsOf($picture, 216)

# The next few lines will list all extended atributes assumes that no indices higher than 1000 exist.
<#
0..1000 | % { 
  if ($n = $picture.GetDetailsOf($null, $_)) { 
    [pscustomobject] @{ Index = $_; Name = $n } 
  } 
}
#>
$AtribToModNUM=0,12,15,20,25
$AtribToModSTR="Name", "Date taken", "Authors", "Copyright"
$NumberOfAtribs = @($AtribToModNUM).Count
$count=0
Do {
    write-output $count, $AtribToModSTR[$count], $Shellfolder.GetDetailsOf($shellfile, $AtribToModNUM[$count])
    $count++
} while ($count -le $NumberOfAtribs)

<# Attributes I want to access/modify:     
0 Name, 4 Date created, 12 Date taken, 15 Year, 20 Authors, 25 Copyright, 
188 Creators, 189 Date, 208 Media created 
256 Event, 257 Exposure bias, 258 Exposure program,
259 Exposure time, 260 F-stop,
261 Flash mode, 262 Focal length,
263 35mm focal length, 264 ISO speed, 265 Lens maker,
266 Lens model, 267 Light source,
268 Max aperture, 269 Metering mode,
270 Orientation, 271 People, 272 Program mode, 273 Saturation, 
274 Subject distance, 275 White balance
#>

#Set atributes 
$newName = "2024-03-24_Hamers-Norway" #attribute 0
$newCopyright = "2024 Jeff Hamer" #attribute 25
$newAuthors = "Jeff Hamer " # Attribute 20

# Update the author name metadata
#Set-ItemProperty -Path $filepath -Name "Authors" -Value $newAuthor
$Shellfolder.SetDetailsOf($shellfile, 0)=$newName
$Shellfolder.SetDetailsOf($shellfile, 20)=$newAuthors
$Shellfolder.SetDetailsOf($shellfile, 25)=$newCopyright
$Shellfolder.

# Verify the update
#Get-ItemProperty -Path $filePath | Get-Member

