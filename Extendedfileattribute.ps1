$folder = (New-Object -ComObject Shell.Application).NameSpace("$pwd")
# Note: Assumes that no indices higher than 1000 exist.
0..1000 | % { 
  if ($n = $folder.GetDetailsOf($null, $_)) { 
    [pscustomobject] @{ Index = $_; Name = $n } 
  } 
}