#Credit Doug Maurer source https://stackoverflow.com/questions/65414088/efficient-script-to-get-all-extended-file-properties
Function Get-FileMetaData {
    [cmdletbinding()]
    Param
    (
        [parameter(valuefrompipeline,ValueFromPipelineByPropertyName,Position=1,Mandatory)]
        $InputObject
    )

    begin
    {
        $shell = New-Object -ComObject Shell.Application
    }

    process
    {
        foreach($object in $InputObject)
        {
            if($object -is [string])
            {
                try
                {
                    $object = Get-Item $object -ErrorAction Stop
                }
                catch
                {
                    Write-Warning "Error while processing $object : $($_.exception.message)"
                    break
                }
            }

            try
            {
                Test-Path $object -ErrorAction Stop
            }
            catch
            {
                Write-Warning "Error while processing $($object.fullname) : $($_.exception.message)"
                break
            }

            switch ($object)
            {
                {$_ -is [System.IO.DirectoryInfo]}{
                    write-host Processing folder $object.FullName -ForegroundColor Cyan
                    $currentfolder = $shell.namespace($object.FullName)
                    $items = $currentfolder.items()
                }
                {$_ -is [System.IO.FileInfo]}{
                    write-host Processing file $object.FullName -ForegroundColor Cyan
                    $parent = Split-Path $object
                    $currentfolder = $shell.namespace($parent)
                    $items = $currentfolder.ParseName((Split-Path $object -Leaf))
                }
            }

            try
            {
                foreach($item in $items)
                {
                    0..512 | ForEach-Object -Begin {$ht = [ordered]@{}}{
                        if($value = $currentfolder.GetDetailsOf($item,$_))
                        {
                            if($propname = $currentfolder.GetDetailsOf($null,$_))
                            {
                                $ht.Add($propname,$value)
                            }
                        
                        }    
                    } -End {[PSCustomObject]$ht}
                }
            }
            catch
            {
                Write-Warning "Error while processing $($item.fullname) : $($_.exception.message)"
            }
        }
    }

    end
    {
        $shell = $null
    }
}

 Get-FileMetaData "C:\photoedit\DSC_3848.JPG"
  Get-FileMetaData "C:\photoedit"