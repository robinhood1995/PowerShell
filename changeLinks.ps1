#powershell -noexit -ExecutionPolicy ByPass -File lnk_change.ps1
$oldPrefix = "019:8080"
$newPrefix = "003:8095"
$searchPath = "C:\Users\ExtusSteLin01\Desktop\Greensboro"

$shell = new-object -com wscript.shell
write-host "Updating shortcut target" -foregroundcolor red -backgroundcolor black

dir $searchPath -filter *.lnk -recurse | foreach {
$lnk = $shell.createShortcut( $_.fullname )
$oldPath= $lnk.Arguments
$lnkRegex = [regex]::escape( $oldPrefix )

#write-host $oldPath + " " + $lnkRegex

    if ( $oldPath -match $lnkRegex ) {
        $newPath = $oldPath -replace $lnkRegex, $newPrefix

        write-host "Found: " + $_.fullname -foregroundcolor yellow -backgroundcolor black
        write-host " Replace: " + $oldPath
        write-host " With: " + $newPath
        $lnk.Arguments = $newPath
        $lnk.Save()
    }
}