$SQLServer = "hskiw-espdb-pg1\kiwprod"
$SQLDatabase = "Stora_Live"
$plant = "Tallinn"
$SpecFolder = "Specs"
$extension = ".pdf"
$folder = "\\HSKIW-ESPFS-P01.group.corp.storaenso.com\drawings\Baltic\$plant\$SpecFolder"
$archivefolder = "\\HSKIW-ESPFS-P01.group.corp.storaenso.com\drawings\Baltic\$plant\$SpecFolder\Archive"

# Look into ESP database
$sqlConn = New-Object System.Data.SqlClient.SqlConnection
$sqlConn.ConnectionString = “Server=$SQLServer;Integrated Security=true;Initial Catalog=$SQLDatabase”
$sqlConn.Open()
#$sqlcmd = $sqlConn.CreateCommand()
$sqlcmd = New-Object System.Data.SqlClient.SqlCommand
$sqlcmd.Connection = $sqlConn

#Query
$query = “select rtrim(ltrim(designnumber)) NewDesignNumber, isnull(userbd12,'') OldDesignNumber 
                from ebxproductDesign pd 
                inner join ebxproductdesignplant pdp
                on pdp.productdesignid=pd.id
                inner join orgplant p 
                on p.id=pdp.plantid 
                where p.name='$plant'
                and isnull(userbd12,'')!=''”
$sqlcmd.CommandText = $query
#This is if you wish to display the result
$adp = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcmd
$data = New-Object System.Data.DataSet
$adp.Fill($data) | Out-Null

foreach ($Row in $data.Tables[0].Rows)
{

$oldname=$row.OldDesignNumber.Replace(' ','')
$newname=$row.NewDesignNumber.Replace(' ','')

Write-Host "Renaming file $oldname To" $row.NewDesignNumber
#Write-Host "Copying file from $folder\$oldname$Extension to $archivefolder"
#Copy-Item $folder\$oldname$extension -Destination $archivefolder
Write-Host "Renaming file from $folder\$oldname$Extension to $folder\$newname$Extension"
Rename-Item $folder\$oldname$Extension -NewName $folder\$newname$Extension

}