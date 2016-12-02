 <#
.SYNOPSIS
  Queries the SCCM database for a PC's last logged in user and writes back to AD
.DESCRIPTION
  This script will query the SCCM database for the username of the person who last logged into
  each machine. This is based on the heartbeat, which is accurate to ~4 minutes unless your heartbeat
  value is different.
  After the user data is pulled from SCCM, the script modifies the associated computer record.
  Both the ManagedBy and Description attributes are updated. 
  Managedby is set to the user's email address
  Description is set to the last logged in user's username and the modification date.
.EXAMPLE
  .\Get-SCCMLastLogin -SQLServer server.domain.com -SQLDBName CM1
.PARAMETER SQLServer
  The FQDN of the SCCM SQL server
.PARAMETER SQLDBName
  The SCCM Database Name
.NOTES
  Version:        1.0
  Author:         dd4495 
  Creation Date:  September 2016
#>
param (
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,Position=1)]
    [string]$SQLServer,
    [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName,Position=1)]
    [string]$SQLDBName
    )

# Query the SCCM database for the last logon time of users
# b.User_Name0 != 'Administrator'will skip the logins for Administrator
# c.name0 not like '[ABCD]1%' skips server machines
$SqlQuery =@"
select CURRENT_TIMESTAMP, b.user_name0 as [UserName], b.Mail0 as [Email], c.Name0 as [MachineName], DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), c.Last_Logon_Timestamp0) as [Date]
  from v_R_System c
	left outer join v_R_User b on (b.user_name0 = c.user_name0)
		where b.User_Name0 is not null
		and c.Last_Logon_Timestamp0 >= CAST(CURRENT_TIMESTAMP AS DATE)
		and b.User_Name0 != 'Administrator'
		and c.name0 not like '[ABCD]1%'
		and b.Mail0 >='0'
			order by c.Last_Logon_Timestamp0 desc;
"@

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection

$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd

$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)

$SqlConnection.Close()

$data = $DataSet.Tables[0]

$date = Get-Date -Format ('d')

for ($i=0; $i -lt $data.Rows.Count; $i++) {
    $pc = $data.Rows[$i].MachineName
    $un = $data.Rows[$i].UserName
    $ea = $data.Rows[$i].Email
    $desc = "Owner: $un, Date Modified: $date"

    Write-Output "Set-ADComputer -Instance  $pc -ManagedBy $ea -Description $desc"
}