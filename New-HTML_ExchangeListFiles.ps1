<#
.SYNOPSIS
  Gets all the groups in Exchange and writes member info to an HTML formatted file
.DESCRIPTION
  This script will get all the Distribution groups and Dynamic distributions groups in Exchange
  and output the member lists to HTML files. It also creates an index file that contains links 
  to each individual list page.
.NOTES
  Version:        1.0
  Author:         dd4495 
  Creation Date:  December 2016
#>

# Get all the distribution groups in Exchange
$distGroups = (Get-DistributionGroup -ResultSize unlimited | Select-Object alias, name)
$dynDistGroups = (Get-DynamicDistributionGroup | Select-Object alias,name)
$outFileLoc = 'F:\temp\webtest'
$indexFile = "$outFileLoc\index.html"
$liList = @()

# Get the members of each distribution group
foreach ($group in $distGroups) {
    $alias = $group.alias
    $name = $group.name
    $shortFile = "$alias.html"
    $outFile = "$outFileLoc\$shortFile"
    $Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
$name
</title>
"@

    # get all the group members of every group and stick them in an HTML formatted table
    Get-DistributionGroup -Identity $alias | 
        Get-DistributionGroupMember -ResultSize unlimited | 
            Select-Object -Property Name, PrimarySmtpAddress, Office, Phone | 
                ConvertTo-Html -Property Name, PrimarySmtpAddress, Office, Phone -Head $Header | Out-File $outFile
    
    # Make an array of the friendly names of the groups and stick it in a link
    $li = "<li><a href='$shortFile'>$name</a></li>"
    $liList +=$li
}

# Get the members of each dynamic distribution group
foreach ($grp in $dynDistGroups){
    $alias = $grp.alias
    $name = $grp.name
    $shortFile = "$alias.html"
    $outFile = "$outFileLoc\$shortFile"
    $Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
$name
</title>
"@
    
    $blah = Get-DynamicDistributionGroup $alias 
    # need to account for broken dynamic lists so I don't have empty files    
    Get-Recipient -RecipientPreviewFilter $blah.RecipientFilter -ResultSize unlimited | 
        Select-Object -Property Name, PrimarySmtpAddress, Office, Phone |
            ConvertTo-Html -Property Name, PrimarySmtpAddress, Office, Phone -Head $Header | Out-File $outFile
    # Make an array of the friendly names of the groups and stick it in a link
    $li = "<li><a href='$shortFile'>$name</a></li>"
    $liList +=$li
}

# There's probably a better way to do this
# Piping $index start through Out-file clobbers the existing file
# that way if there's any new groups, they'll get added
# The CSS doesn't work, but I'm leaving it in just in case I figure out why it's borked

$indexStart = @"
<html>
 <head>
<link rel="stylesheet" type="text/css" href="index.css">
</head>
  <body>
    <ul>

"@ | Out-File -FilePath $indexFile

# Sort the group list alphabetically to prevent the dynamic lists from appearing at the bottom
$liList = $liList | Sort-Object
# append the array of group names
Add-Content -Path $indexFile -Value "$liList"

# append the closing html tags
$indexEnd = @"
   
    </ul>
  </body>
</html>
"@ | Add-Content -Path $indexFile
