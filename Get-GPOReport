#################################
#### Morgan Atwood 5/25/2017 ####
#################################

[String]$date = Get-Date -Format y
$date = $date -replace ','


#-----------------------------------------------
# Functions
#-----------------------------------------------

#Get-GPOInfo gets an XML report of the GPO and uses it to return specific data in an array
function Get-GPOInfo
{
 param($GPOGUID)

 #Gets the XML version of the GPO Report
 $GPOReport = get-gporeport -guid $GPOGUID -reporttype XML
 #Converts it to an XML variable for manipulation
 $GPOXML = [xml]$GPOReport
 
 #Create array to store info
 $GPOInfo = @() 

 #Get's info from XML and adds to array
 #General Information
 $Name = $GPOXML.GPO.Name
 $GPOInfo += , $Name
 
 $GUID = $GPOXML.GPO.Identifier.Identifier.'#text'
 $GPOInfo += , $GUID
 
 [DateTime]$Created = $GPOXML.GPO.CreatedTime
 $GPOInfo += , $Created.ToString("G")
 
 [DateTime]$Modified = $GPOXML.GPO.ModifiedTime
 $GPOInfo += , $Modified.ToString("G")
 
 
 #Computer Configuration
 $ComputerEnabled = $GPOXML.GPO.Computer.Enabled
 $GPOInfo += , $ComputerEnabled
 #Links
 if ($GPOXML.GPO.LinksTo) {
  $Links = $GPOXML.GPO.LinksTo | %{ $_.SOMPath }
  $Links = [string]::join("`n", $Links)
  $LinksEnabled = $GPOXML.GPO.LinksTo | %{ $_.Enabled }
  $LinksEnabled = [string]::join("`n", $LinksEnabled)
  $LinksNoOverride = $GPOXML.GPO.LinksTo | %{ $_.NoOverride }
  $LinksNoOverride = [string]::join("`n", $LinksNoOverride)
 } else {
  $Links = "<none>"
  $LinksEnabled = "<none>"
  $LinksNoOverride = "<none>"
 }
 $GPOInfo += , $Links

 
 #Security Info

 #$Owner = $GPOXML.GPO.SecurityDescriptor.Owner.Name.'#text'
 #$GPOInfo += , $Owner

# $SecurityFilter = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions.Trustee.Name
 #$SecurityFilter = [Sting]::join("`n",  $SecurityFilter)
 #$GPOInfo += , $SecurityFilter
 
 #$SecurityInherits = $GPOXML.GPO.SecurityDescriptor.Permissions.InheritsFromParent
 #$SecurityInherits = [string]::join("`n", $SecurityInherits)
 #$GPOInfo += , $SecurityInherits
 
 $SecurityGroups = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | %{ $_.Trustee.Name.'#text' }
 $SecurityGroups = [string]::join("`n", $SecurityGroups)
 $GPOInfo += , $SecurityGroups
 
 #$SecurityType = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | % { $_.Type.PermissionType }
# $SecurityType = [string]::join("`n", $SecurityType)
 #$GPOInfo += , $SecurityType
 
# $SecurityPerms = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | % { $_.Standard.GPOGroupedAccessEnum }
 #$SecurityPerms = [string]::join("`n", $SecurityPerms)
 #$GPOInfo += , $SecurityPerms
 
    return $GPOInfo
 
}



#-----------------------------------------------
# Get's list of GPO's
#-----------------------------------------------

write-host Getting GPO Information...
$GPOs = get-gpo -all |  Where-Object {$_.displayname -like 'ENG*'} 
write-host `tGPOs: $GPOs.Count

#-----------------------------------------------
# Creates an array and populates it with GPO information arrays
#-----------------------------------------------

$AllGPOs = @()

write-host Getting GPO XML Reports...

$GPOCount = 0
$GPOs | foreach-object {

 $GPOCount++
 write-host `t$GPOCount : $_.DisplayName / $_.ID
 $GPOGUID = $_.ID
 $ThisGPO = get-gpoinfo $_.ID
 $AllGPOs += ,$ThisGPO

}

#-----------------------------------------------------
# Exports all information to Excel (nicely formatted)
#-----------------------------------------------------

write-host Exporting information to Excel...

#Excel Constants
$White = 2
$Bluegrass = 50
$Center = -4108
$Top = -4160

$e = New-Object -comobject Excel.Application
$e.Visible = $false #Change to Hide Excel Window
$e.DisplayAlerts = $False
$wkb = $E.Workbooks.Add()
$wks = $wkb.Worksheets.Item(1)

#Builds Top Row
$wks.Cells.Item(1,1) = "GPO Summary Report $date"
$wks.Cells.Item(2,1) = "GPO Name"
$wks.Cells.Item(2,2) = "GUID"
$wks.Cells.Item(2,3) = "Created"
$wks.Cells.Item(2,4) = "Last Modified"
$wks.Cells.Item(2,5) = "Enabled"
$wks.Cells.Item(2,6) = "Links"
$wks.Cells.Item(2,7) = "Secuirty Filter"



#Formats Top Row
$wks.Range("A1:Y1").font.bold = "true"
$wks.Range("A1:Y1").font.size = "24"
$wks.Range("A1:Y1").font.ColorIndex = $Bluegrass 
$wks.Range("A1:Y1").interior.ColorIndex = $White

$wks.Range("A2:Y2").font.bold = "true"
$wks.Range("A2:Y2").font.ColorIndex = $White
$wks.Range("A2:Y2").interior.ColorIndex = $Bluegrass

#Fills in Data from Array
$row = 3
$AllGPOs | foreach {
  $wks.Cells.Item($row,1) = $_[0]
  $wks.Cells.Item($row,2) = $_[1]
  $wks.Cells.Item($row,3) = $_[2]
  $wks.Cells.Item($row,4) = $_[3]
  $wks.Cells.Item($row,5) = $_[4]
  $wks.Cells.Item($row,6) = $_[5]
  $wks.Cells.Item($row,7) = $_[6]
  $wks.Cells.Item($row,8) = $_[7]
  $wks.Cells.Item($row,9) = $_[8]
  $wks.Cells.Item($row,10) = $_[9]
  $wks.Cells.Item($row,11) = $_[10]
  $wks.Cells.Item($row,12) = $_[11]
  $wks.Cells.Item($row,13) = $_[12]
  $wks.Cells.Item($row,14) = $_[13]
  $wks.Cells.Item($row,15) = $_[14]
  $row++
}

#Adjust Formatting to make it easier to read
$wks.Range("F:F").Columns.ColumnWidth = 150
$wks.Range("G:G").Columns.ColumnWidth = 150


[void]$wks.Range("A:Y").Columns.AutoFit()
$wks.Range("A:U").Columns.VerticalAlignment = $Top
$wks.Range("F:H").Columns.HorizontalAlignment = $Center
$wks.Range("J:L").Columns.HorizontalAlignment = $Center
$wks.Range("R:R").Columns.HorizontalAlignment = $Center
$wks.Range("V:X").Columns.HorizontalAlignment = $Center

#Save the file
$SaveFile = "C:\scripts\gpo_report\output\GPO Report $date.xlsx"
$wkb.SaveAs($SaveFile)

$e.Quit()

Start-Sleep 5


Write-Host "Sending Mail"


$emailaddress = ""
$from = ""
$smtpServer = ""
$subject = "GPO Summary Report"


send-mailmessage $emailaddress $subject -Attachment $SaveFile -from $from -SmtpServer $smtpServer -priority High 
