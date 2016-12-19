$Now = Get-Date
$Days = "30"
$TargetFolder = "\\pharmanas1\landingzone\SCCMTeam\DeleteSpecialReports"
$Extension = "*.mhtml"
$LastWrite = $Now.AddDays(-$Days)
$Files = Get-Childitem $TargetFolder -Include $Extension -Recurse | Where {$_.LastWriteTime -le "$LastWrite"}
foreach ($File in $Files) 
  {
  if ($File -ne $NULL)
      {
      Remove-Item $File.FullName | out-null
      }
  else
      {
      }
  }