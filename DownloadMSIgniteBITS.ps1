# Download MS Ignite Content with BITS
# 
# Niklas Akerlund v 0.6 2015-05-18
# 
# Borrowed code for making the folders from Peter Schmidt (Exchange MVP, blog: www.msdigest.net) DownloadTechEdEurope14VideoAndSlides.ps1
# Thanks to Markus Bäker for fixing some issues with the code for file name and -ALL !! 

[CmdletBinding()]
param(
  [switch]$CH9,
  [switch]$KEY,
  [switch]$FDN,
  [switch]$DEV,
  [switch]$HYBRID,
  [switch]$VOICE,
  [switch]$CLOUD,
  [switch]$IAAS,
  [switch]$AUTOMATION,
  [switch]$ALL,

  [switch]$PPT,
  [switch]$MP4,
  [string]$Dest='C:\Ignitetest')

# Check if the folder exists
if(!(Get-Item $Dest -ErrorAction Ignore))
{
  New-Item -Path $Dest -ItemType Directory
}

#download arrays 
$dlpsessions = @()
$dlvsessions = @()
# Get all sessions to massage 
$psessions = @()
$vsessions = @()

$psessions =  Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' 
$psessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides?page=2'
$vsessions = Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' 
$vsessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high?Page=2'
#$ = 'C:\techedtest'

if($CH9){
  $dlpsessions = $psessions  | Where-Object comments -cmatch 'C9'
  $dlvsessions = $vsessions  | Where-Object comments -cmatch 'C9'
}
if($KEY){
  $dlpsessions += $psessions | Where-Object comments -cmatch 'KEY'
  $dlvsessions += $vsessions | Where-Object comments -cmatch 'KEY'
}
if($FDN){
  $dlpsessions += $psessions | Where-Object comments -cmatch 'FDN'
  $dlvsessions += $vsessions | Where-Object comments -cmatch 'FDN'
}
if($DEV){
  $dlpsessions += $psessions | Where-Object category -Contains 'development'
  $dlvsessions += $vsessions | Where-Object category -Contains 'development'
}
if($HYBRID){
  $dlpsessions += $psessions | Where-Object category -Contains 'hybrid'
  $dlvsessions += $vsessions | Where-Object category -Contains 'hybrid'
}
if($VOICE){
  $dlpsessions += $psessions | Where-Object category -Contains 'voice'
  $dlvsessions += $vsessions | Where-Object category -Contains 'voice'
}
if($CLOUD){
  $dlpsessions += $psessions | Where-Object category -Contains 'cloud'
  $dlvsessions += $vsessions | Where-Object category -Contains 'cloud'
}
if($IAAS){
  $dlpsessions += $psessions | Where-Object category -Contains 'infrastructure-as-a-service'
  $dlvsessions += $vsessions | Where-Object category -Contains 'infrastructure-as-a-service'
}
if($AUTOMATION){
  $dlpsessions += $psessions | Where-Object category -Contains 'automation'
  $dlvsessions += $vsessions | Where-Object category -Contains 'automation'
}


if ($All){
    $dlpsessions = $psessions
    $dlvsessions = $vsessions
}

#$psessions 

if($PPT){
  foreach ($dlpsession in $dlpsessions){
      # create folder
      $code = $dlpsession.comments.split('/') | Select-Object -last 1	
      $folder = $code + ' - ' + $dlpsession.title.Replace(':', '-').Replace('?', '').Replace('/', '-').Replace('<', '').Replace('|', '').Replace('"','').Replace('*','')
		  $folder = $folder.substring(0, [System.Math]::Min(100, $folder.Length))
		  $folder = $folder.trim()
      $folder = join-path $Dest $folder
      if(!(Get-Item $folder -ErrorAction Ignore)){
          New-Item -Path $folder -ItemType Directory
      }
      #tage pptx
      [string]$pptx = $dlpsession.GetElementsByTagName('enclosure').url 
      $target=join-path $folder ($code+".pptx") 
      if(!(get-item ($folder +'\' + $code +'.pptx') -ErrorAction Ignore)){
        Start-BitsTransfer -Source $pptx -Destination $target -DisplayName "PPT $Code" -Description $folder
      } else{
        Write-Output " $code ppt already downloaded"
      }
  }
}
if($MP4){
  foreach ($dlvsession in $dlvsessions){
      $code = $dlvsession.comments.split('/') | Select-Object -last 1	
      $folder = $code + ' - ' + $dlvsession.title.Replace(':', '-').Replace('?', '').Replace('/', '-').Replace('<', '').Replace('|', '').Replace('"','').Replace('*','')
		  $folder = $folder.substring(0, [System.Math]::Min(100, $folder.Length))
		  $folder = $folder.trim()
      $folder = join-path $Dest $folder
      if(!(Get-Item $folder -ErrorAction Ignore)){
          New-Item -Path $folder -ItemType Directory
      }
      [string]$video = $dlvsession.GetElementsByTagName('enclosure').url
      $target=join-path $folder ($code+".mp4") 
      #$video
      if(!(get-item ($folder +'\' + $code +'.mp4') -ErrorAction Ignore)){
        Start-BitsTransfer -Source $video -Destination $target -DisplayName "MP4 $Code" -Description $folder
      }else{
        Write-Output " $code video already downloaded"
      }
  }
}