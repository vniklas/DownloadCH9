﻿# Download MS Ignite Content with BITS
# 
# Niklas Akerlund v 0.4 2015-05-05
# 
# Borrowed code for making the folders from Peter Schmidt (Exchange MVP, blog: www.msdigest.net) DownloadTechEdEurope14VideoAndSlides.ps1
#
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

  [switch]$PPT,
  [switch]$MP4,
  [string]$Dest='C:\Ignitetest')

# Check if the folder exists
if(!(Get-Item $Dest -ErrorAction Ignore))
{
  New-Item -Path $Dest -ItemType Directory
}
$psessions = @()
$vsessions = @()
#$ = 'C:\techedtest'

if($CH9){
  $psessions =  Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides'  | Where-Object comments -cmatch "C9"
  $vsessions =  Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high'  | Where-Object comments -cmatch "C9"
}
if($KEY){
  $psessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' | Where-Object comments -cmatch "KEY"
  $vsessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' | Where-Object comments -cmatch "KEY"
}
if($FDN){
  $psessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' | Where-Object comments -cmatch "FDN"
  $vsessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' | Where-Object comments -cmatch "FDN"
}
if($DEV){
  $psessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' | Where-Object category -Contains "development"
  $vsessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' | Where-Object category -Contains "development"
}
if($HYBRID){
  $psessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' | Where-Object category -Contains "hybrid"
  $vsessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' | Where-Object category -Contains "hybrid"
}
if($VOICE){
  $psessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' | Where-Object category -Contains "voice"
  $vsessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' | Where-Object category -Contains "voice"
}
if($CLOUD){
  $psessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' | Where-Object category -Contains "cloud"
  $vsessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' | Where-Object category -Contains "cloud"
}
if($IAAS){
  $psessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' | Where-Object category -Contains "infrastructure-as-a-service"
  $vsessions += Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' | Where-Object category -Contains "infrastructure-as-a-service"
}


if ($All){
  $psessions =  Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' 
  $vsessions = Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' 
}

#$psessions 

if($PPT){
  foreach ($psession in $psessions){
      # create folder
      $code = $psession.comments.split("/") | Select-Object -last 1	
      $folder = $code + " - " + $psession.title.Replace(":", "-").Replace("?", "").Replace("/", "-").Replace("<", "").Replace("|", "").Replace('"',"").Replace("*","")
		  $folder = $folder.substring(0, [System.Math]::Min(100, $folder.Length))
		  $folder = $folder.trim()
      $folder = $Dest + "\" + $folder
      if(!(Get-Item $folder -ErrorAction Ignore)){
          New-Item -Path $folder -ItemType Directory
      }
      #tage pptx
      [string]$pptx = $psession.GetElementsByTagName('enclosure').url 
      if(!(get-item ($folder +"\" + $code +".pptx") -ErrorAction Ignore)){
        Start-BitsTransfer -Source $pptx -Destination $folder -DisplayName "PPT $Code" -Description $folder
      } else{
        Write-Output " $code ppt already downloaded"
      }
  }
}
if($MP4){
  foreach ($vsession in $vsessions){
      $code = $vsession.comments.split("/") | Select-Object -last 1	
      $folder = $code + " - " + $vsession.title.Replace(":", "-").Replace("?", "").Replace("/", "-").Replace("<", "").Replace("|", "").Replace('"',"").Replace("*","")
		  $folder = $folder.substring(0, [System.Math]::Min(100, $folder.Length))
		  $folder = $folder.trim()
      $folder = $Dest + "\" + $folder
      if(!(Get-Item $folder -ErrorAction Ignore)){
          New-Item -Path $folder -ItemType Directory
      }
      [string]$video = $vsession.GetElementsByTagName('enclosure').url
      #$video
      if(!(get-item ($folder +"\" + $code +".mp4") -ErrorAction Ignore)){
        Start-BitsTransfer -Source $video -Destination $folder -DisplayName "MP4 $Code" -Description $folder
      }else{
        Write-Output " $code video already downloaded"
      }
  }
}