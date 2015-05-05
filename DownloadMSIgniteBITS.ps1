﻿# Download MS Ignite Content with BITS
# 
# Niklas Akerlund v 0.3 2015-05-03
# 
# Borrowed code for making the folders from Peter Schmidt (Exchange MVP, blog: www.msdigest.net) DownloadTechEdEurope14VideoAndSlides.ps1
#
[CmdletBinding()]
param(
  [switch]$CH9,
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
  $psessions =  Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides'  | where comments -cmatch "C9"
  $vsessions =  Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high'  | where comments -cmatch "C9"
}

if ($All){
  $psessions =  Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/slides' 
  $vsessions = Invoke-RestMethod 'http://s.ch9.ms/Events/Ignite/2015/RSS/mp4high' 
}

#$psessions 

if($PPT){
  foreach ($psession in $psessions){
      # create folder
      $code = $psession.comments.split("/") | select -last 1	
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
      $code = $vsession.comments.split("/") | select -last 1	
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