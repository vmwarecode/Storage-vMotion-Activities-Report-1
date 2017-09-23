# Author: Dennis Zimmer, company: opvizor GmbH
# HTML structure based on:
# http://www.myexchangeworld.com/2010/03/powershell-disk-space-html-email-report/
# Script based on sVMotion history
# Source: http://www.lucd.info/2013/03/31/get-the-vmotionsvmotion-history/

# Variables
$from = "sender@"
$to = "receiver@"
$subject = "VMware SVMotion Activities - $Date"
$smtphost = "mailhost"
$svmotionRepFileName = "svmRep.html"
$vcenter = "localhost"

$result = @()
rm $svmotionRepFileName -force
New-Item -ItemType file $svmotionRepFileName -Force

# Initialize
#add-pssnapin vmware.vimautomation.core 
connect-viserver $vcenter


function Get-VIEventPlus {
<#  
.SYNOPSIS  Returns vSphere events   
.DESCRIPTION The function will return vSphere events. With 
  the available parameters, the execution time can be
  improved, compered to the original Get-VIEvent cmdlet.
.NOTES  Author:  Luc Dekens  
.PARAMETER Entity
  When specified the function returns events for the
  specific vSphere entity. 
.PARAMETER EventType
  This parameter limits the returned events to those
  specified on this parameter.
.PARAMETER Start
  The start date of the events to retrieve
.PARAMETER Finish
  The end date of the events to retrieve.
.PARAMETER Recurse
  A switch indicating if the events for the children of
  the Entity will also be returned
.PARAMETER FullMessage
  A switch indicating if the full message shall be compiled.
  This switch can improve the execution speed if the full
  message is not needed.  
.EXAMPLE
  PS> Get-VIEventPlus -Entity $vm
.EXAMPLE
  PS> Get-VIEventPlus -Entity $cluster -Recurse:$true
#>

  param(
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$Entity,
    [string[]]$EventType,
    [DateTime]$Start,
    [DateTime]$Finish = (Get-Date),
    [switch]$Recurse,
    [switch]$FullMessage = $false
  )

  process {
    $eventnumber = 100
    $events = @()
    $eventMgr = Get-View EventManager
    $eventFilter = New-Object VMware.Vim.EventFilterSpec
    $eventFilter.disableFullMessage = ! $FullMessage
    $eventFilter.entity = New-Object VMware.Vim.EventFilterSpecByEntity
    $eventFilter.entity.recursion = &{if($Recurse){"all"}else{"self"}}
    $eventFilter.eventTypeId = $EventType
    if($Start -or $Finish){
      $eventFilter.time = New-Object VMware.Vim.EventFilterSpecByTime
      $eventFilter.time.beginTime = $Start
      $eventFilter.time.endTime = $Finish
    }

    $entity | %{
      $eventFilter.entity.entity = $_.ExtensionData.MoRef
      $eventCollector = Get-View ($eventMgr.CreateCollectorForEvents($eventFilter))
      $eventsBuffer = $eventCollector.ReadNextEvents($eventnumber)
      while($eventsBuffer){
        $events += $eventsBuffer
        $eventsBuffer = $eventCollector.ReadNextEvents($eventnumber)
      }
      $eventCollector.DestroyCollector()
    }
    $events
  }
}

function Get-MotionHistory {
<#  
.SYNOPSIS  Returns the vMotion/svMotion history   
.DESCRIPTION The function will return information on all
  the vMotions and svMotions that occurred over a specific 
  interval for a defined number of virtual machines
.NOTES  Author:  Luc Dekens  
.PARAMETER Entity
  The vSphere entity. This can be one more virtual machines,
  or it can be a vSphere container. If the parameter is a 
  container, the function will return the history for all the
  virtual machines in that container.
.PARAMETER Days
  An integer that indicates over how many days in the past
  the function should report on.
.PARAMETER Hours
  An integer that indicates over how many hours in the past
  the function should report on.
.PARAMETER Minutes
  An integer that indicates over how many minutes in the past
  the function should report on.
.PARAMETER Sort
  An switch that indicates if the results should be returned
  in chronological order.
.EXAMPLE
  PS> Get-MotionHistory -Entity $vm -Days 1
.EXAMPLE
  PS> Get-MotionHistory -Entity $cluster -Sort:$false
.EXAMPLE
  PS> Get-Datacenter -Name $dcName |
  >> Get-MotionHistory -Days 7 -Sort:$false
#>

  param(
    [CmdletBinding(DefaultParameterSetName="Days")]
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$Entity,
    [Parameter(ParameterSetName='Days')]
    [int]$Days = 1,
    [Parameter(ParameterSetName='Hours')]
    [int]$Hours,
    [Parameter(ParameterSetName='Minutes')]
    [int]$Minutes,
    [switch]$Recurse = $false,
    [switch]$Sort = $true
  )

  begin{
    $history = @()
    switch($psCmdlet.ParameterSetName){
      'Days' {
        $start = (Get-Date).AddDays(- $Days)
      }
      'Hours' {
        $start = (Get-Date).AddHours(- $Hours)
      }
      'Minutes' {
        $start = (Get-Date).AddMinutes(- $Minutes)
      }
    }
    $eventTypes = "DrsVmMigratedEvent","VmMigratedEvent","VmRelocatedEvent"
  }

  process{
    $history += Get-VIEventPlus -Entity $entity -Start $start -EventType $eventTypes -Recurse:$Recurse |
    Select CreatedTime,
    @{N="Type";E={
        if($_.SourceDatastore.Name -eq $_.Ds.Name){"vMotion"}else{"svMotion"}}},
    @{N="UserName";E={if($_.UserName){$_.UserName}else{"System"}}},
    @{N="VM";E={$_.VM.Name}},
    @{N="SrcVMHost";E={$_.SourceHost.Name.Split('.')[0]}},
    @{N="TgtVMHost";E={if($_.Host.Name -ne $_.SourceHost.Name){$_.Host.Name.Split('.')[0]}}},
    @{N="SrcDatastore";E={$_.SourceDatastore.Name}},
    @{N="TgtDatastore";E={if($_.Ds.Name -ne $_.SourceDatastore.Name){$_.Ds.Name}}}
  }

  end{
    if($Sort){
      $history | Sort-Object -Property CreatedTime
    }
    else{
      $history
    }
  }
}
# Function to write the HTML Header to the file
Function writeHtmlHeader
{
param($svmotionRepFileName)
$date = ( get-date ).ToString('yyyy/MM/dd')
Add-Content $svmotionRepFileName "<html>"
Add-Content $svmotionRepFileName "<head>"
Add-Content $svmotionRepFileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $svmotionRepFileName '<title>opvizor - VMware Storage vMotion Report</title>'
add-content $svmotionRepFileName '<STYLE TYPE="text/css">'
add-content $svmotionRepFileName "<!--"
add-content $svmotionRepFileName "td {"
add-content $svmotionRepFileName "font-family: Tahoma;"
add-content $svmotionRepFileName "font-size: 12px;"
add-content $svmotionRepFileName "border-top: 1px solid #999999;"
add-content $svmotionRepFileName "border-right: 1px solid #999999;"
add-content $svmotionRepFileName "border-bottom: 1px solid #999999;"
add-content $svmotionRepFileName "border-left: 1px solid #999999;"
add-content $svmotionRepFileName "padding-top: 0px;"
add-content $svmotionRepFileName "padding-right: 0px;"
add-content $svmotionRepFileName "padding-bottom: 0px;"
add-content $svmotionRepFileName "padding-left: 0px;"
add-content $svmotionRepFileName "}"
add-content $svmotionRepFileName "body {"
add-content $svmotionRepFileName "margin-left: 5px;"
add-content $svmotionRepFileName "margin-top: 5px;"
add-content $svmotionRepFileName "margin-right: 0px;"
add-content $svmotionRepFileName "margin-bottom: 10px;"
add-content $svmotionRepFileName ""
add-content $svmotionRepFileName "table {"
add-content $svmotionRepFileName "border: thin solid #000000;"
add-content $svmotionRepFileName "}"
add-content $svmotionRepFileName "-->"
add-content $svmotionRepFileName "</style>"
Add-Content $svmotionRepFileName "</head>"
Add-Content $svmotionRepFileName "<body>"

add-content $svmotionRepFileName "<table width='100%'>"
add-content $svmotionRepFileName "<tr bgcolor='#CCCCCC'>"
add-content $svmotionRepFileName "<td colspan='7' height='25' align='center'>"
add-content $svmotionRepFileName "<font face='tahoma' color='#003399' size='4'><strong>VMware Storage vMotion Report - $date</strong></font>"
add-content $svmotionRepFileName "</td>"
add-content $svmotionRepFileName "</tr>"
add-content $svmotionRepFileName "</table>"

}

# Function to write the HTML Header to the file
Function writeTableHeader
{
param($svmotionRepFileName)
Add-Content $svmotionRepFileName "<table>"
Add-Content $svmotionRepFileName "<tr bgcolor=#CCCCCC>"
Add-Content $svmotionRepFileName "<td width='25%' align='center'>DataCenter</td>"
Add-Content $svmotionRepFileName "<td width='25%' align='center'>no. of svMotioned VMs</td>"
Add-Content $svmotionRepFileName "<td width='25%' align='center'>no. of svMotion tasks</td>"
Add-Content $svmotionRepFileName "<td width='25%' align='center'>svMotion TransferSize GB</td>"
Add-Content $svmotionRepFileName "</tr>"
}

Function writeHtmlFooter
{
param($svmotionRepFileName)

Add-Content $svmotionRepFileName "</body>"
Add-Content $svmotionRepFileName "</html>"
}

Function writeSvmInfo
{
param($svmotionRepFileName,$dcname, $vmcount, $svmotioncount, $svmotionsizegb)

 Add-Content $svmotionRepFileName "<tr>"
 Add-Content $svmotionRepFileName "<td>$dcname</td>"
 Add-Content $svmotionRepFileName "<td>$vmcount</td>"
 Add-Content $svmotionRepFileName "<td>$svmotioncount</td>"
 Add-Content $svmotionRepFileName "<td>$svmotionsizegb</td>"
 Add-Content $svmotionRepFileName "</tr>"
 }
 
 Add-Content $svmotionRepFileName "</tr>"

Function sendEmail
{ param($from,$to,$subject,$smtphost,$htmlFileName)
$body = Get-Content $htmlFileName
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
$msg.isBodyhtml = $true
$smtp.send($msg)
}

# Header
writeHtmlHeader $svmotionRepFileName

# Table 
 writeTableHeader $svmotionRepFileName
$result = @()
 foreach ($dc in get-datacenter) {
	$sum = 0
	$vms = $dc | get-vm
	$vmsview = $vms | get-view
	$vmotions = get-motionhistory -Entity ($vms) -days 7 | ?{$_.type -eq "svMotion"}
		$row = "" | select dcname, vmcount, svmotioncount, svmotionsizegb
		$row.dcname = $dc.name
		$row.vmcount = ($vmotions | select vm -Unique | measure).count
		$row.svmotioncount = ($vmotions | measure).count
		
		foreach ($vmotion in $vmotions) {
			$sum += [int](((($vmsview | ?{$_.name -eq $vmotion.vm}).layoutex.file | measure -Property size -sum).sum)/1GB)
		}
		$row.svmotionsizegb = $sum
		
	# write Report
	$result += $row

	if ($vmotions) {writeSvmInfo $svmotionRepFileName $row.dcname $row.vmcount $row.svmotioncount $row.svmotionsizegb}
	}
 # Add final row of table
	Add-Content $svmotionRepFileName "</table>"

# Footer
	writeHtmlFooter $svmotionRepFileName

	$date = ( get-date ).ToString('yyyy/MM/dd')

# Send Email	
sendEmail $from $to $subject $smtphost $svmotionRepFileName

disconnect-viserver -confirm:$false