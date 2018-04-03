#Exit 

function ConvertTo-DataTable
{
	[CmdLetBinding(DefaultParameterSetName="None")]
	param(
	 [Parameter(Position=0,Mandatory=$true)][System.Array]$Source,
	 [Parameter(Position=1,ParameterSetName='Like')][String]$Match=".+",
	 [Parameter(Position=2,ParameterSetName='NotLike')][String]$NotMatch=".+"
	)
	if ($NotMatch -eq ".+"){
	$Columns = $Source[0] | Select * | Get-Member -MemberType NoteProperty | Where-Object {$_.Name -match "($Match)"}
	}
	else {
	$Columns = $Source[0] | Select * | Get-Member -MemberType NoteProperty | Where-Object {$_.Name -notmatch "($NotMatch)"}
	}
	$DataTable = New-Object System.Data.DataTable
	foreach ($Column in $Columns.Name)
	{
	 $DataTable.Columns.Add("$($Column)") | Out-Null
	}
	#For each row (entry) in source, build row and add to DataTable.
	foreach ($Entry in $Source)
	{
	 $Row = $DataTable.NewRow()
	 foreach ($Column in $Columns.Name)
	 {
	 $Row["$($Column)"] = if($Entry.$Column -ne $null){($Entry | Select-Object -ExpandProperty $Column) -join ', '}else{$null}
	 }
	 $DataTable.Rows.Add($Row)
	}
	#Validate source column and row count to DataTable
	if ($Columns.Count -ne $DataTable.Columns.Count){
	 throw "Conversion failed: Number of columns in source does not match data table number of columns"
	}
	else{ 
	 if($Source.Count -ne $DataTable.Rows.Count){
	 throw "Conversion failed: Source row count not equal to data table row count"
	 }
	 #The use of "Return ," ensures the output from function is of the same data type; otherwise it's returned as an array.
	 else{
	 Return ,$DataTable
	 }
	 }
}

function PrintTable([System.Data.DataTable] $Dt)
{
	if ($Dt -eq $null -Or $Dt.Rows.Count -eq 0){
		return "\n| No Errors.... |\n| --- |"
	
	}
	$TableStr = "|"
	$TableStr2 = "|"
	foreach ($col in $Dt.Columns)
	{
		$TableStr += " " + $col.Caption + " |"
		$TableStr2 += " :---: |"
	}

	$RowStr = ""
	foreach ($row in $Dt.Rows){
		$RowStr += "|"
		foreach ($col in $Dt.Columns)
		{
			$RowStr += " " + $row[$col.Caption] + " |"
		}
		$RowStr += "\n"
	}

	$out ="\n" + $TableStr + "\n" + $TableStr2 + "\n" + $RowStr
	return $out
}



Function Execute-Command ($commandPath, $commandArguments, $timeOutMin = 50)
{
	$info = New-Object System.Diagnostics.ProcessStartInfo -Property @{
				"FileName" = "$commandPath"
				"Arguments" = "$commandArguments"
				"UseShellExecute" = $false
				"RedirectStandardOutput" = $true
                "UserName" = "appveyor"
                "PasswordInClearText"  = "r@@zbeh1234"
				}
	$pr = New-Object System.Diagnostics.Process
	$pr.StartInfo = $info

	Register-ObjectEvent -InputObject $pr -EventName OutputDataReceived -action {Write-Output $Event.SourceEventArgs.Data} | Out-Null
	Register-ObjectEvent -InputObject $pr -EventName Exited -action {$global:myprocessrunning = $false} | Out-Null

	$global:myprocessrunning = $true
	$process = $pr.start()
	$pr.BeginOutputReadLine()

	$timeoutseconds = $timeOutMin * 60 # testing should be done in 50minutes!
	$processTimeout = $timeoutseconds * 1000
	while (($global:myprocessrunning -eq $true) -and ($processTimeout -gt 0)) {
		# We must use lots of shorts sleeps rather than a single long one otherwise events are not processed
		$processTimeout -= 50
		Start-Sleep -m 50
	}
	if ($processTimeout -le 0) {
		Write-Verbose "Process exceeded its timeout and forcefully killed!" -Verbose
		$pr.Kill()
	}
}

Function Execute-CommandShell ($commandPath, $commandArguments, $timeOutMin = 50)
{
	$info = New-Object System.Diagnostics.ProcessStartInfo -Property @{
				"FileName" = "$commandPath"
				"Arguments" = "$commandArguments"
				"UseShellExecute" = $true
				"RedirectStandardOutput" = $false
				}
	$pr = New-Object System.Diagnostics.Process
	$pr.StartInfo = $info

	Register-ObjectEvent -InputObject $pr -EventName OutputDataReceived -action {Write-Output $Event.SourceEventArgs.Data} | Out-Null
	Register-ObjectEvent -InputObject $pr -EventName Exited -action {$global:myprocessrunning = $false} | Out-Null

	$global:myprocessrunning = $true
	$process = $pr.start()
	#$pr.BeginOutputReadLine()

	$timeoutseconds = $timeOutMin * 60 # testing should be done in 50minutes!
	$processTimeout = $timeoutseconds * 1000
	while (($global:myprocessrunning -eq $true) -and ($processTimeout -gt 0)) {
		# We must use lots of shorts sleeps rather than a single long one otherwise events are not processed
		$processTimeout -= 50
		Start-Sleep -m 50
	}
	if ($processTimeout -le 0) {
		Write-Verbose "Process exceeded its timeout and forcefully killed!" -Verbose
		$pr.Kill()
	}
}


$initial_json = @"
{
  "state": "pending",
  "target_url": "https://some url.com",
  "description": "",
  "context": "CI/AWSTests"
}
"@

$jgateway = ConvertFrom-Json $initial_json
$jgateway.description = "BehaviouralGateway tests pending."
$jgateway.context = "CI/AWS/BehaviouralGateway"

$jsipp = ConvertFrom-Json $initial_json
$jsipp.description = "SipP tests pending."
$jsipp.context = "CI/AWS/SipP"

$jMSAutoRun = ConvertFrom-Json $initial_json
$jMSAutoRun.description = "MediaServerAutoRun tests pending."
$jMSAutoRun.context = "CI/AWS/MediaServerAutoRun"

$jAAAutoRun = ConvertFrom-Json $initial_json
$jAAAutoRun.description = "AudioAnalyzerAutoRun tests pending."
$jAAAutoRun.context = "CI/AWS/AudioAnalyzerAutoRun"

$jMSService = ConvertFrom-Json $initial_json
$jMSService.description = "MediaServer Service Installation test pending."
$jMSService.context = "CI/AWS/MediaServerService"



$github_access_token = '?access_token=xxxxxxxxxxxxxxx'




# here you are Normal Media Server!
	Write-Verbose "Adding github statuses..." -Verbose
	
##	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jgateway) -Uri $github_status_url  -UseBasicParsing | Out-Null
##	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jsipp) -Uri $github_status_url  -UseBasicParsing | Out-Null
##	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jMSAutoRun) -Uri $github_status_url -UseBasicParsing | Out-Null
##	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jAAAutoRun) -Uri $github_status_url -UseBasicParsing | Out-Null
##	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jMSService) -Uri $github_status_url -UseBasicParsing | Out-Null
	


	$jMSService.state = "success"
	$jMSService.description = "MediaServer Service Start-Stop passed. :heavy_check_mark:"
##	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jMSService) -Uri $github_status_url -UseBasicParsing | Out-Null
	Start-Sleep -s 5
	

	$noindicator = ":grey_exclamation:"
	$bechmark_result | Add-Member -MemberType NoteProperty "Avg %D" -Value $noindicator
	$bechmark_result | Add-Member -MemberType NoteProperty "Max %D" -Value $noindicator

	$benchmark_table = ConvertTo-DataTable -Source $bechmark_result

	Write-Verbose "Benchmark Done." -Verbose



$comment_body = @"
{
  "body": ""
}
"@

$js_comment_body = ConvertFrom-Json $comment_body


	$testresults = @{
        "Sipp" = @{
            passed = 0
            failed = 0
			ignored = 0
            }
        "BGW" = @{
            passed = 0
            failed =0
			ignored = 0
            }
        "MSAR" = @{
            passed = 0
            failed =0
			ignored = 0
            }
        "AAAR" = @{
            passed = 0
            failed =0
			ignored = 0
            }
    }
	
	
	$jsipp.state = "success"
	$jsipp.description = "SipP tests finished. Passed : " + $testresults["Sipp"].passed + " Failed: " + $testresults["Sipp"].failed +  " Ignored: " + $testresults["Sipp"].ignored + " ."
	if ( ($testresults["Sipp"].failed -gt 0) -or ( ($testresults["Sipp"].passed -eq 0) -and ($testresults["Sipp"].ignored -eq 0) )){
		$jsipp.state = "failure"
		$jsipp.description = $jsipp.description + " :x:"
    }
	else {
		$jsipp.description = $jsipp.description + " :heavy_check_mark:"
	}
	
	$jgateway.state = "success"
	$jgateway.description = "BehaviouralGateway tests finished. Passed : " + $testresults["BGW"].passed + " Failed: " + $testresults["BGW"].failed + " Ignored: " + $testresults["BGW"].ignored + " ."
	if ( ($testresults["BGW"].failed -gt 0) -or ( ($testresults["BGW"].passed -eq 0) -and ($testresults["BGW"].ignored -eq 0) )){
		$jgateway.state = "failure"
		$jgateway.description = $jgateway.description + " :x:"
    }
	else {
		$jgateway.description = $jgateway.description + " :heavy_check_mark:"
	}

	$jMSAutoRun.state = "success"
	$jMSAutoRun.description = "MediaServerAutoRun tests finished. Passed : " + $testresults["MSAR"].passed + " Failed: " + $testresults["MSAR"].failed + " Ignored: " + $testresults["MSAR"].ignored + " ."
	if ( ($testresults["MSAR"].failed -gt 0) -or ( ($testresults["MSAR"].passed -eq 0) -and ($testresults["MSAR"].ignored -eq 0) )){
		$jMSAutoRun.state = "failure"
		$jMSAutoRun.description = $jMSAutoRun.description + " :x:"
    }
	else {
		$jMSAutoRun.description = $jMSAutoRun.description + " :heavy_check_mark:"
	}
	
	$jAAAutoRun.state = "success"
	$jAAAutoRun.description = "AudioAnalyzerAutoRun tests finished. Passed : " + $testresults["AAAR"].passed + " Failed: " + $testresults["AAAR"].failed + " Ignored: " + $testresults["AAAR"].ignored + " ."
	if ( ($testresults["AAAR"].failed -gt 0) -or ( ($testresults["AAAR"].passed -eq 0) -and ($testresults["AAAR"].ignored -eq 0) )){
		$jAAAutoRun.state = "failure"
		$jAAAutoRun.description = $jAAAutoRun.description + " :x:"
    }
	else {
		$jAAAutoRun.description = $jAAAutoRun.description + " :heavy_check_mark:"
	}
	
#	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jgateway) -Uri $github_status_url -UseBasicParsing | Out-Null
#	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jsipp) -Uri $github_status_url  -UseBasicParsing | Out-Null
#	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jMSAutoRun) -Uri $github_status_url -UseBasicParsing | Out-Null
#	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jAAAutoRun) -Uri $github_status_url -UseBasicParsing | Out-Null

	

If ($comments_url -ne $null)
{
	$dtEvents = New-Object System.Data.Datatable
	[void]$dtEvents.Columns.Add("Source")
	[void]$dtEvents.Columns.Add("FirstInstance")
	[void]$dtEvents.Columns.Add("Count")
	[void]$dtEvents.Columns.Add("FirstTime")
	
	$AppEvents = Get-EventLog -LogName "Application" -AsBaseObject
	foreach ($AppEvent in $AppEvents){
		if ( ($AppEvent.EntryType -eq "Error") ){
			[void]$dtEvents.Rows.Add($AppEvent.Source,$AppEvent.Message.Replace("`n"," ").subString(0, [System.Math]::Min(162, $AppEvent.Message.Length)) ,"0",$AppEvent.TimeWritten)
			}

	}
	
	$hash.DataTable.Merge($dtEvents)

	$found_AWS_Comment = $null
	$comments_list = (Invoke-WebRequest  -Method GET -Uri ($comments_url + $github_access_token) -UseBasicParsing).Content | ConvertFrom-Json

	ForEach ($comment in $comments_list)
	{
		if ($comment.body -like "# AWS Test Results*")
		{
			$comments_url = $comment.url
			$found_AWS_Comment = $comment
			break
		}
	}

	$diskusage = (($diskbefore.Free - $diskafter.Free) / 1024)
	$diskfilesusage = $diskfilesafter - $diskfilesbefore
	
	if ($Env:My_MediaServerMode -eq "Normal"){
		if ($diskusage -ge 450000) {$diskusage_alertcode = "&#x1F621;"} else {$diskusage_alertcode = "&#x1F60A;"}
		if ($diskfilesusage -ge 3000) {$diskfileusage_alertcode = "&#x1F621;"} else {$diskfileusage_alertcode = "&#x1F60A;"}
		
		$js_DiskUsage = [string]::Format("Disk usage was {0:n0}Kb $diskusage_alertcode \n New files created are {1:n0} $diskfileusage_alertcode",$diskusage,$diskfilesusage)

		$Report_XML_Results = (Get-Content $ReportfilePath) | Out-String
		$commet_downloads = "Download report of [Standalone MRC Client](https://${jfrog_user}:$jfrog_pass@$jfrogurl_nofile/Normal/reports.zip) \n Download report of [TTS Media Server](https://${jfrog_user}:$jfrog_pass@$jfrogurl_nofile/TTS/reports.zip) \n "
		$js_comment_xml = "Followings are test cases with errors ( Download [full xml](https://${jfrog_user}:$jfrog_pass@$jfrogurl_nofile) for list of all passed and failed.) : \n ``````xml $Report_XML_Results``````"
		$js_comment_MSErrors = "### Followings are Standalone MRC Client Errors :\n" + "\n\n" + $js_DiskUsage + "\n" 
		$js_comment_table = PrintTable $benchmark_table
		$js_comment_perf = "Performance Index is " + $bechmark_perfIndex_text + "\n\nBenchmark results of Single MRC Client are :\n" + $js_comment_table + "\n"
		$js_comment_MSTTSErrors = "### Followings are TTS Media Server Errors :\n TTS_REPORTS_PENDING..."
		$js_comment_body.body = "# AWS Test Results (" + ($Env:MY_APPVEYOR_PULL_REQUEST_HEAD_COMMIT).SubString(0,7) + ")\n"
		$js_comment_body.body = $js_comment_body.body + "Testing finished at $(Get-Date  -format g)\n\n"
		$js_comment_body.body = $js_comment_body.body + $jMSService.description + " \n " + $jsipp.description + " \n " + $jgateway.description + " \n " + $jMSAutoRun.description + " \n " + $jAAAutoRun.description+" \n\n $commet_downloads \n\n $js_comment_xml\n" + $js_comment_MSErrors + $js_comment_perf + "\n" + $js_comment_MSTTSErrors
	##	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $js_comment_body).replace("\\n","\n") -Uri ($comments_url + $github_access_token) -UseBasicParsing | ConvertFrom-Json
	}
	elseif ($found_AWS_Comment -ne $null){
		if ($diskusage -ge 50000) {$diskusage_alertcode = "&#x1F621;"} else {$diskusage_alertcode = "&#x1F60A;"}
		if ($diskfilesusage -ge 100) {$diskfileusage_alertcode = "&#x1F621;"} else {$diskfileusage_alertcode = "&#x1F60A;"}
		
		$js_DiskUsage = [string]::Format("Disk usage was {0:n0}Kb $diskusage_alertcode \n New files created are {1:n0} $diskfileusage_alertcode",$diskusage,$diskfilesusage)
		
		$js_comment_body.body = ($found_AWS_Comment.body).replace("TTS_REPORTS_PENDING...","\n\n $js_DiskUsage")
##		Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $js_comment_body).replace("\\n","\n") -Uri ($comments_url + $github_access_token) -UseBasicParsing | ConvertFrom-Json
	}
	
}

