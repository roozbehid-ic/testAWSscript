#Exit 
Start-Transcript -Path "C:\StartCodeCov_AWS.txt"
Import-Module -Name C:\ProgramData\Amazon\EC2-Windows\Launch\Module\Ec2Launch.psd1

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

function Disable-InternetExplorerESC {
    $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
    $UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
	$WizardKey = "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main"
    Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 0
    Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 0
	New-Item -Path $WizardKey -Force
	Set-ItemProperty -Path $WizardKey -Name "DisableFirstRunCustomize" -Value 1
    #Stop-Process -Name Explorer | Out-Null
    Write-Output "IE Enhanced Security Configuration (ESC) has been disabled." 
}

Disable-InternetExplorerESC

Write-Output "Disable firewall on private and public domains." 
Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False

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



 
# if you change it here also change it in StartCodeCo_AWS.ps1
$user = "build"
$pass= "xxxxx"
$secpasswd = ConvertTo-SecureString $pass -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($user, $secpasswd)

$jfrog_user = "reader"
$jfrog_pass= "xxxx"
$jfrogurl = "xxxx.jfrog.io/xxxx/binary-snapshot-local/AWSTestResults/$Env:My_APPVEYOR_REPO_COMMIT/TestSummary.xml"
$jfrogurl_nofile = "xxxx.jfrog.io/xxxx/binary-snapshot-local/AWSTestResults/$Env:My_APPVEYOR_REPO_COMMIT/"
$jfrog_URI = New-Object System.Uri("https://"+$jfrogurl) 

$initial_json = @"
{
  "state": "pending",
  "target_url": "https://${jfrog_user}:$jfrog_pass@$jfrogurl_nofile",
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


$hash = [hashtable]::Synchronized(@{})
$hash.DataTable = New-Object System.Data.DataTable "ny"
$hash.Flag = $true
$hash.Host = $host

$runspace = [runspacefactory]::CreateRunspace()
$runspace.Open()
$runspace.SessionStateProxy.SetVariable('hash',$hash)
$powershell = [powershell]::Create()
$powershell.Runspace = $runspace
$powershell.AddScript($MediaServerErrorCheckScript) | Out-Null
$handle_powershellInvoke = $powershell.BeginInvoke()




	$running = $False
	$n_exception = 0
	do {
		Start-Sleep -s 10
		try{
			$state = (Get-ec2instance -Filter @{Name = "network-interface.addresses.private-ip-address"; Values = "$Env:My_AWS_Normal_IP"} -Region "$Env:My_AWS_Region").Instances[0].State
			if ($state.Name -eq "running"){
				$running = $True
				$n_exception = 0
			}
			else{
				$running = $False
				Write-Output "Get-ec2instance state is $state"
			}
			
			if (Test-Path c:\assume_normal_done.txt){
				$running = $False
			}
			
		}catch{
			$n_exception = $n_exception + 1
			Write-Output "Got Exception # $n_exception with message = " + $_.Exception.Message
			if ($n_exception -eq 3){
				$running = $False
				Write-Verbose "3 times I've got exception trying to ask for AWS with ip=$Env:My_AWS_Normal_IP , so I assume something is wrong! aborting...." -Verbose
			}
		}
	} while ($running)

}

	# here you are Normal Media Server!
	Write-Verbose "Adding github statuses..." -Verbose
	
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jgateway) -Uri $github_status_url  -UseBasicParsing | Out-Null
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jsipp) -Uri $github_status_url  -UseBasicParsing | Out-Null
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jMSAutoRun) -Uri $github_status_url -UseBasicParsing | Out-Null
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jAAAutoRun) -Uri $github_status_url -UseBasicParsing | Out-Null
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jMSService) -Uri $github_status_url -UseBasicParsing | Out-Null
	


	$jMSService.state = "success"
	$jMSService.description = "MediaServer Service Start-Stop passed. :heavy_check_mark:"
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jMSService) -Uri $github_status_url -UseBasicParsing | Out-Null
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
	
	try{
		[xml]$userfile = Get-Content C:\projects\test\$Env:MY_APPVEYOR_REPO_BRANCH\TestSummary.xml
		$ReportfilePath = "C:\Projects\test\$Env:MY_APPVEYOR_REPO_BRANCH\Report.xml"
		 
		Invoke-WebRequest  -uri $jfrog_URI -Method Put -InFile "C:\projects\test\$Env:MY_APPVEYOR_REPO_BRANCH\TestSummary.xml" -Credential $credential -ContentType "multipart/form-data" -UseBasicParsing | Out-Null
		 
		$testIgnoreList = New-Object System.Collections.Generic.HashSet[string]
		[regex]$rx="/ignore:(?<ID>.*)"
		$matches = ( (Get-Content C:\projects\media-server.wiki\Continuous-Integration.md) | select-string -Pattern $rx)

		foreach ($match in $matches.Matches){
			$testIgnoreList.Add($match.groups["ID"].Value) | Out-Null
		}

		# Create The Document
		$XmlWriter = New-Object System.XMl.XmlTextWriter($ReportfilePath,$Null)
		 
		# Set The Formatting
		$xmlWriter.Formatting = "Indented"
		$xmlWriter.Indentation = "4"
		 
		# Write the XML Decleration
		$xmlWriter.WriteStartDocument()
		$xmlWriter.WriteStartElement("testsuite") 
		


		foreach( $innerTest in $userfile.SummaryResult.InnerTests.InnerTest) 
		{
		#Add-AppveyorTest -Name $innerTest.TestName -Framework MSTest -Filename $userfile.SummaryResult.TestName  -Outcome $innerTest.TestResult -ErrorMessage $innerTest.ErrorMessage

			# Write the Document

			
			if ($innerTest.TestName.EndsWith("_1") -eq $true ){
				$test_type = "BGW"
			}
			elseif ($innerTest.TestName.EndsWith(".xml") -eq $true){
				$test_type = "Sipp"
			}
			else{
				$test_type = "MSAR"
			}
			
			if ($innerTest.TestName -eq "ScriptEngineClient"){
				$test_type = "BGW"
			}

			if ($innerTest.TestName -eq "AudioAnalyzerUnitTests"){
				$test_type = "AAAR"
			}
			
			if ($innerTest.TestName -eq "AudioAnalyzerAutoRun"){
 				$test_type = "AAAR"
 			}
 			
 			if ($innerTest.TestName -eq "MediaServerScriptedTester"){
 				$test_type = "BGW"
 			}
 			
 			if ($innerTest.TestName -eq "MediaServerSIPpTester"){
 				$test_type = "Sipp"
 			}
 			
 			if ($innerTest.TestName -eq "MediaServerAutoRun"){
 				$test_type = "MSAR"
 			}			
			
			if ($innerTest.TestResult -eq "Passed"){ 
				$testresults[$test_type].passed = $testresults[$test_type].passed + 1
			}
			else{ # ---- FAILED ----
				if ($innerTest.TestName -eq "ScriptEngineClient"){
					if ( ($testresults[$test_type].failed -eq 0) -and ($testresults[$test_type].ignored -eq 0)){ 
						$testresults[$test_type].failed = $testresults[$test_type].failed + 1
					} #dont increment failed ones on BGW because of ScriptEngineClient if there were already failed ones 
				}
				else{
					if ($testIgnoreList.Contains($innerTest.TestName)){
						$testresults[$test_type].ignored = $testresults[$test_type].ignored + 1
					}
					else {
						$testresults[$test_type].failed = $testresults[$test_type].failed + 1
					}
				}
				
				$xmlWriter.WriteStartElement("testcase")
				$xmlWriter.WriteAttributeString("name",$innerTest.TestName)
				$xmlWriter.WriteElementString("error",$innerTest.ErrorMessage)
				$xmlWriter.WriteFullEndElement() # <-- Closing Servers
			}
			#$xmlWriter.WriteElementString("system-err",$innerTest.ErrorMessage)
			
		} #end of for loop
		 
		# Write Close Tag for Root Element
		$xmlWriter.WriteEndElement() # <-- Closing RootElement
		 
		# End the XML Document
		$xmlWriter.WriteEndDocument()
		 
		# Finish The Document
		$xmlWriter.Finalize
		$xmlWriter.Flush
		$xmlWriter.Close()
	}
	catch{
		Write-Output $_.Exception.Message
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
	
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jgateway) -Uri $github_status_url -UseBasicParsing | Out-Null
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jsipp) -Uri $github_status_url  -UseBasicParsing | Out-Null
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jMSAutoRun) -Uri $github_status_url -UseBasicParsing | Out-Null
	Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $jAAAutoRun) -Uri $github_status_url -UseBasicParsing | Out-Null

	



If ($comments_url -ne $null)
{
	$dtEvents = New-Object System.Data.Datatable
	[void]$dtEvents.Columns.Add("Source")
	[void]$dtEvents.Columns.Add("FirstInstance")
	[void]$dtEvents.Columns.Add("Count")
	[void]$dtEvents.Columns.Add("FirstTime")
	
	$AppEvents = Get-EventLog -LogName "Application" -AsBaseObject
	foreach ($AppEvent in $AppEvents){
		if ( ($AppEvent.EntryType -eq "Error") -and ( ($AppEvent.Source -eq ".NET Runtime")) ){
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

	$js_MSErrors = PrintTable $hash.DataTable
	$diskusage = (($diskbefore.Free - $diskafter.Free) / 1024)
	$diskfilesusage = $diskfilesafter - $diskfilesbefore
	
	if ($Env:My_MediaServerMode -eq "Normal"){
		if ($diskusage -ge 450000) {$diskusage_alertcode = "&#x1F621;"} else {$diskusage_alertcode = "&#x1F60A;"}
		if ($diskfilesusage -ge 3000) {$diskfileusage_alertcode = "&#x1F621;"} else {$diskfileusage_alertcode = "&#x1F60A;"}
		
		$js_DiskUsage = [string]::Format("Disk usage was {0:n0}Kb $diskusage_alertcode \n New files created are {1:n0} $diskfileusage_alertcode",$diskusage,$diskfilesusage)

		$Report_XML_Results = (Get-Content $ReportfilePath) | Out-String
		$commet_downloads = "Download report of [Standalone MRC Client](https://${jfrog_user}:$jfrog_pass@$jfrogurl_nofile/Normal/reports.zip) \n Download report of [TTS Media Server](https://${jfrog_user}:$jfrog_pass@$jfrogurl_nofile/TTS/reports.zip) \n "
		$js_comment_xml = "Followings are test cases with errors ( Download [full xml](https://${jfrog_user}:$jfrog_pass@$jfrogurl_nofile) for list of all passed and failed.) : \n ``````xml $Report_XML_Results``````"
		$js_comment_MSErrors = "### Followings are Standalone MRC Client Errors :\n" + $js_MSErrors + "\n\n" + $js_DiskUsage + "\n" 
		$js_comment_table = PrintTable $benchmark_table
		$js_comment_perf = "Performance Index is " + $bechmark_perfIndex_text + "\n\nBenchmark results of Single MRC Client are :\n" + $js_comment_table + "\n"
		$js_comment_MSTTSErrors = "### Followings are TTS Media Server Errors :\n TTS_REPORTS_PENDING..."
		$js_comment_body.body = "# AWS Test Results (" + ($Env:MY_APPVEYOR_PULL_REQUEST_HEAD_COMMIT).SubString(0,7) + ")\n"
		$js_comment_body.body = $js_comment_body.body + "Testing finished at $(Get-Date  -format g)\n\n"
		$js_comment_body.body = $js_comment_body.body + $jMSService.description + " \n " + $jsipp.description + " \n " + $jgateway.description + " \n " + $jMSAutoRun.description + " \n " + $jAAAutoRun.description+" \n\n $commet_downloads \n\n $js_comment_xml\n" + $js_comment_MSErrors + $js_comment_perf + "\n" + $js_comment_MSTTSErrors
		Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $js_comment_body).replace("\\n","\n") -Uri ($comments_url + $github_access_token) -UseBasicParsing | ConvertFrom-Json
	}
	elseif ($found_AWS_Comment -ne $null){
		if ($diskusage -ge 50000) {$diskusage_alertcode = "&#x1F621;"} else {$diskusage_alertcode = "&#x1F60A;"}
		if ($diskfilesusage -ge 100) {$diskfileusage_alertcode = "&#x1F621;"} else {$diskfileusage_alertcode = "&#x1F60A;"}
		
		$js_DiskUsage = [string]::Format("Disk usage was {0:n0}Kb $diskusage_alertcode \n New files created are {1:n0} $diskfileusage_alertcode",$diskusage,$diskfilesusage)
		
		$js_comment_body.body = ($found_AWS_Comment.body).replace("TTS_REPORTS_PENDING...","$js_MSErrors \n\n $js_DiskUsage")
		Invoke-WebRequest  -Method POST -Body (ConvertTo-Json $js_comment_body).replace("\\n","\n") -Uri ($comments_url + $github_access_token) -UseBasicParsing | ConvertFrom-Json
	}
	
}


Stop-Transcript