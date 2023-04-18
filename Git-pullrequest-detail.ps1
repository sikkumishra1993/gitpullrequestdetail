Function sendPullRequestDetails
{
   param (
       [Parameter(Mandatory=$True)][string]$SmtpServerparam,
       [Parameter(Mandatory=$True)][string]$SMTPPortparam,
       [Parameter(Mandatory=$True)][string]$EmailFromparam,
       [Parameter(Mandatory=$True)][string]$EmailToparam,
       [Parameter(Mandatory=$True)][string]$EmailPassword,
       [Parameter(Mandatory=$True)][string]$gitRepoURL

   )

   process
   {

    try
        {

        $allpr = Invoke-RestMethod -Method Get -uri $gitRepoURL

        #Get Date

        $FromDate = (Get-Date).AddDays(-7).ToString("yyyy-MM-dd")
        $ToDate = Get-Date -format "yyyy-MM-dd"

        #Configuration Variables for E-mail credential and function call

        [string][ValidateNotNullOrEmpty()] $pass = $EmailPassword
        $userPassword = ConvertTo-SecureString -String $pass -AsPlainText -Force
        [string][ValidateNotNullOrEmpty()] $name = $EmailFromparam
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $name,$userPassword

        $SmtpServer = $SmtpServerparam
        $SMTPPort = $SMTPPortparam
        $EmailFrom = $EmailFromparam
        $EmailTo = $EmailToparam 

        $EmailSubject = "Pull request details from "+$FromDate+" to "+$ToDate

#HTML Template

$EmailBody = @"
<html>
<body>

<p>Dear Shashi,</p>        
<p>Please find below the detail of Open, Closed and in-progress pull request for last 7 days.</p>


<table style="width: 68%" style="text-align: center; border-collapse: collapse; border: 1px solid #008080;">
    <tr>
    <td colspan="4" bgcolor="#008080" style=" text-align: center;color: #FFFFFF; font-size: large; height: 35px;">
        Closed pull request - Daily Report on: VarReportDate 
    </td>
    </tr>
    <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 10px">
    <td style="text-align: center; width: 201px; height: 39px">  <b> Pull Request Title</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> ID</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> State</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> Closed-on</b></td>
    </tr>
"@

for (($i = 6); $i -ge 0; $i--)

{

$cdate = (Get-Date).AddDays(-$i).ToString("yyyy-MM-dd")

$closedthisweek = $allpr | where-Object closed_at -match $cdate
$prcountclosed += ($closedthisweek.id).count

foreach ($dt in $closedthisweek){

$EmailBody += @" 
    <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
    <td style="text-align: center; height: 35px; width: 233px;"> VarTitle</td>
    <td style="text-align: center; height: 35px; width: 233px;"> VarID</td>
    <td style="text-align: center; height: 35px; width: 233px;"> VarState</td>
    <td style="text-align: center; height: 39px; width: 233px;"> VarClosedAt</td>
    </tr>
"@

$EmailBody= $EmailBody.Replace("VarTitle",$dt.title)
$EmailBody= $EmailBody.Replace("VarID",$dt.id)
$EmailBody= $EmailBody.Replace("VarState",$dt.state)
$EmailBody= $EmailBody.Replace("VarClosedAt",($dt.closed_at -split "T")[0])

}
}

$EmailBody += @"
    <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
    <td style="text-align: center; height: 35px; width: 233px;"> <b> Total Pul Requests</b></td>
    <td style="text-align: center; height: 39px; width: 233px;"> Vartotalpr</td>
    </tr>
</table>
"@

$EmailBody= $EmailBody.Replace("Vartotalpr",$prcountclosed)

$EmailBody += @"
<table style="width: 68%" style="border-collapse: collapse; border: 1px solid #008080;">
    <tr>
    <td colspan="4" bgcolor="#008080" style="text-align: center; color: #FFFFFF; font-size: large; height: 35px;">
        Open pull request - Daily Report on: VarReportDate 
    </td>
    </tr>
    <tr style="text-align: center; border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 10px">
    <td style="text-align: center; width: 201px; height: 39px">  <b> Pull Request Title</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> ID</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> State</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> Opened-on</b></td>
    </tr>
"@

for (($i = 6); $i -ge 0; $i--)

{

$cdate = (Get-Date).AddDays(-$i).ToString("yyyy-MM-dd")

$createdthisweek = $allpr | where-Object {($_.state -match "open") -and ($_.created_at -match $cdate)}
$prcountcreated += ($createdthisweek.id).count

foreach ($dt in $createdthisweek){

$EmailBody += @" 
    <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
    <td style="text-align: center; height: 35px; width: 233px;"> VarTitle</td>
    <td style="text-align: center; height: 35px; width: 233px;"> VarID</td>
    <td style="text-align: center; height: 35px; width: 233px;"> VarState</td>
    <td style="text-align: center; height: 39px; width: 233px;"> VarClosedAt</td>
    </tr>
"@

$EmailBody= $EmailBody.Replace("VarTitle",$dt.title)
$EmailBody= $EmailBody.Replace("VarID",$dt.id)
$EmailBody= $EmailBody.Replace("VarState",$dt.state)
$EmailBody= $EmailBody.Replace("VarClosedAt",($dt.created_at -split "T")[0])

}
}

$EmailBody += @"

    <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
    <td style="text-align: center; height: 35px; width: 233px;"> <b> Total Pul Requests</b></td>
    <td style="text-align: center; height: 39px; width: 233px;"> Vartotalpr</td>
    </tr>
</table>
"@

$EmailBody= $EmailBody.Replace("Vartotalpr",$prcountcreated)

$EmailBody += @"
<table style="width: 68%" style="border-collapse: collapse; border: 1px solid #008080;">
    <tr>
    <td colspan="4" bgcolor="#008080" style="text-align: center; color: #FFFFFF; font-size: large; height: 35px;">
        In-Progress pull request - Daily Report on: VarReportDate 
    </td>
    </tr>
    <tr style="text-align: center; border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 10px">
    <td style="text-align: center; width: 201px; height: 39px">  <b> Pull Request Title</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> ID</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> State</b> </td>
    <td style="text-align: center; width: 201px; height: 39px">  <b> Updated-on</b></td>
    </tr>
"@

for (($i = 6); $i -ge 0; $i--)

{

$cdate = (Get-Date).AddDays(-$i).ToString("yyyy-MM-dd")

$inprogress = $allpr | where-Object {($_.state -match "open") -and ($_.updated_at -match $cdate)}
$prcountinprogress += ($inprogress.id).count

foreach ($dt in $inprogress){

$EmailBody += @" 
    <tr style="text-align: center; border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
    <td style="text-align: center; height: 35px; width: 233px;"> VarTitle</td>
    <td style="text-align: center; height: 35px; width: 233px;"> VarID</td>
    <td style="text-align: center; height: 35px; width: 233px;"> VarState</td>
    <td style="text-align: center; height: 39px; width: 233px;"> VarClosedAt</td>
    </tr>
"@

$EmailBody= $EmailBody.Replace("VarTitle",$dt.title)
$EmailBody= $EmailBody.Replace("VarID",$dt.id)
$EmailBody= $EmailBody.Replace("VarState",$dt.state)
$EmailBody= $EmailBody.Replace("VarClosedAt",($dt.updated_at -split "T")[0])

}
}

$EmailBody += @"
    <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
    <td style="text-align: center; height: 35px; width: 233px;"> <b> Total Pul Requests</b></td>
    <td style="text-align: center; height: 39px; width: 233px;"> Vartotalpr</td>
    </tr>
</table>
<p>Thanks & Regards,</p>
<p>Shashi Bhushan Mishra</p>
</body>
</html>
"@

        $EmailBody= $EmailBody.Replace("Vartotalpr",$prcountinprogress)
        $EmailBody= $EmailBody.Replace("VarReportDate",$ToDate)
  
        #Send E-mail from PowerShell script
        Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -Body $EmailBody -BodyAsHtml -SmtpServer $SmtpServer -port $SMTPPort -Credential $cred -UseSsl
        echo "Email is sent to : $EmailTo"
        echo "Email is sent by : $EmailFrom"
        echo "Subject of email : $EmailSubject"
        }

        catch
           {
               Write-Error "[ERROR] Failed to send the mail"
               throw $_.Exception.Message
           }
        
    }
}

#**********************************************#
#************Function body*********************#
#********Function execution begins here********#
#**********************************************#

try
{

#declairing variable to set the git repo url
$giturl = "https://api.github.com/repos/actionsdemos/calculator/pulls?state=all&per_page=100"

#declairing variable to set the credential to authenticate with SMTP server, in this case I have used gmail app credential which is stored in windows secret store
$secret = Get-Secret -Name gmailsecret -AsPlainText SecretStore

#declairing variables to be passed to function call
$SmtpServer = "smtp.gmail.com"
$SMTPPort = "587"
$EmailFrom = "sikkumishra1994@gmail.com"
$EmailTo =  "sikkumishra1993@gmail.com"

sendPullRequestDetails -SmtpServerparam $SmtpServer -SMTPPortparam $SMTPPort -EmailFromparam $EmailFrom -EmailToparam $EmailTo -EmailPassword $secret -gitRepoURL $giturl

}

catch
       {
           Write-Error "[ERROR] Failed to send the mail"
           throw $_.Exception.Message
       }
