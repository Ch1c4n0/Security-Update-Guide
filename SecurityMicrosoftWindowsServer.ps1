### Install the module
#Install-Module MSRCSecurityUpdates -Force

### Load the module
Import-Module -Name MsrcSecurityUpdates

# Month Format for MSRC
$Month = Get-Date -Format 'yyyy-MMM'

# Enter the Operating System you specifically want to focus on
$ClientOS_Type = "Windows Server"

# Environment Variables
$Css = "<style>
body {
    font-family: Arial, sans-serif;
    font-size: 10px;
    color: #FFFFFF;
    background: #000000;
}
#title {
    color: #FFFFFF;
    font-size: 30px;
    font-weight: bold;
    height: 50px;
    margin-left: 0px;
    padding-top: 10px;
}
#subtitle {
    font-size: 16px;
    margin-left: 0px;
    padding-bottom: 10px;
    color: #FFFFFF;
}
table {
    width: 100%;
    border-collapse: collapse;
}
table td, table th {
    border: 1px solid #FFFFFF;
    padding: 3px 7px 2px 7px;
}
table th {
    text-align: center;
    padding-top: 5px;
    padding-bottom: 4px;
    background-color: #FFFFFF;
    color: #000000;
}
table tr.alt td {
    color: #000000;
    background-color: #EAF2D3;
}
tr.critical {
    color: white;
    background-color: red;
}
</style>"

$Title = "<span style='font-weight:bold;font-size:24pt'>CVE List for Windows Server " + $Month + "</span>"
$Logo = "<img src='C:\Users\Chicano\OneDrive\MCT\cloudnaquebradav2.png' alt='Logo' height='100' width='100'>"
$Header = "<div id='banner'>$Logo</div>`n" +
          "<div id='title'>$Title</div>`n" +
          "<div id='subtitle'>Report generated: $(Get-Date)</div>"

# Main Script Logic
$ID = Get-MsrcCvrfDocument -ID $Month
$ProductName = Get-MsrcCvrfAffectedSoftware -Vulnerability $ID.Vulnerability -ProductTree $ID.ProductTree | Where-Object { $_.Severity -in 'Critical', 'Important' -and ($_.FullProductName -match $ClientOS_Type) }

$Report = $ProductName | Select CVE, FullProductName, Severity, Impact, @{Name='KBArticle'; Expression={($_.KBArticle.ID | Select-Object -Unique) -join ', '}}, @{Name='BaseScore'; Expression={$_.CvssScoreSet.Base}}, @{Name='TemporalScore'; Expression={$_.CvssScoreSet.Temporal}}, @{Name='Vector'; Expression={$_.CvssScoreSet.Vector}} | ConvertTo-Html -As Table -Fragment | ForEach-Object {
    if ($_ -match "<td.*?Critical.*?>") {
        $_ -replace "<tr>", "<tr class='critical'>"
    } else {
        $_
    }
}

# Combine CSS, Header, and Report into a full HTML document
$HtmlContent = @"
<html>
<head>
    $Css
</head>
<body>
    $Header
    $Report
</body>
</html>
"@

# Save the HTML content to a file
$HtmlFilePath = "C:\Users\Chicano\OneDrive\vscode\PowerShell\WindowsServer.html"
$HtmlContent | Out-File -FilePath $HtmlFilePath -Encoding utf8

# Send the HTML content via email
#$SmtpServer = "smtp.seuservidor.com"
#$From = "seuemail@dominio.com"
#$To = "destinatario@dominio.com"
#$Subject = "Relatório de Segurança Microsoft"
#$Body = $HtmlContent
#$Attachment = $HtmlFilePath

#Send-MailMessage -SmtpServer $SmtpServer -From $From -To $To -Subject $Subject -Body $Body -BodyAsHtml -Attachments $Attachment