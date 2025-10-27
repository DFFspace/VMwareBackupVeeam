#VARIABLE TABLE
$Gebruikers     = "" #Fill in username to authenticate with Veeam
$Wachtwoord     = "" #Fill in password to authenticate with Veeam
$VsphereServer  = "" #Fill in the hostname or IP address from the vSphere server
$VMNamen        = "" #Fill in the virtual machine name
$ClusterA       = "" #Fill in the name of the Cluster
$ClusterB       = "" #Fill in the name of the Cluster
$JobName        = "" #Fill in the name of the Job from Veeam
$BackupServer   = "" #Fill in the hostname of the backup server
$Logfile        = "c:\Temp\Log.txt"
$DatumTijd      = Get-Date
$Smtp           = "" #Fill in SMTP server
$From           = "" #Fill in the from email address
$To             = "" #Fill in the email address where to send the mail to
$Subject        = "Script started at $DatumTijd"
$EmCredU        = "" #Fill in username for email account to send mail with 
$EmCredP        = "" #Fill in password for email account to send mail with 

#LOGGING DOCUMENTATIE
Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring -Encoding UTF8
}

    LogWrite @"
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head>
<meta http-equiv="Content-Type" content="text/html; charset="UTF-8">
<meta name="ProgId" content="Word.Document">
<meta name="Generator" content="Microsoft Word 15">
<meta name="Originator" content="Microsoft Word 15">
<style>
<!--
 /* Font Definitions */
 @font-face
  {font-family:"Cambria Math";
  panose-1:2 4 5 3 5 4 6 3 2 4;
  mso-font-charset:0;
  mso-generic-font-family:roman;
  mso-font-pitch:variable;
  mso-font-signature:3 0 0 0 1 0;}
@font-face
  {font-family:Calibri;
  panose-1:2 15 5 2 2 2 4 3 2 4;
  mso-font-charset:0;
  mso-generic-font-family:swiss;
  mso-font-pitch:variable;
  mso-font-signature:-536858881 -1073732485 9 0 511 0;}
@font-face
  {font-family:Tahoma;
  panose-1:2 11 6 4 3 5 4 4 2 4;
  mso-font-charset:0;
  mso-generic-font-family:swiss;
  mso-font-pitch:variable;
  mso-font-signature:-520081665 -1073717157 41 0 66047 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
  {mso-style-unhide:no;
  mso-style-qformat:yes;
  mso-style-parent:"";
  margin:0cm;
  margin-bottom:.0001pt;
  mso-pagination:widow-orphan;
  font-size:11.0pt;
  font-family:"Calibri",sans-serif;
  mso-fareast-font-family:Calibri;
  mso-fareast-theme-font:minor-latin;}
p.msonormal0, li.msonormal0, div.msonormal0
  {mso-style-name:msonormal;
  mso-style-unhide:no;
  mso-margin-top-alt:auto;
  margin-right:0cm;
  mso-margin-bottom-alt:auto;
  margin-left:0cm;
  mso-pagination:widow-orphan;
  font-size:11.0pt;
  font-family:"Calibri",sans-serif;
  mso-fareast-font-family:Calibri;
  mso-fareast-theme-font:minor-latin;}
span.smalllabel
  {mso-style-name:small_label;
  mso-style-unhide:no;}
.MsoChpDefault
  {mso-style-type:export-only;
  mso-default-props:yes;
  font-size:10.0pt;
  mso-ansi-font-size:10.0pt;
  mso-bidi-font-size:10.0pt;}
@page WordSection1
  {size:612.0pt 792.0pt;
  margin:70.85pt 70.85pt 70.85pt 70.85pt;
  mso-header-margin:35.4pt;
  mso-footer-margin:35.4pt;
  mso-paper-source:0;}
div.WordSection1
  {page:WordSection1;}
-->
</style>

</head>

<body lang="NL" style="tab-interval:35.4pt">

<div class="WordSection1">

<p class="MsoNormal" style="margin-left:120.0pt;text-indent:-120.0pt;tab-stops:
120.0pt;mso-layout-grid-align:none;text-autospace:none"><b><span style="color:black">Van:<span style="mso-tab-count:1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></b><span style="color:black">$From<o:p></o:p></span></p>

<p class="MsoNormal" style="margin-left:120.0pt;text-indent:-120.0pt;tab-stops:
120.0pt;mso-layout-grid-align:none;text-autospace:none"><b><span style="color:black">Verzonden:<span style="mso-tab-count:1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></b><span style="color:black">$DatumTijd<o:p></o:p></span></p>

<p class="MsoNormal" style="margin-left:120.0pt;text-indent:-120.0pt;tab-stops:
120.0pt;mso-layout-grid-align:none;text-autospace:none"><b><span style="color:black">Aan:<span style="mso-tab-count:1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></b><span style="color:black">FILL_IN_A_NAME<o:p></o:p></span></p>

<p class="MsoNormal" style="margin-left:120.0pt;text-indent:-120.0pt;tab-stops:
120.0pt;mso-layout-grid-align:none;text-autospace:none"><b><span style="color:black">Onderwerp:<span style="mso-tab-count:1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></b><span style="color:black">[Success] Veeam Backup script is succesvol uitgevoerd <o:p></o:p></span></p>

<p class="MsoNormal"><o:p>&nbsp;</o:p></p>

<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%;border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:
 0cm 0cm 0cm 0cm">
 <tbody><tr style="mso-yfti-irow:0;mso-yfti-firstrow:yes">
  <td style="padding:0cm 0cm 0cm 0cm">
  <table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%;border-collapse:collapse;mso-yfti-tbllook:1184;
   mso-padding-alt:0cm 0cm 0cm 0cm">
   <tbody><tr style="mso-yfti-irow:0;mso-yfti-firstrow:yes;height:52.5pt">
    <td width="80%" valign="bottom" style="width:80.0%;background:#00B050;
    padding:0cm 0cm 12.75pt 11.25pt;height:52.5pt">
    <p class="MsoNormal"><b><span style="font-size:12.0pt;font-family:&quot;Tahoma&quot;,sans-serif;
    mso-fareast-font-family:&quot;Times New Roman&quot;;color:white">Backup job: $JobName <o:p></o:p></span></b></p>
    <div style="margin-top:3.75pt">
    <p class="MsoNormal"><b><span style="font-size:9.0pt;font-family:&quot;Tahoma&quot;,sans-serif;
    mso-fareast-font-family:&quot;Times New Roman&quot;;color:white">Gemaakt door
    Pasquale D. op 31-12-2018 13:37.<o:p></o:p></span></b></p>
    </div>
    </td>
    <td valign="bottom" style="background:#00B050;padding:0cm 0cm 12.75pt 11.25pt;
    height:52.5pt">
    <p class="MsoNormal"><b><span style="font-size:12.0pt;font-family:&quot;Tahoma&quot;,sans-serif;
    mso-fareast-font-family:&quot;Times New Roman&quot;;color:white">Success<o:p></o:p></span></b></p>
    <div style="margin-top:3.75pt">
    <p class="MsoNormal"><b><span style="font-size:9.0pt;font-family:&quot;Tahoma&quot;,sans-serif;
    mso-fareast-font-family:&quot;Times New Roman&quot;;color:white">Veeam_Backup script
    <o:p></o:p></span></b></p>
    </div>
    </td>
   </tr>

"@

#DATUM + TIJD
    LogWrite @"
    <tr style="mso-yfti-irow:1;mso-yfti-lastrow:yes">
    <td colspan="2" style="padding:0cm 0cm 0cm 0cm">
    <table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%;border-collapse:collapse;mso-yfti-tbllook:
     1184;mso-padding-alt:0cm 0cm 0cm 0cm">
     <tbody><tr style="mso-yfti-irow:0;mso-yfti-firstrow:yes;height:26.25pt">
      <td colspan="9" style="border:solid #A7A9AC 1.0pt;mso-border-alt:solid #A7A9AC .75pt;
      mso-border-bottom-alt:solid #A7A9AC .25pt;background:#F3F4F4;padding:
      3.75pt 0cm 0cm 11.25pt;height:26.25pt">
      <p class="MsoNormal"><span style="font-size:12.0pt;font-family:&quot;Tahoma&quot;,sans-serif;
      mso-fareast-font-family:&quot;Times New Roman&quot;;color:#626365"><b>📜 Script started at $DatumTijd 📜</b><o:p></o:p></span></p>
      </td>
     </tr>
"@

LogWrite @"
      <td width="1%" nowrap="" valign="top" style="width:1.0%;border:solid #A7A9AC 1.0pt;
      border-top:none;mso-border-top-alt:solid #A7A9AC .75pt;mso-border-alt:
      solid #A7A9AC .75pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:12.75pt">
      <p class="MsoNormal"><span style="font-size:9.0pt;font-family:&quot;Tahoma&quot;,sans-serif;
      mso-fareast-font-family:&quot;Times New Roman&quot;">
"@

#VERPLAATS VMS NAAR ClusterA
    LogWrite "<b>[💡 Moving VMs to $ClusterA 💡]</b><br>"
    LogWrite "<b>1.</b> Trying to connect to Vsphere Center Server<br>"
if (connect-viserver $Vsphereserver -User $Gebruikers -Password $Wachtwoord){
    LogWrite "<b>2.</b> Successfully connected to Vsphere Center Server ✔️<br>"
}else{
    LogWrite "<b>2.</b> Failed connecting to Vsphere Center Server ❌<br>"
    Exit
}
    LogWrite "<b>3.</b> Grabbing the VMs and move them to $ClusterA<br>"
if (Get-VM -Name $VMnamen | Move-VM -Destination $ClusterA){
    LogWrite "<b>4.</b> Successfully moved all VMs to $ClusterA ✔️<br>"
}else{
    LogWrite "<b>4.</b> Failed moving all VMs to $ClusterA ❌<br>"
    Exit
}


#UITVOERING VAN VEEAM BACK-UP JOB
    LogWrite "<br>"
    LogWrite "<b>[💡 Veeam Center Back-up 💡]</b><br>"
    LogWrite "<b>5.</b> Trying to connect to Veeam Back-up Center<br>"
Add-PSSnapin -Name VeeamPSSnapIn -ErrorAction SilentlyContinue
Disconnect-VBRServer | out-null
connect-vbrserver -server $BackupServer -user $Gebruikers -password $Wachtwoord
    LogWrite "<b>6.</b> Successfully connected to Veeam Back-up Center ✔️<br>"
$JobObject      = Get-VBRJob | where {$_.Name -eq $JobName}    
    LogWrite "<b>7.</b> Trying to run $JobName Job<br>"
if (Start-VBRJob $JobObject){
    LogWrite "<b>8.</b> Successfully ran $JobName Job ✔️<br>"
}else{
    LogWrite "<b>8.</b> Failed running $JobName Job ❌<br>"
    Exit
}    


#VERPLAATS VMS WEER TERUG NAAR ClusterB
    LogWrite "<br>"
    LogWrite "<b>[💡 Moving VMs back to $ClusterB 💡]</b><br>"
    LogWrite "<b>9.</b> Trying to connect to Vsphere Center Server<br>"
if (connect-viserver $Vsphereserver -User $Gebruikers -Password $Wachtwoord){
    LogWrite "<b>10.</b> Successfully connected to Vsphere Center Server ✔️<br>"
}else{
    LogWrite "<b>10.</b> Failed connecting to Vsphere Center Server ❌<br>"
    Exit
}
    LogWrite "<b>11.</b> Grabbing the VMs and move them back to $ClusterB<br>"
if (Get-VM -Name $VMnamen | Move-VM -Destination $ClusterB){
    LogWrite "<b>12.</b> Successfully moved all VMs back to $ClusterB ✔️<br>"
}else{
    LogWrite "<b>12.</b> Failed to moving all VMs back to $ClusterB ❌<br>"
    Exit
}

$DatumTijd      = Get-Date
    LogWrite @"
</span></b><span style="font-size:9.0pt;font-family:&quot;Tahoma&quot;,sans-serif;mso-fareast-font-family:
      &quot;Times New Roman&quot;"><o:p></o:p></span></p>
      </td>
<tr style="mso-yfti-irow:4;height:26.25pt">
      <td nowrap="" colspan="9" style="border:solid #A7A9AC 1.0pt;border-top:none;
      mso-border-top-alt:solid #A7A9AC .75pt;mso-border-alt:solid #A7A9AC .75pt;
      background:#F3F4F4;padding:3.75pt 0cm 0cm 11.25pt;height:26.25pt">
      <p class="MsoNormal"><span style="font-size:12.0pt;font-family:&quot;Tahoma&quot;,sans-serif;
      mso-fareast-font-family:&quot;Times New Roman&quot;;color:#626365"><b>📜 Script ended at $DatumTijd 📜</b><o:p></o:p></span></p>
      </td>
     </tr> 
    </tbody></table>
    </td>
   </tr>
  </tbody></table>
  </td>
 </tr>
 <tr style="mso-yfti-irow:1">
  <td style="padding:0cm 0cm 0cm 0cm">
  <p class="MsoNormal"><span style="mso-fareast-font-family:&quot;Times New Roman&quot;">&nbsp;<o:p></o:p></span></p>
  </td>
 </tr>
 
</tbody></table>

<p class="MsoNormal"><span style="mso-fareast-font-family:&quot;Times New Roman&quot;"><o:p>&nbsp;</o:p></span></p>

</div>




</body></html>

"@

#E-MAIL SETTINGS
$Body = (cat c:\temp\Log.txt|out-string)
send-MailMessage -SmtpServer $Smtp -Credential $EmCredU -To $To -From $From -Subject $Subject -Body $Body -Encoding UTF8 -BodyAsHtml 

#VERWIJDERING LOGFILE
Remove-Item $Logfile
