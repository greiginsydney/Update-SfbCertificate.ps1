# Update-SfbCertificate.ps1
Lync &amp; Skype for Business need their PKI certificates refreshed occasionally, and there's always the risk you'll misspell or accidentally omit a vital SAN. No longer. This script will request a new cert (internal or external) using the existing cert as a template.

Skype for Business &ndash; like so much else these days &ndash; relies on PKI certificates, and the community has risen to the opportunity with some great tools to help us manage them.
Here are two in my &ldquo;essentials&rdquo; kit:
<ol>
- Check out the &ldquo;<a href="https://gallery.technet.microsoft.com/Lync-Certificates-Reporter-502fefaf" target="_blank">Lync Certificates Report</a>&ldquo;: If you&rsquo;ve not found it already, Guy Bachar &amp; Yoav Barzilay (with some input from Anthony  Caragol) have crafted a fantastic script that will read your Lync/SfB topology and query all the servers (using WinRM) then prepare a report showing how long each of the cert&rsquo;s still has to live. 
- If your new OAuth cert isn&rsquo;t replicating or your Front-End service won&rsquo;t start, it might be due to bad cert placement. David Paulino&rsquo;s &ldquo;<a href="https://gallery.technet.microsoft.com/LyncSkype4B-Certificate-81944851" target="_blank">Test-CertificateStore.ps1</a>&rdquo;  to the rescue! 
</ol>
I recently found myself needing to update ALL of the internal certs for a couple of deployments &ndash; one of them quite large &ndash; and I spent many hours checking and re-checking to make sure I&rsquo;d not misspelt or overlooked a SAN in all of the  requests.
Yes, the Deployment Wizard and PowerShell are <em>meant </em>to include all of the expected SANs in the automated request, but your existing certificate might contain others that <em>aren&rsquo;t</em> added automatically:
<ul>
- a site-specific simple URL (<a href="http://www.justin-morris.net/configuring-site-level-simple-urls-in-lync-server-2010/" target="_blank">thank you Justin for the walk-through</a>) 
- the scheduler (<a href="http://blog.mcgreanor.com/2012/11/14/lync-2013-web-scheduler/" target="_blank">nod to Chad</a>) 
- ucupdates-r2 
- or all of the Front-End FQDNs so the one cert can be deployed across the entire EE pool. 
</ul>
&nbsp;
PowerShell to the rescue!
&nbsp;

## What is Update-SfBCertificate.ps1?
&ldquo;Update-SfBCertificate.ps1&rdquo; (herein just &ldquo;Update&rdquo;) basically just acts as a front-end for the &ldquo;Request-CsCertificate&rdquo; commandlet &ndash; but it starts the request process by first reading in the details of the certificate  you want to replace or renew.
The &ldquo;Request-CsCertificate&rdquo; (&ldquo;Request&rdquo;) commandlet requires a lot of parameters to create a complete certificate request, and yet for quite a lot of them the values you probably want to use are already in your existing certificate:  Organisation, OU, City, State, Country being the obvious ones. So the existing values become the template from which this script will request a replacement. If you want to override or replace any of the existing values, any you provide from the command-line  will be used instead of those coming from the existing certificate.
&ldquo;Update&rdquo; also reads the existing SANs from the certificate and adds them to the &ldquo;Domains&rdquo; parameter that will be fed into &ldquo;Request&rdquo; &ndash; but you&rsquo;re free to inject any others you want too just by naming them in  a comma-separated list, as you would if you were running the &ldquo;Request&rdquo; cmdlet directly. (Don&rsquo;t worry about any unwanted repetition &ndash; I de-dupe the list).
I&rsquo;ve forced the &ldquo;-confirm&rdquo; switch so you&rsquo;ll be presented with the values that are about to be fed to Request &amp; have a chance to abort and revise the values. (I even present a couple of warnings if you might be about to proceed  with some unintended values).
Upon successfully requesting a new certificate from your online CA it will automatically installed to the server, and then an on-screen comparison will show you where the two differ. Expected differences (like the &ldquo;not before&rdquo;, &ldquo;not after&rdquo;,  &ldquo;serial number&rdquo; and &ldquo;thumbprint&rdquo; values) will be displayed as warnings, whilst any unexpected differences (like changes to SANs) will be represented as Errors. Any changes the user specified from the command-line will also be shown  as warnings to provide confirmation they were enacted.
If you like what you see, the commandlet to Assign this new certificate will be found on the clipboard, ready to paste and enjoy. If the cert&rsquo;s aren&rsquo;t aligned correctly, you&rsquo;re free to repeat the process, adjusting the parameters until  the values are agreeable. As always, please keep a note of any unsuitable certs that may be issued so you can delete them from the server and revoke them from the CA.

## Features
<ul>
- Choose your starting position: Start with the currently-active cert for a given role (e.g. Default), or specify the thumbnail of a particular cert 
- No need to neaten the thumbprint if you&rsquo;re copying from the certs mmc: just scrape it from there and paste it in, spaces, warts and all (but don&rsquo;t forget to put it inside quotation marks if it has spaces) 
- Automatically ensures all the SANs on the existing cert are copied to the new one: no more omissions or typos 
- Lets you consolidate separate Default, InternalWebServices &amp; ExternalWebServices front-end certs into one new Combined cert, with the SANs from all three 
- Include the &ldquo;-AllPoolMemberServers&rdquo; switch and the script will add the FQDNs of all FE&rsquo;s in the pool as SANs 
- Copies all of the mundane Org, OU, City, State, Country values across for you 
- Automatically reads the server&rsquo;s FQDN and feeds it to the request commandlet as the &ldquo;ComputerFqdn&rdquo; parameter. (Specify &ldquo;-ComputerFqdn $null&rdquo; to disable, say for an Edge server) 
- Pops a warning if you&rsquo;re about to renew a cert with a 1024 bit key. (Friends don&rsquo;t let friends&hellip;) 
- The more the merrier: don&rsquo;t be concerned for the same SAN appearing multiple times in your new certificate &ndash; the script de-dupes the list 
- Works in online and offline mode (although in offline mode you don&rsquo;t get the &ldquo;before and after&rdquo; comparison). The work-around here is to take the offline request to completion then install the new certificate and feed both thumbprints to  my free-standing &ldquo;<a href="https://gallery.technet.microsoft.com/Compare-two-identical-PKI-6dcbfdec" target="_blank">Compare-PkiCertificates.ps1</a>&rdquo; (which is largely the same &ldquo;before and after&rdquo; engine) 
- Shows the existing and new certificates side-by-side so you can quickly see what&rsquo;s unchanged or different between them before you Assign the new one into service 
- If the new certificate meets your requirements, you&rsquo;ll find the commandlet to Assign it is on your Windows clipboard after the script exits! 
- Uses your existing PowerShell colour scheme to represent the identical, expected and unexpected differences 
- Even though long values are truncated on screen, they&rsquo;re compared at their full length before being truncated 
- Breaks out all of the SANs to a line each, clearly showing any that are added or removed 
- Code-signed so it&rsquo;ll run in restricted environments. (Thank you Digicert) 
- Auto-adjusts column widths to make maximum use of your screen width 
</ul>

## Shortcomings, Weaknesses
<ul>
- If you&rsquo;re requesting a new certificate on behalf of another server (as you might normally by using the &ldquo;ComputerFqdn&rdquo; parameter) you&rsquo;ll need to first install its &lsquo;starter&rsquo; cert on this server so it can be read when you  add the &ldquo;-Thumbnail&rdquo; parameter 
- If you are consolidating multiple Front-End certificates into one replacement cert, the existing &lsquo;location&rsquo; values (City, State, etc) are only captured from the first certificate. SANs are however read from all the certs you nominate (e.g. Default,WebServicesInternal,WebServicesExternal)  &amp; de-duped 
- You can&rsquo;t remove SANs that are no longer required 
- You can only request an Offline certificate when running this script on an Edge server (and you&rsquo;ll need to specify &ldquo;-ComputerFqdn $null&rdquo; to stop it attempting to read the Topology). Another approach is to copy the Edge&rsquo;s cert to  an FE and do it from there! 
- Outputs to screen using &ldquo;write-host&rdquo; and not to the pipeline 
- PowerShell v2 (Server 2008R2 &amp; Windows 7) doesn&rsquo;t reveal all of the values to me. Rather than prevent the script from running under v2, I pop a warning 
- Needs to be run as Admin to reliably show all information 
</ul>

## How-To
Here are some examples of how to use it:
This is your minimum to refresh an existing FE cert. This makes no changes to the cert&rsquo;s values (other than the new Friendly name of course):
```
powershell
PS C:\&gt; .\Update-SfBCertificate.ps1 -type default,webservicesinternal,webservicesexternal -FriendlyName "Combined SYD Pool FE cert 27Apr2017" -ca MyCA.contoso.com\My-CA
```
Refresh your existing FE cert. Add all the pool FQDNs and make sure the SIP domains are all there too:
```powershell
PS C:\&gt; .\Update-SfBCertificate.ps1 -type default,webservicesinternal,webservicesexternal -Verbose -FriendlyName "Combined SYD Pool FE cert 27Apr2017" -ca MyCA.contoso.com\My-CA -AllPoolMemberServers -AllSipDomain
```
Has someone recently updated the server&rsquo;s cert and overlooked a SAN? We&rsquo;ve all been there. Use the thumbprint of the old one and inject the SAN:
```powershell
PS C:\&gt; .\Update-SfBCertificate.ps1 -type default,webservicesinternal,webservicesexternal -Verbose -FriendlyName "Combined SYD Pool FE cert 27Apr2017" -ca MyCA.contoso.com\My-CA -Thumbprint "?12 34 56 78 90 12 34 45 78 89" -Domain "OverlookedSAN.contoso.com"
```

## Show me
Here I&rsquo;m requesting a new default,webservicesinternal,webservicesexternal certificate for my Front-End. I&rsquo;m providing the details of the CA, specifying a new Friendly Name, making sure the City comes out as &ldquo;Sydney&rdquo;, adding all SIP  Domains, all FE&rsquo;s in the pool, and adding SANs for meet.contoso &amp; meet.fabrikam:

<img src="https://user-images.githubusercontent.com/11004787/81052514-e606c300-8f06-11ea-8832-66672dfebe29.png" alt="" width="600" />

The script pauses at this point. On-screen (to the left of the green bar here) are all the values that will be fed into &ldquo;Request-CsCertificate&rdquo; if you agree to the Confirm prompt at the bottom.

<img src="https://user-images.githubusercontent.com/11004787/81052556-ffa80a80-8f06-11ea-9235-0dd42098206e.png" alt="" width="600" />

And here's the comparison of old/existing and new. The expected differences between them are highlighted in yellow. Any that aren't expected are shown in red (or whatever your Warning and Error colour preferences are set to, respectively).
If you like the look of the new certificate, right-click to paste the correctly-formatted "Assign-CsCertificate" command to make it active.
&nbsp;

## Script Revision History

### v1.3: 12th May 2018
<ul>
- Added an abort line that kills the script when running in the (unsupported) PowerShell ISE. (Screen-width and coloured output don't work) 
</ul>

### v1.2: 24th December 2017
<ul>
- Improved the way the "Subject" is parsed in ParseCertSubject by trimming leading spaces 
- Added "E=" for those certs where an e-mail address has been provided. (Not applicable to new cert requests via the "Request-CsCertificate") 
- Changed the cert comparison highlighting: no longer shows in 'warning' colour if the user provided a 'new' value but nothing changed in the resulting cert 
- Incorporated my version of Pat's "Get-UpdateInfo". Credit: https://ucunleashed.com/3168 
</ul>

### v1.1: 19th February 2017 - the "Thank You Mike Shivtorov" bugfix &amp; suggestions release
<ul>
- Added 'XmppServer' certificate type, overlooked in the original release 
- Changed Output request file example text from ".pfx" to ".req" 
- Added another example to better document how to generate Edge certificates 
- Improved offline request process: now finds and re-opens the offline request, then feeds it to the comparison engine for display 
- Added an extra Exception trap to the Request-CsCertificate handling: script now reports cleanly if you request an inappropriate Type 
- Sorted Key Usages before sending them to the Compare engine in an effort to reduce false positives 
- Fixed bug where KeySize was reported in red instead of yellow when the user provided a new value 
</ul>

### v1.0: 7th May 2016. Initial public release
<ul>
- Improved the error reporting when Request-CsCertificate fails 
</ul>
<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/update-sfbcertificate-ps1/](https://greiginsydney.com/update-sfbcertificate-ps1/).

