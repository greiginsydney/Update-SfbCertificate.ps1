# Update-SfbCertificate.ps1
Lync & Skype for Business need their PKI certificates refreshed occasionally, and there's always the risk you'll misspell or accidentally omit a vital SAN. No longer. This script will request a new cert (internal or external) using the existing cert as a template.

Skype for Business &ndash; like so much else these days &ndash; relies on PKI certificates, and the community has risen to the opportunity with some great tools to help us manage them.

Here are two in my "essentials" kit:

- Check out the "<a href="https://gallery.technet.microsoft.com/Lync-Certificates-Reporter-502fefaf" target="_blank">Lync Certificates Report</a>": If you've not found it already, Guy Bachar & Yoav Barzilay (with some input from Anthony  Caragol) have crafted a fantastic script that will read your Lync/SfB topology and query all the servers (using WinRM) then prepare a report showing how long each of the cert's still has to live.

- If your new OAuth cert isn't replicating or your Front-End service won't start, it might be due to bad cert placement. David Paulino's "<a href="https://gallery.technet.microsoft.com/LyncSkype4B-Certificate-81944851" target="_blank">Test-CertificateStore.ps1</a>"  to the rescue!

I recently found myself needing to update ALL of the internal certs for a couple of deployments &ndash; one of them quite large &ndash; and I spent many hours checking and re-checking to make sure I'd not misspelt or overlooked a SAN in all of the  requests.

Yes, the Deployment Wizard and PowerShell are <em>meant </em>to include all of the expected SANs in the automated request, but your existing certificate might contain others that <em>aren't</em> added automatically:

- a site-specific simple URL (<a href="http://www.justin-morris.net/configuring-site-level-simple-urls-in-lync-server-2010/" target="_blank">thank you Justin for the walk-through</a>)

- the scheduler (<a href="http://blog.mcgreanor.com/2012/11/14/lync-2013-web-scheduler/" target="_blank">nod to Chad</a>)

- ucupdates-r2

- or all of the Front-End FQDNs so the one cert can be deployed across the entire EE pool.

PowerShell to the rescue!


## What is Update-SfBCertificate.ps1?
"Update-SfBCertificate.ps1" (herein just "Update") basically just acts as a front-end for the "Request-CsCertificate" commandlet &ndash; but it starts the request process by first reading in the details of the certificate  you want to replace or renew.

The "Request-CsCertificate" ("Request") commandlet requires a lot of parameters to create a complete certificate request, and yet for quite a lot of them the values you probably want to use are already in your existing certificate:  Organisation, OU, City, State, Country being the obvious ones. So the existing values become the template from which this script will request a replacement. If you want to override or replace any of the existing values, any you provide from the command-line  will be used instead of those coming from the existing certificate.

"Update" also reads the existing SANs from the certificate and adds them to the "Domains" parameter that will be fed into "Request" &ndash; but you're free to inject any others you want too just by naming them in  a comma-separated list, as you would if you were running the "Request" cmdlet directly. (Don't worry about any unwanted repetition &ndash; I de-dupe the list).

I've forced the "-confirm" switch so you'll be presented with the values that are about to be fed to Request & have a chance to abort and revise the values. (I even present a couple of warnings if you might be about to proceed  with some unintended values).

Upon successfully requesting a new certificate from your online CA it will automatically installed to the server, and then an on-screen comparison will show you where the two differ. Expected differences (like the "not before", "not after",  "serial number" and "thumbprint" values) will be displayed as warnings, whilst any unexpected differences (like changes to SANs) will be represented as Errors. Any changes the user specified from the command-line will also be shown  as warnings to provide confirmation they were enacted.

If you like what you see, the commandlet to Assign this new certificate will be found on the clipboard, ready to paste and enjoy. If the cert's aren't aligned correctly, you're free to repeat the process, adjusting the parameters until  the values are agreeable. As always, please keep a note of any unsuitable certs that may be issued so you can delete them from the server and revoke them from the CA.

## Features

- Choose your starting position: Start with the currently-active cert for a given role (e.g. Default), or specify the thumbnail of a particular cert

- No need to neaten the thumbprint if you're copying from the certs mmc: just scrape it from there and paste it in, spaces, warts and all (but don't forget to put it inside quotation marks if it has spaces)

- Automatically ensures all the SANs on the existing cert are copied to the new one: no more omissions or typos

- Lets you consolidate separate Default, InternalWebServices & ExternalWebServices front-end certs into one new Combined cert, with the SANs from all three

- Include the "-AllPoolMemberServers" switch and the script will add the FQDNs of all FE's in the pool as SANs

- Copies all of the mundane Org, OU, City, State, Country values across for you

- Automatically reads the server's FQDN and feeds it to the request commandlet as the "ComputerFqdn" parameter. (Specify "-ComputerFqdn $null" to disable, say for an Edge server)

- Pops a warning if you're about to renew a cert with a 1024 bit key. (Friends don't let friends&hellip;)

- The more the merrier: don't be concerned for the same SAN appearing multiple times in your new certificate &ndash; the script de-dupes the list

- Works in online and offline mode (although in offline mode you don't get the "before and after" comparison). The work-around here is to take the offline request to completion then install the new certificate and feed both thumbprints to  my free-standing "<a href="https://gallery.technet.microsoft.com/Compare-two-identical-PKI-6dcbfdec" target="_blank">Compare-PkiCertificates.ps1</a>" (which is largely the same "before and after" engine)

- Shows the existing and new certificates side-by-side so you can quickly see what's unchanged or different between them before you Assign the new one into service

- If the new certificate meets your requirements, you'll find the commandlet to Assign it is on your Windows clipboard after the script exits!

- Uses your existing PowerShell colour scheme to represent the identical, expected and unexpected differences

- Even though long values are truncated on screen, they're compared at their full length before being truncated

- Breaks out all of the SANs to a line each, clearly showing any that are added or removed

- Code-signed so it'll run in restricted environments. (Thank you Digicert)

- Auto-adjusts column widths to make maximum use of your screen width

## Shortcomings, Weaknesses

- If you're requesting a new certificate on behalf of another server (as you might normally by using the "ComputerFqdn" parameter) you'll need to first install its &lsquo;starter' cert on this server so it can be read when you  add the "-Thumbnail" parameter

- If you are consolidating multiple Front-End certificates into one replacement cert, the existing &lsquo;location' values (City, State, etc) are only captured from the first certificate. SANs are however read from all the certs you nominate (e.g. Default,WebServicesInternal,WebServicesExternal)  & de-duped

- You can't remove SANs that are no longer required

- You can only request an Offline certificate when running this script on an Edge server (and you'll need to specify "-ComputerFqdn $null" to stop it attempting to read the Topology). Another approach is to copy the Edge's cert to  an FE and do it from there!

- Outputs to screen using "write-host" and not to the pipeline

- PowerShell v2 (Server 2008R2 & Windows 7) doesn't reveal all of the values to me. Rather than prevent the script from running under v2, I pop a warning

- Needs to be run as Admin to reliably show all information

## How-To
Here are some examples of how to use it:
This is your minimum to refresh an existing FE cert. This makes no changes to the cert's values (other than the new Friendly name of course):

```powershell
PS C:\> .\Update-SfBCertificate.ps1 -type default,webservicesinternal,webservicesexternal -FriendlyName "Combined SYD Pool FE cert 27Apr2017" -ca MyCA.contoso.com\My-CA
```

Refresh your existing FE cert. Add all the pool FQDNs and make sure the SIP domains are all there too:

```powershell
PS C:\> .\Update-SfBCertificate.ps1 -type default,webservicesinternal,webservicesexternal -Verbose -FriendlyName "Combined SYD Pool FE cert 27Apr2017" -ca MyCA.contoso.com\My-CA -AllPoolMemberServers -AllSipDomain
```

Has someone recently updated the server's cert and overlooked a SAN? We've all been there. Use the thumbprint of the old one and inject the SAN:

```powershell
PS C:\> .\Update-SfBCertificate.ps1 -type default,webservicesinternal,webservicesexternal -Verbose -FriendlyName "Combined SYD Pool FE cert 27Apr2017" -ca MyCA.contoso.com\My-CA -Thumbprint "?12 34 56 78 90 12 34 45 78 89" -Domain "OverlookedSAN.contoso.com"
```

## Show me
Here I'm requesting a new default,webservicesinternal,webservicesexternal certificate for my Front-End. I'm providing the details of the CA, specifying a new Friendly Name, making sure the City comes out as "Sydney", adding all SIP  Domains, all FE's in the pool, and adding SANs for meet.contoso & meet.fabrikam:

<img src="https://user-images.githubusercontent.com/11004787/81052514-e606c300-8f06-11ea-8832-66672dfebe29.png" alt="" width="600" />

The script pauses at this point. On-screen (to the left of the green bar here) are all the values that will be fed into "Request-CsCertificate" if you agree to the Confirm prompt at the bottom.

<img src="https://user-images.githubusercontent.com/11004787/81052556-ffa80a80-8f06-11ea-9235-0dd42098206e.png" alt="" width="600" />

And here's the comparison of old/existing and new. The expected differences between them are highlighted in yellow. Any that aren't expected are shown in red (or whatever your Warning and Error colour preferences are set to, respectively).
If you like the look of the new certificate, right-click to paste the correctly-formatted "Assign-CsCertificate" command to make it active.
&nbsp;
## <span style="font-size: 1.17em;">Script Revision History
#### v1.3: 12th May 2018

- Added an abort line that kills the script when running in the (unsupported) PowerShell ISE. (Screen-width and coloured output don't work)

#### v1.2: 24th December 2017

- Improved the way the "Subject" is parsed in ParseCertSubject by trimming leading spaces

- Added "E=" for those certs where an e-mail address has been provided. (Not applicable to new cert requests via the "Request-CsCertificate")

- Changed the cert comparison highlighting: no longer shows in 'warning' colour if the user provided a 'new' value but nothing changed in the resulting cert

- Incorporated my version of Pat's "Get-UpdateInfo". Credit: https://ucunleashed.com/3168

#### v1.1: 19th February 2017 - the "Thank You Mike Shivtorov" bugfix & suggestions release

- Added 'XmppServer' certificate type, overlooked in the original release

- Changed Output request file example text from ".pfx" to ".req"

- Added another example to better document how to generate Edge certificates

- Improved offline request process: now finds and re-opens the offline request, then feeds it to the comparison engine for display

- Added an extra Exception trap to the Request-CsCertificate handling: script now reports cleanly if you request an inappropriate Type

- Sorted Key Usages before sending them to the Compare engine in an effort to reduce false positives

- Fixed bug where KeySize was reported in red instead of yellow when the user provided a new value

#### v1.0: 7th May 2016. Initial public release

- Improved the error reporting when Request-CsCertificate fails

&nbsp;

<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/update-sfbcertificate-ps1/](https://greiginsydney.com/update-sfbcertificate-ps1/).
