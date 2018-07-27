Infoscherm Lync/Skype for Buisiness Intergration
==============

These scripts are used to enable the lync intergration of the infoscherm (the tv in the board room). 
The scripts read a local link client and see if someone is calling and who, to this end it writes to files in the tvpc account's public html which are read out bij the infoscherm webpage and displayed.
For more information see https://bestwiki.gewis.nl/~bestuur/wiki/index.php?title=Infoscherm

Installation
------------
- Make sure the PC automatically logs in to the tvpc account
- Make sure Lync is installed and logged in to the GEWIS account (only tested with 2010 and 2013)
- Put the powershell scripts in the root of the C: drive 
- Put the batch script in the startup folder (Win+R "shell:startup")
- In browser go to infoscherm.gewis.nl
- See if it works
- You might need to install some additinal SDK's or change some paths to SDK's in the LyncStatusDetect.ps1 script
