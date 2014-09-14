FJUpdater
=========

Script to testing the relevance of the installed versions of Java &amp; Flash with Sun&amp;Adobe web-sites or local source, and automatic updates (via the Internet (HTTP) or local area network (SMB), respectively) if necessary. The script can save installer of current version Java &amp; Flash in the specified folder (+ file with the version number)  which makes it possible to build a mechanism to automatically deploy updates Java &amp; Flash.

Installation & Usage
--------------------
1) Download files and place it on server ("\\\\SERVER\\INSTALL\\" for example)

2) Create share for installation files ("\\\\SERVER\\INSTALL_CLIENT$\\FJUpdater\\" by default) and allow it for reading to all and for write to "reference computer" (see step 5)

3) Edit *"FJUpdater_config.vbs"* file, tune mail settings and *csInstallerPath* (from step 2)

4) Start downloading the installation files:

```sh
cscript \\SERVER\INSTALL\FJUpdater.wsf /WEBModeSaveInstallForce /mail:1 /debug:3
```

5) Configure a regular check for updates on the **reference computer** (virtual machine)
```sh
cscript \\SERVER\INSTALL\FJUpdater.wsf /webmode:1 /WEBModeSaveInstall:1 /mail:1 /debug:3
```

6) Adjust the installation of updates for **other computers** (at startup, using GPO)
```sh
cscript \\SERVER\INSTALL\FJUpdater.wsf /webmode:0 /mail:1 /debug:2
```

7) To see current version of installed plugings use
```sh
cscript \\SERVER\INSTALL\FJUpdater.wsf /ShowVersion
```
for local computer or
```sh
cscript \\SERVER\INSTALL\FJUpdater.wsf /ShowVersion:Computer_Name_Or_IP
```
for remote computer "Computer_Name_Or_IP"



Changelog
--------------

v.2.5

  Single .VBS file was separated into 2 parts - source code (FJUpdater_main.vbs) and config (FJUpdater_config.vbs), which combined using .WSF file.

v.2.6

  added Java 64-bit support



License
----

GPL

