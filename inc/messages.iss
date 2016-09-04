; =====================================================================
; == inc\messages.iss
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================

en.OfficeIsRunning=Your Office application must be closed in order to continue installation. If you click 'OK', the application will be shut down.
de.OfficeIsRunning=Ihre Office-Anwendung muß geschlossen werden, um mit der Installation fortzufahren. Wenn Sie auf 'OK' klicken, wird die Anwendung geschlossen.

en.SingleOrMulti=Single-user or system-wide install
en.SingleOrMultiSubcaption=Install for the current user only or for all users
en.SingleOrMultiDesc=Please indicate the scope of this installation:
en.SingleOrMultiSingle=Single user (only for me)
en.SingleOrMultiAll=All users (system-wide)
en.Office2007Required=This add-in requires Office 2007 or later. Setup will now terminate.

; CannotInstallPage [EN]
en.CannotInstallCaption=Administrator privileges required
en.CannotInstallDesc=You do not have the necessary rights to install additional required runtime files.
en.CannotInstallMsg=Additional runtime files from Microsoft are required to run this add-in. You may continue the installation, but the add-in won't start unless the required runtime files are installed by an administrator. Note: On Windows Vista and newer, right-click the installer file and choose 'Run as administrator'.
en.CannotInstallCont=Continue anyway, although it won't work without the required runtime files
en.CannotInstallAbort=Abort the installation (come back when the admin has installed the files)

; DownloadInfoPage [EN]
en.RequiredCaption=Additional runtime files required
en.RequiredDesc=Additional runtime files for the .NET framework from Microsoft are required in order to run this add-in.
en.RequiredMsg=%d file(s) totalling about %s MiB need to be downloaded from the Microsoft servers. Click 'Next' to start downloading.

; InstallInfoPage [EN]
en.InstallCaption=Runtime files downloaded
en.InstallDesc=The required runtime files are ready to install.
en.InstallMsg=Click 'Next' to beginn the installation.

en.StillNotInstalled=The required additional runtime files are still not installed. Setup will continue, but unless you ensure that the runtimes are properly installed, the add-in will not function properly.
en.DownloadNotValidated=A downloaded file has unexpected content. It may have not been downloaded correctly, or someone might have hampered with it. You may click 'Back' and then 'Next' to download it again.

; General messages [DE]
de.SingleOrMulti=Einzelner oder alle Benutzer
de.SingleOrMultiSubcaption=Geben Sie an, für wen die Installation sein soll
de.SingleOrMultiDesc=Bitte geben Sie an, ob das Addin nur für Sie oder für alle Benutzer installiert werden soll.
de.SingleOrMultiSingle=Ein Benutzer (nur für mich)
de.SingleOrMultiAll=Alle Benutzer (systemweit)
de.Office2007Required=Dieses Add-in läuft nur auf Office 2007 und neueren Versionen.

; "Download required" messages (.NET and VSTOR runtimes) [DE]
de.CannotInstallCaption=Administratorrechte benötigt
de.CannotInstallDesc=Sie haben nicht die erforderlichen Benutzerrechte, um weitere benötigte Laufzeitdateien zu installieren.
de.CannotInstallMsg=Sie können mit der Installation fortfahren, aber das Addin wird nicht starten, solange die VSTO-Laufzeitdateien nicht von einem Admin installiert wurden. Tipp: Wenn Sie Windows Vista oder neuer verwenden, klicken Sie mit der rechten Maustaste auf die Installationsdatei und wählen "Als Administrator ausführen".
de.CannotInstallCont=Trotzdem installieren, obwohl es nicht funktionieren wird
de.CannotInstallAbort=Installation abbrechen

; DownloadInfoPage [EN]
de.RequiredCaption=Weitere Laufzeitdateien erforderlich
de.RequiredDesc=Weitere Laufzeitdateien für das .NET-Framework von Microsoft werden benötigt, um das Addin verwenden zu können.
de.RequiredMsg=%d Datei(en) mit ca. %s MiB muß/müssen von den Microsoft-Servern heruntergeladen werden. Klicken Sie 'Weiter', um den Download zu beginnen.

; InstallInfoPage [EN]
de.InstallCaption=Weitere .NET-Laufzeitdateien heruntergeladen
de.InstallDesc=Die zusätzlichen benötigten Dateien von Microsoft können jetzt installiert werden.
de.InstallMsg=Klicken Sie 'Weiter', um mit der Installation zu beginnen.

de.StillNotInstalled=Die zusätzlichen benötigten Dateien wurden leider nicht korrekt installiert. Die Installation des Addins wird jetzt zwar fortgesetzt, aber solange Sie nicht die erforderlichen Laufzeitdateien installieren, wird es nicht funktionieren.
de.DownloadNotValidated=Es wurde unerwarteter Inhalt in einer heruntergeladenen Datei gefunden. Die Installation kann so nicht fortgesetzt werden. Sie können aber 'Zurück' und dann 'Weiter' klicken, um den Download neu zu beginnen.
