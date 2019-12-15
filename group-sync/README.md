# Gruppensync

Für jede Stadt existiert eine Office-365-Gruppe in der alle Mitglieder hinzugefügt werden.
Zusätzlich existiert eine E-Mail-aktivierte Sicherheitsgruppe über die die Leserechte auf die anderen Stadt-Gruppen geregelt sind und die als Mailverteiler dienen.
Damit die Mitglieder beider Gruppen synchron sind existiert ein Powershell-Skript (group_sync_running.ps1), das diese synct.

Die Vorstands-Gruppen werden vom Bundesvorstand verwaltet. Hier funktioniert der Sync in die andere Richtung (E-Mail-aktivierte Sicherheitsgruppe -> Office-365-Gruppe).

Weitere Infos zum hier: https://weitblicker.sharepoint.com/WeitblickWiki/Intranet/Administration/Sharepoints/Gruppenadministration.aspx

Weitere Infos zum Setup hier: https://weitblicker.sharepoint.com/WeitblickWiki/Intranet/Administration/Azure.aspx

## Manueller Sync

Zusätzlich sind Skripte verfügbar, mit denen die Office-365 und Sicherheitsgruppen manuell synchronisiert werden können (default: Alle, mitglieder und vorstand !). Diese wurden während der Umstellung zu Office-365 Gruppen verwendet.