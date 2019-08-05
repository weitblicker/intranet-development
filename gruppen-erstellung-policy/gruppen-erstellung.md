# Erstellung von Gruppen

Wir haben die Erstellung von Gruppen durch Standarduser deaktiviert. Durch das PowerShell Skript [GroupCreationControl.ps1](./GroupCreationControl.ps1) kann die Erstellung für Standarduser aktiviert und deaktiviert werden.

Hierzu einfach die Variable `$AllowGroupCreation` entweder auf `True` (Erstellung ist möglich) oder `False` (Erstellung ist deaktiviert) setzen.

***Anforderungen***

Bitte beachte, dass ggf. die folgenden Anforderungen erfüllt werden müssen:

1. Du benötigst ggf. eine "Preview" bzw. Vorschauversion der Azure ActiveDirectory PowerShell Module. Falls dies so ist, kannst du diese entsprechend der weiterführenden Informationen (Microsoft Dokumentation) installieren.
2. Vor Ausführung muss ggf. die Execution Policy deiner PowerShell verändert werden (bspw. ByPass) 

***Weiterführende Informationen:***

- [Microsoft Dokumentation (Deutsch)](https://docs.microsoft.com/de-de/office365/admin/create-groups/manage-creation-of-groups?view=o365-worldwide)
- [Microsoft Dokumentation (Englisch)](https://docs.microsoft.com/en-us/office365/admin/create-groups/manage-creation-of-groups?view=o365-worldwide)
