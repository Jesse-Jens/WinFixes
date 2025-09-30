WinFixes

Useful Windows fixes, files and single‑command runs.

Teams Meeting Add‑in Repair

This repository includes an automated repair script for the Microsoft Teams Meeting Add‑in for Outlook. Use it to reinstall or repair the add‑in on any Windows machine without manually copying files or clicking through installers.

Voorbereiding

Plaats het ZIP‑bestand TeamsMeetingAddin.zip in de root van deze repository. Dit archief moet de map TeamsMeetingAdd‑in en de corresponderende MSI‑installer bevatten.

Zorg dat de bestandsnaam exact overeenkomt; het script verwijst naar TeamsMeetingAddin.zip in deze repository.

Een‑regel reparatie (IRM)

Voer de volgende PowerShell‑opdracht uit om de reparatie in één stap te starten. Er zijn geen parameters nodig:

irm https://raw.githubusercontent.com/Jesse-Jens/WinFixes/main/install-teamsmeeting-addin.ps1 | iex


Deze opdracht doet het volgende:

Downloadt het script install-teamsmeeting-addin.ps1 rechtstreeks uit dit GitHub‑repository.

Voert het script uit in de huidige PowerShell‑sessie.

Het script downloadt het ZIP‑bestand, extraheert de map TeamsMeetingAdd‑in naar %LOCALAPPDATA%\Microsoft en start de MSI‑installer in stille modus (/fa /quiet /norestart) om de add‑in te repareren. Als de stille reparatie niet slaagt, wordt de installer interactief gestart zodat je deze zelf kunt voltooien.

Hoe werkt het?

Het script maakt gebruik van PowerShell’s Invoke-WebRequest en Expand-Archive om bestanden te downloaden en uit te pakken. Door de download‑URL van de zip in het script zelf te definiëren, hoef je geen parameters meer mee te geven. Het commando irm ... | iex (alias voor Invoke-RestMethod en Invoke-Expression) zorgt ervoor dat de inhoud van het script direct wordt uitgevoerd nadat het is opgehaald.

Zie de header van het script voor verdere uitleg over de werkwijze en aanpassingen.