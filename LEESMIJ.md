# Zakkingskaart: zo draai je het script

De Zakkingskaart laat zien hoe snel gebouwen in Nederland zakken. Voor heel Nederland gebruikt hij de NAP-peilmerken van Rijkswaterstaat; voor Amsterdam komt daar een veel fijnere laag bij: de meetbouten van de gemeente.

Het script `zakkingskaart.py` haalt alle meetbouten en metingen op bij de gemeente Amsterdam (en in website-modus ook de NAP-peilmerken). Het berekent per meetbout hoe snel de gevel zakt en maakt daar een kaart van die je in je browser opent. Je hebt alleen Python 3.8 of nieuwer nodig, geen extra pakketten.

## 1. Eenmalig: Python controleren

**Mac**
1. Open **Terminal** (Cmd+spatie, typ "Terminal").
2. Typ `python3 --version` en druk op Enter.
   - Zie je `Python 3.8` of hoger? Dan ben je klaar.
   - Vraagt de Mac om "ontwikkelaarstools" te installeren, klik dan op Installeer en probeer het daarna opnieuw.
   - Werkt het niet? Installeer Python via https://www.python.org/downloads/. Ga daarna in Finder naar **Programma's > Python 3.x** en dubbelklik op **Install Certificates.command**.

**Windows**
1. Installeer Python via https://www.python.org/downloads/ en vink **"Add python.exe to PATH"** aan.
2. Open **Opdrachtprompt** en typ `py --version`.

## 2. Script klaarzetten

Maak een map aan, bijvoorbeeld `Zakkingskaart` in je Documenten, en zet `zakkingskaart.py` erin.

Ga in Terminal naar die map:

```
cd ~/Documents/Zakkingskaart
```

(Windows: `cd %USERPROFILE%\Documents\Zakkingskaart`)

## 3. Eerst een proefrun (ongeveer 1 minuut)

```
python3 zakkingskaart.py --proef
```

(Windows: `py zakkingskaart.py --proef`)

Dit haalt een klein deel van de data op en opent een kaart. De meeste punten zijn dan grijs, en dat is normaal. Het doel is alleen te controleren dat alles werkt.

## 4. De volledige kaart

```
python3 zakkingskaart.py
```

- De eerste keer duurt het **ongeveer 10 tot 30 minuten**, omdat het om honderdduizenden metingen gaat. Je ziet de voortgang in het venster.
- Wordt het onderbroken (laptop dicht, Ctrl+C, wifi weg)? Start hetzelfde commando opnieuw, dan gaat het verder waar het was.
- Daarna opent de kaart vanzelf in je browser.

De uitvoer staat in de map `zakkingskaart_uitvoer`:

| Bestand | Wat |
|---|---|
| `zakkingskaart.html` | de kaart; dubbelklik om opnieuw te openen |
| `meetbouten_analyse.csv` | alle cijfers per meetbout, opent in Excel |
| `cache/` | de opgehaalde data, zodat een volgende run in seconden klaar is |

Voor de achtergrondkaart heb je internet nodig.

## 5. Handige varianten

| Doel | Commando |
|---|---|
| Nieuwste metingen ophalen | `python3 zakkingskaart.py --vernieuw` |
| Alleen Zuid en Oost tonen (sneller in de browser) | `python3 zakkingskaart.py --bbox 52.32,4.83,52.38,5.00` |
| Kleiner bestand, zonder grafiekjes | `python3 zakkingskaart.py --zonder-grafieken` |
| Website maken (zie hieronder) | `python3 zakkingskaart.py --site` |
| Alle opties | `python3 zakkingskaart.py --help` |

## Website maken (`--site`)

```
python3 zakkingskaart.py --site
```

Dit maakt in plaats van één groot bestand een kleine website in de map `site`:

| Bestand | Wat |
|---|---|
| `index.html` | de kaart zelf, zonder data, laadt snel |
| `data.json` | alle meetbouten met hun cijfers |
| `reeksen.json` | de meetreeksen voor de grafiekjes; wordt pas geladen nadat de kaart er staat |
| `leaflet/` | de kaartbibliotheek, zelf gehost, zodat bezoekers geen verzoeken naar een extern CDN doen |
| `meetbouten_analyse.csv` | alle cijfers per meetbout, als download |
| `.nojekyll` | zorgt dat GitHub Pages de bestanden ongewijzigd publiceert |

De data blijft bewaard in de map `cache`, die bewust **buiten** `site` staat, zodat hij niet wordt gepubliceerd. Bij elke volgende run:
- worden de meetbouten opnieuw opgehaald (klein, en status of adres kan wijzigen);
- worden van de metingen **alleen de nieuwe** opgehaald: alles vanaf de laatst bekende datum min 2 jaar (`--terugkijk 730`), zodat ook laat ingevoerde metingen meekomen. Nieuwe en gewijzigde metingen worden samengevoegd met de cache;
- worden alle metingen eens per 90 dagen toch volledig opnieuw opgehaald (`--volledig-na 90`). Zo komen ook correcties en verwijderingen in oudere metingen mee. `--vernieuw` doet dat meteen;
- stopt het script **zonder iets te publiceren** als de gemeente veel minder levert dan de vorige keer (minder dan 90% van de meetbouten of 95% van de metingen). De bestaande site blijft dan staan. Klopt het toch, draai dan met `--forceer`.

De eerste keer haalt het script ook Leaflet 1.9.4 op en controleert het die bestanden op de officiële hashes. Kloppen die niet, dan stopt het.

**Lokaal bekijken:** dubbelklikken op `index.html` werkt niet, omdat de pagina `data.json` ophaalt. Start daarom een mini-webserver:

```
python3 -m http.server 8000 --directory site
```

en open http://localhost:8000 in je browser. Stoppen doe je met Ctrl+C.

**In de repository** horen alleen het script en de documentatie. Zet deze regels in `.gitignore`:

```
cache/
cache_proef/
site/
zakkingskaart_uitvoer/
```

## Automatisch publiceren via GitHub

De workflow `.github/workflows/publiceer.yml` draait elke maandagochtend, bij een wijziging in het script, en handmatig via het tabblad **Actions** (knop **Run workflow**, met de opties "alles opnieuw ophalen" en "forceer"). Hij:
1. haalt de cache op van de branch `data-cache`;
2. draait `zakkingskaart.py --site`;
3. bewaart de bijgewerkte cache weer op `data-cache` (één commit, steeds overschreven, door `github-actions[bot]`). Die wekelijkse push telt ook als activiteit, zodat GitHub het schema niet na 60 dagen stilzet;
4. publiceert de map `site` op GitHub Pages.

Eenmalig instellen: **Settings → Pages → Source: GitHub Actions**. Voor feedback zet je onder **Settings → Secrets and variables → Actions** het secret `ZK_FEEDBACK_URL` (zie hieronder). Mislukt een run, dan stuurt GitHub een e-mail naar het adres van het account. De vorige versie van de site blijft dan gewoon online.

## NAP-peilmerken (heel Nederland)

In website-modus haalt het script ook de NAP-peilmerken van Rijkswaterstaat op: ongeveer 50.000 bronzen bouten in huizen, bruggen en andere bouwwerken, met per peilmerk alle gemeten hoogtes (ruim 500.000 metingen). Bron: de openbare service `geo.rijkswaterstaat.nl/arcgis/rest/services/GDR/nap/MapServer` (lagen 0 en 1 plus tabel 4). De licentie is CC0.

- **Verversen:** peilmerken worden maar eens per 5 tot 15 jaar gemeten, in grote meetcampagnes. Daarom haalt het script niet elke week alles op (ruim 250 verzoeken, ongeveer een uur). Om de 3 weken doet het een korte controle (`--nap-controle 21`, een paar verzoeken): het vraagt het aantal peilmerken en metingen op en de nieuwste datums. Is daar iets aan veranderd, dan haalt het alles opnieuw op. Daarnaast haalt het eens per kwartaal toch alles op (`--nap-ververs 91`), voor correcties die het aantal niet veranderen. De eerste keer ophalen kan meerdere runs duren (zie hieronder).
- **Rekenmethode:** van elk peilmerk worden, als ze er zijn, alleen de **herziene** waarden gebruikt (project-id met `=`, het huidige NAP-stelsel). Anders staan oude en herrekende hoogtes door elkaar en ontstaan nepsprongen van ongeveer 2 cm. Metingen binnen één jaar worden samengevoegd. Een reeks met meer dan 1 cm verschil binnen één campagne, een stap van meer dan 25 mm/jaar of meer dan 2 mm/jaar stijging telt als **onzeker** en wordt grijs. De recente snelheid wordt berekend zoals bij de meetbouten, met minimaal 8 jaar tussen de metingen.
- **Op de kaart:** de laag staat standaard aan en de kaart opent op heel Nederland. Uitgezoomd (onder zoomniveau 11) zie je gekleurde vakjes: de mediaan van de peilmerken in dat vak. Klik op een vakje of zoom in voor de losse peilmerken (open rondjes). De meetreeksen (`nap_reeksen.json`) worden pas geladen als iemand op een peilmerk klikt. Zet een bezoeker de laag zelf uit, dan onthoudt de browser die keuze. Je kunt ook zoeken op peilmerknummer, bijvoorbeeld `025G0385`. Via **"Alle meetgegevens bekijken"** in de popup haalt de browser van de bezoeker alle metingen van dat peilmerk live op bij Rijkswaterstaat (één verzoek) en toont ze in een tabel, met de ruwe JSON eronder. Oude waarden van vóór de herziening van het NAP-stelsel zijn grijs, de periode van de recente snelheid is geel. Lukt het ophalen niet, dan toont het venster de opgeslagen meetreeks.
- **Rustig ophalen:** de beveiliging van Rijkswaterstaat remt adressen af die veel nieuwe verbindingen openen (je ziet dan `timed out` of `Connection refused`). Het script doet daarom alle verzoeken via één blijvende verbinding, wacht 1 seconde tussen verzoeken (`--nap-pauze`) en bladert op objectid. Hangt het toch, dan wacht het 2, 5, 10 en 15 minuten. Per run besteedt het hoogstens `--nap-max-minuten` aan de peilmerken (standaard 60; de workflow gebruikt 100).
- **Als het misgaat:** wat binnen is, wordt bewaard in `nap_voortgang.json.gz` in de cache, en de volgende run gaat daar verder. De workflow heeft daarvoor een dagelijkse run (05:41 UTC) die alleen iets doet zolang dat bestand bestaat; je hoeft dus niets te doen. Sneller kan door de workflow met de hand te starten (zonder 'vernieuw', anders begint hij opnieuw). Tot de data compleet is, wordt de oude cache gebruikt; zonder cache wordt de laag overgeslagen en blijft de rest van de site gewoon werken. Uitzetten kan met `--zonder-nap`.

## Disclaimer en vragen over de gegevens

- **Akkoord, elke 30 dagen:** een bezoeker ziet een kort venster ("Voordat je de kaart opent", drie punten). Daaronder staat dat je, door verder te gaan, akkoord gaat met gebruik van de kaart zoals hij is, zonder er rechten aan te kunnen ontlenen. Pas na een klik op "Akkoord, naar de kaart" verschijnt de kaart. De browser onthoudt het akkoord 30 dagen (localStorage, geen cookie; sleutel `zakkingskaart.intro` met versie en tijdstip). Daarna, in een andere browser, in een privévenster of na het wissen van sitegegevens verschijnt het venster opnieuw. Pas je de tekst later inhoudelijk aan, verhoog dan de versie in de code, zodat iedereen hem opnieuw ziet. Een link naar `#over` opent direct het Over-venster.
- **Geen oordeel of advies:** de site geeft geen advies, ook niet om een onderzoek te laten doen. In elke pop-up staat één regel "Openbare meetgegevens, zonder oordeel of advies", en rechtsonder de kaart "Geen oordeel of advies".
- **Vragen en fouten gaan naar de bron:** het Over-venster en de meetgegevens-vensters verwijzen rechtstreeks naar de bron. Voor meetbouten is dat de gemeente Amsterdam, gebruik.basisinformatie@amsterdam.nl, het contactadres bij de dataset Kernregistratie Meetbouten op data.overheid.nl. Voor NAP-peilmerken is dat het contactformulier van de Servicedesk Data van Rijkswaterstaat. Wat de bron aanpast of verwijdert, verdwijnt bij de volgende automatische update vanzelf van de kaart. De site belooft zelf niets en noemt geen termijn.
- **Noodknop (niet op de site genoemd):** wil je zelf toch eens een meetpunt weghalen, dan kan dat met de niet-openbare repository-variabele **`ZK_VERBERGEN`** (Settings → Secrets and variables → Actions → Variables). Zet er per regel een adres zoals het op de kaart staat in (dan verdwijnen alle meetbouten op dat adres), een meetboutnummer of een peilmerknummer. Start daarna de workflow met de hand. Het log toont alleen het aantal regels. Lokaal kan hetzelfde met `verbergen.txt` naast het script, of met `--verbergen <bestand>`.

## Bezoekersstatistieken

GitHub Pages en de domeinregistrar geven geen bezoekersaantallen. Daarom telt de site anoniem via hetzelfde Google Apps Script als de feedback (`feedback/Code.gs`). Per bezoek meldt de kaart alleen dat hij is geopend, of dat op mobiel of desktop gebeurt en wat de herkomst is (Google, andere zoekmachine, social, andere site of direct). Bij het verlaten volgt hoe vaak er is gezocht, een punt is bekeken en meetgegevens zijn geopend. Opgeslagen worden alleen totalen per dag, plus per maand de domeinnamen van verwijzende sites. Er worden geen cookies, IP-adressen of kenmerken per bezoeker opgeslagen, en lokaal of door bots wordt niets geteld.

**Aanzetten:**
1. Zet de nieuwe `Code.gs` in je Apps Script-project en kies **Implementeren → Implementaties beheren → bewerken (potlood) → Versie: Nieuwe versie → Implementeren**. De URL blijft gelijk.
2. Kies bovenin de editor de functie **installeer** en klik op **Uitvoeren**. Geef de gevraagde toestemming.
3. Je krijgt een mail met de link naar je privé-statistiekenpagina. Elke maandagochtend volgt een weekoverzicht.

De teller gebruikt de feedback-URL (`ZK_FEEDBACK_URL`). Wil je hem uitzetten, gebruik dan `--zonder-teller`. Deel de statistiekenlink niet: wie hem heeft, kan de cijfers zien.

## Vindbaar in Google

De site heeft een paginatitel, een beschrijving, een `robots.txt` en een `sitemap.xml`. Om Google de site te laten kennen en zoekstatistieken te zien:
1. Ga (ingelogd met het Google-account van de kaart) naar Google Search Console en kies **Property toevoegen → URL-voorvoegsel**: `https://zakkingskaart.github.io/`.
2. Kies als verificatiemethode **HTML-tag** en kopieer alleen de code tussen `content="` en `"`.
3. Zet die code in GitHub onder **Settings → Secrets and variables → Actions → tabblad Variables** als variabele `ZK_GOOGLE_VERIFICATIE`, en start de workflow handmatig.
4. Klik in Search Console op **Verifiëren** en dien onder **Sitemaps** `sitemap.xml` in.

## Eigen domein (optioneel)

1. Verifieer eerst het domein bij GitHub: **profiel-Settings → Pages → Add a domain** (een TXT-record). Dat voorkomt dat iemand anders het domein aan een eigen GitHub-site kan koppelen.
2. Zet bij je domeinregistrar deze DNS-regels:
   - `A` voor het domein zelf: `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153`
   - `AAAA`: `2606:50c0:8000::153`, `2606:50c0:8001::153`, `2606:50c0:8002::153`, `2606:50c0:8003::153`
   - `CNAME` voor `www`: `zakkingskaart.github.io`
3. Vul in de repo bij **Settings → Pages → Custom domain** het domein in. Zet **Enforce HTTPS** aan zodra dat kan.
4. Zet de variabele `ZK_SITE_URL` op het nieuwe adres (bijvoorbeeld `https://zakkingskaart.nl/`) en start de workflow. Het oude github.io-adres stuurt dan vanzelf door.

Stop je later met het domein, haal het dan **eerst** weg bij Custom domain en laat het daarna pas verlopen.

## Feedback aanzetten (optioneel)

De kaart heeft een knop **"★ Feedback"**. Bezoekers geven daarmee 1 tot 5 sterren, met eventueel een korte toelichting (maximaal 280 tekens). Er wordt geen naam of e-mailadres gevraagd.

De feedback komt binnen via een klein Google Apps Script (`feedback/Code.gs`) dat draait onder het Google-account van de kaart. Het script mailt elke beoordeling naar de mailbox van dat account. Het e-mailadres staat daardoor niet in de code en niet op de website; bezoekers zien het nergens. Er komt geen extra formulierdienst aan te pas. Meer dan 50 berichten per dag worden genegeerd, tegen spam.

**Eenmalig instellen** (ingelogd met het Google-account van de kaart):
1. Ga naar https://script.google.com en kies **Nieuw project**. Noem het bijvoorbeeld "Zakkingskaart feedback".
2. Vervang de inhoud van `Code.gs` door die van `feedback/Code.gs` en klik op **Opslaan**.
3. Kies **Implementeren → Nieuwe implementatie**. Klik op het tandwiel bij "Type selecteren" en kies **Web-app**. Vul in:
   - **Uitvoeren als:** Ik
   - **Wie heeft toegang:** Iedereen
4. Klik op **Implementeren** en geef toestemming. Google waarschuwt dat de app "niet geverifieerd" is: kies **Geavanceerd → Ga naar Zakkingskaart feedback**. Dat is normaal voor een eigen script.
5. Kopieer de **URL van de web-app** (eindigt op `/exec`). Open die in je browser; je hoort `{"ok":true,"bericht":"feedback-ontvanger actief"}` te zien.
6. Zet die URL als secret `ZK_FEEDBACK_URL` in GitHub, of maak lokaal de kaart met:

```
python3 zakkingskaart.py --site --feedback-url https://script.google.com/macros/s/.../exec
```

Pas je het script later aan, kies dan **Implementeren → Implementaties beheren → bewerken (potlood) → Versie: Nieuwe versie**. Zo blijft de URL hetzelfde.

De URL komt in de website te staan; dat kan niet anders. Hij bevat alleen een willekeurige code, geen e-mailadres. Iemand met de URL kan je hooguit berichten sturen, tot de daglimiet.

## 6. De kaart gebruiken

- **Zoeken:** typ een adres, straat, postcode of plaats ergens in Nederland, een meetbout-nummer of een peilmerknummer. Amsterdamse straten worden direct in de eigen gegevens gevonden. Al het andere gaat via de PDOK Locatieserver van de overheid (alleen de zoektekst wordt verstuurd).
- **Open rondjes** zijn NAP-peilmerken (heel Nederland). **Lijnen** (alleen Amsterdam) zijn straatdelen per bouwblok, gekleurd op de mediaan van de meetbouten daarin. **Punten** zijn de losse meetbouten.
- **Klik op een lijn** voor een lijst van alle bouten in dat straatdeel. **Klik op een punt** voor de cijfers en een grafiek van de zakking.
- **Kleuren** (standaard):
  - rood: 2,0 mm/jaar of meer
  - geel: 1,0 tot 2,0 mm/jaar
  - groen: minder dan 1,0 mm/jaar
  - groen met donkere rand of stippellijn: **gestabiliseerd**, eerst minstens 1,5 mm/jaar en nu minder dan 0,5 mm/jaar
  - grijs: geen recente data (laatste meting meer dan 10 jaar oud, bout vervallen, of meetreeks te kort)
- De namen zijn "Zakt snel", "Zakt matig" en "Zakt nauwelijks". De grenzen van 1 en 2 mm/jaar zijn ook klassegrenzen op de meetboutenkaart van de gemeente Amsterdam (`meetbouten.map` in github.com/Amsterdam/mapserver), die ook van groen via geel naar rood kleurt. De gemeente kleurt wel op de gemiddelde snelheid sinds de eerste meting.
- Via **"Grenzen aanpassen"** verander je alle grenzen. Daar kun je ook oudere metingen alsnog laten kleuren.

## 7. Hoe de snelheid berekend wordt

**Recente snelheid** = het verschil tussen de laatste meting en de meest recente meting die minstens 4 jaar eerder ligt, gedeeld door de tijd daartussen. **Eerdere snelheid** = van de eerste meting tot het begin van die recente periode.

Houd rekening met het volgende:
- Een meting is op ongeveer 1 mm nauwkeurig. Over korte periodes is de snelheid dus wat onzeker.
- Een meetbout zit in één gevel. De kleur zegt iets over die plek, niet over elk pand in de straat.
- Grijs betekent "onbekend", niet "goed".
- "Gestabiliseerd" is een aanwijzing voor herstel, geen bewijs. Check het bouwdossier of vraag de VvE.

## Problemen?

| Melding | Oplossing |
|---|---|
| `CERTIFICATE_VERIFY_FAILED` | Mac: dubbelklik op **Install Certificates.command** (zie stap 1) |
| `command not found: python3` | Python installeren (stap 1), op Windows `py` gebruiken |
| `server antwoordt 429/503` | Het script wacht en probeert het zelf opnieuw |
| `HTTP 400` of een andere foutmelding | Kopieer de laatste regels uit het venster en bewaar ze om het probleem uit te zoeken |
| Kaart blijft leeg | Controleer je internetverbinding; de kaartbibliotheek wordt online geladen |
