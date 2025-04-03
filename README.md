# Alfa-college FAQ Builder

![Version](https://img.shields.io/badge/versie-2.2.0-blue)
![SPFx](https://img.shields.io/badge/SPFx-1.16.1-green)
![Node.js](https://img.shields.io/badge/Node.js-16.x-yellow)

Een SharePoint Framework webpart voor het weergeven van FAQ's in een gestructureerd accordeon formaat met categoriefilters.

## Functionaliteiten

- **Categoriefiltering:** Filtert vragen per categorie via tabs
- **Ondersteuning voor multi-select categorieën:** Items kunnen in meerdere categorieën tegelijk worden getoond
- **Zoekfunctionaliteit:** Doorzoekt zowel vragen als antwoorden
- **Automatisch uitklappen:** Zoekresultaten worden automatisch uitgeklapt
- **Gebruikersvriendelijke interface:** Eenvoudig te navigeren accordeon-stijl voor FAQ's
- **Logging mogelijkheid:** Optionele logging van items die gebruikers bekijken via webhook

## Screenshots

![AlfaFaq Webpart](https://example.com/placeholder-screenshot.png)

## Installatie

1. Download het `.sppkg` bestand uit de `/sharepoint/solution/` map
2. Upload het naar je organisatie's App Catalog
3. Voeg het webpart toe aan een pagina

## Configuratie

Het webpart heeft de volgende configuratiemogelijkheden:

- **Lijst selectie:** Kies de SharePoint lijst met FAQ gegevens
- **Categorie kolom:** Kies een keuzeveld (single of multi-select) voor categoriefiltering
- **Titel kolom:** De kolom die gebruikt wordt voor de vragen
- **Inhoud kolom:** De kolom die gebruikt wordt voor de antwoorden
- **Sorteerkolom:** Bepaal op welke kolom de items gesorteerd worden
- **Sorteerrichting:** Bepaal of sortering oplopend of aflopend is
- **Layout opties:** Configureer of meerdere items tegelijk uitgeklapt kunnen worden
- **Logging:** Optionele configuratie voor het loggen van bekeken items via webhook

## Ontwikkelomgeving

### Vereisten

- Node.js v16.x (`nvm install 16.13.0`)
- Yeoman en SPFx generator (`npm install -g yo @microsoft/generator-sharepoint`)
- Gulp CLI (`npm install -g gulp-cli`)

### Setup

```bash
# Clone de repository
git clone <repository-url>

# Navigeer naar de projectmap
cd AlfaFaq

# Installeer dependencies
npm install --legacy-peer-deps

# Start de lokale ontwikkelserver
gulp serve
```

### Bouwen voor productie

```bash
# Bouw het project voor productie
gulp bundle --ship
gulp package-solution --ship
```

Het `.sppkg` bestand wordt gemaakt in de `/sharepoint/solution/` map.

## Recent toegevoegde functionaliteiten

### V2.2.0 - April 2025
- Ondersteuning voor multi-select keuzevelden toegevoegd voor categoriefiltering
- Verbeterde filtering en tabweergave

### V2.1.0 - Eerdere release
- Automatische logging van bekeken items via webhook
- Verbeterde zoekfunctionaliteit

## Ontwikkelaars

Ontwikkeld door Richard ten Klooster, [Goldhand](https://goldhand.nl)

## Licentie

Copyright © 2025 Alfa-college
