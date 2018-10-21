# Google Docs Scripts

## Ordini v1.1
#### Creating offer and order documents from a catalogue sheet

### What's new
- Added a menu item to clean the sheet
- Thousands separators

### Quick Start

- install nodejs version > 8.10 < 9 
- clone repo: `git clone https://github.com/olivettiscuoladigitale/hw-gestione-offerte` 
- inside dir Ordini: `npm install` 
- Done

### Compiling

- tsc --pretty
- Done

### Pulling code to GoogleDocs
- clasp login # (this is needed only once)
- clasp pull

or 

- clasp login # (this is needed only once)
- clasp clone <scriptId>

### Pushing code to GoogleDocs
- clasp login # (this is needed only once)
- clasp push

### Order Usage 

- Open "Catalogo Olivetti File"
- Select items checkboxes and quantity for each item
- Select "Nuovo Ordine" from "-Ordini OSD-" menu
- Fill-in the module
- A new order sheet will be created in your Google Documents

### Offer Usage 

- Open "Catalogo Olivetti File"
- Select items checkboxes and quantity for each item
- Select "Nuova Offerta" from "-Ordini OSD-" menu
- Fill-in the module
- A new offer doc  will be created in your Google Documents

