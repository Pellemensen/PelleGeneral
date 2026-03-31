# 50/30/20 Budget Tracker — Handleiding
### Jouw persoonlijke financiële overzicht in Google Sheets

---

## Welkom!

Gefeliciteerd met je aankoop van de 50/30/20 Budget Tracker. Dit template helpt je om grip te krijgen op je geld — zonder ingewikkelde formules of financiële kennis. Alles is al voor je ingesteld. Jij hoeft alleen maar je inkomsten en uitgaven in te vullen.

Deze handleiding legt stap voor stap uit hoe het werkt.

---

## Wat is de 50/30/20 regel?

De 50/30/20 regel is een simpele methode om je geld te verdelen:

| Categorie | Percentage | Wat valt hieronder? |
|---|---|---|
| **Needs (behoeften)** | 50% | Huur, boodschappen, verzekeringen, gas/water/licht |
| **Wants (wensen)** | 30% | Uit eten, kleding, Netflix, hobby's |
| **Savings & Debt (sparen)** | 20% | Spaargeld, schulden aflossen, pensioen |

**Voorbeeld:** Verdien je €3.000 per maand netto?
- €1.500 gaat naar vaste lasten
- €900 mag je besteden aan leuke dingen
- €600 zet je opzij of los je af

---

## Stap 1 — Een kopie maken

> ⚠️ Bewerk nooit het originele bestand. Maak altijd eerst een eigen kopie.

1. Open de link die je bij je aankoop hebt ontvangen
2. Klik bovenaan op **Bestand → Kopie maken**
3. Geef het een naam (bijv. "Budget 2026 - Pelle")
4. Klik op **Kopie maken**

Je hebt nu je eigen versie die je vrij kunt aanpassen.

---

## Stap 2 — De tabbladen begrijpen

Onderaan het scherm zie je vier tabbladen:

### 📊 Dashboard
Dit is je hoofdpagina. Hier zie je in één oogopslag:
- Hoeveel je hebt uitgegeven aan Needs, Wants en Savings
- Of je binnen je budget zit
- Hoeveel geld je nog over hebt

### 📋 Transactions Log
Hier voer je al je inkomsten en uitgaven in. Elke keer dat je iets uitgeeft of ontvangt, voeg je hier een regel toe.

### ⚙️ Settings
Hier stel je in:
- Welke percentages je wilt hanteren (standaard 50/30/20)
- Welke categorieën bij Needs, Wants of Savings horen

### 🔒 Reference
Dit tabblad staat verborgen en bevat achtergrondgegevens voor de formules. Je hoeft hier niets aan te doen.

---

## Stap 3 — Basisinstellingen aanpassen

Ga naar het tabblad **Settings**.

**Valuta instellen:**
Bovenaan zie je "Set Currency". Verander de $ naar € (of een andere valuta naar keuze).

**Percentages aanpassen (optioneel):**
Wil je andere percentages dan 50/30/20? Pas de getallen aan in de gele vakjes. Zorg dat ze altijd optellen tot 100.

**Categorieën bekijken:**
Je ziet een lijst met categorieën ingedeeld per bucket (Needs / Wants / Savings). Je kunt namen aanpassen naar wat bij jouw situatie past.

---

## Stap 4 — Je eerste transactie invoeren

Ga naar het tabblad **Transactions Log**.

Vul per regel het volgende in:

| Kolom | Wat invullen | Voorbeeld |
|---|---|---|
| **Date** | Datum van de transactie | 15-01-2026 |
| **Description** | Korte omschrijving | Weekboodschappen |
| **Type** | Kies uit de lijst: Income of Expense | Expense |
| **Bucket** | Kies: Needs, Wants of Savings & Debt | Needs |
| **Category** | Kies een categorie uit de lijst | Groceries |
| **Amount** | Bedrag (altijd positief) | 85 |

> 💡 **Tip:** Gebruik de dropdown lijsten in de kolommen Type, Bucket en Category. Klik op het vakje en je ziet een pijltje verschijnen.

De kolommen **Month** en **Bucket (auto)** worden automatisch ingevuld — die hoef je niet aan te raken.

---

## Stap 5 — Het Dashboard bekijken

Ga naar het tabblad **Dashboard**.

**Maand selecteren:**
Bovenaan zie je "Selected Month" met een dropdown menu. Kies de maand die je wilt bekijken. Alle bedragen worden automatisch bijgewerkt.

**Wat je ziet:**

- **NEEDS / WANTS / SAVINGS & DEBT** (grote bedragen bovenaan): dit zijn je werkelijke uitgaven voor die maand
- **Income Summary**: al je inkomstenbronnen met verwacht vs. werkelijk bedrag
- **Needs Summary & Wants Summary**: al je uitgaven per categorie
- **Progress**: hoe ver je bent als percentage van je budget
- **Bar**: een visuele balk — groen = binnen budget, rood = over budget
- **Amount Left**: hoeveel geld je nog over hebt om uit te geven

**Kleurcodering:**
- 🟢 **Groen** = positief / binnen budget
- 🔴 **Rood** = negatief / over budget
- 🟡 **Geel** = let op, je zit bijna op je limiet

---

## Stap 6 — Verwacht bedrag instellen

De **Progress** en **Bar** kolommen werken alleen als je een verwacht bedrag invult.

Ga naar **Dashboard → Needs Summary of Wants Summary** en vul bij jouw categorieën een bedrag in onder **Expected ($)**.

**Voorbeeld Rent:**
- Expected: €800
- Actual: €725 → Progress: 91% → groene balk ✓
- Actual: €850 → Progress: 106% → rode balk ✗

---

## Veelgestelde vragen

**Mijn bedragen staan op $0,00 terwijl ik wel transacties heb ingevoerd.**
Controleer of de datum in Transactions Log correct is ingevuld én of de geselecteerde maand op het Dashboard overeenkomt met die datum.

**Ik zie #ERROR! staan in een cel.**
Druk op de cel en druk op **Cmd+R** (Mac) of **F5** (Windows) om de formule te vernieuwen. Als het probleem blijft, controleer dan of je de cel per ongeluk hebt bewerkt.

**Kan ik categorieën toevoegen?**
Ja. Ga naar **Settings** en voeg een naam toe onderaan een van de lijsten. Vergeet niet om ook de kolom **Category** in het Dashboard bij te werken als je die categorie wilt volgen.

**Kan ik meerdere jaren bijhouden?**
Ja. Maak voor elk jaar een nieuwe kopie van het template via **Bestand → Kopie maken**.

**De voortgangsbalk verschijnt niet.**
Je moet eerst een bedrag invullen bij **Expected** voor die categorie. Zonder verwacht bedrag weet het systeem niet hoe ver je bent.

---

## Tips voor starters

✅ **Voer transacties meteen in** — hoe langer je wacht, hoe makkelijker je iets vergeet

✅ **Check je Dashboard wekelijks** — 5 minuten per week is genoeg om grip te houden

✅ **Wees eerlijk met jezelf** — een kop koffie van €4,50 is ook een uitgave

✅ **Begin realistisch** — zet bij Expected bedragen in die je echt haalt, niet wat je hoopt

✅ **Gebruik Rollover** — heb je geld overgehouden? Zet dat bedrag in het Rollover-vak zodat het meerekent voor de volgende maand

---

## Technische vereisten

- Google account (gratis)
- Google Sheets (gratis, werkt in de browser)
- Geen installatie nodig

---

## Vragen of problemen?

Stuur een berichtje via Etsy. We helpen je graag verder!

---

*50/30/20 Budget Tracker • Versie 1.0 • 2026*
