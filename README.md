# Dokumentanonymiserare

Ett lokalt Windows-verktyg med grafiskt gränssnitt (Tkinter) för att anonymisera och återställa dokument, ritningar och texter innan de delas eller laddas upp till t.ex. AI-tjänster. All bearbetning sker **helt lokalt på din egen dator**.

---

## Funktioner

- **Stödda filformat:** `.docx` (Word), `.pdf`, `.xlsx` (Excel), `.tif` / `.tiff` (ritningar/bilder), `.txt` och `.md`.
- **Versalbevarande ersättning:** Matchar skiftläge automatiskt (`VERSALER`, `gemener`, `Inledande versal`) för både textinnehåll och filnamn.
- **Word (`.docx`):** Bearbetar brödtext, tabeller, sidhuvud/sidfot och textrutor, rensar dokumentmetadata (författare, titel m.m.) och varnar om inbäddade bilder finns.
- **Excel (`.xlsx`):** Bearbetar cellvärden (bevarar formler som börjar med `=`), byter fliknamn och rensar metadata.
- **PDF (`.pdf`):**
  - Textbaserade PDF:er bearbetas med äkta *redaction* (originaltext tas bort ur filen, inte bara täcks över).
  - Skannade PDF-sidor (utan textlager) hanteras automatiskt som bild med ritningshuvudsmaskering och valfri OCR.
  - Metadata rensas.
- **Ritningar (`.tif` / `.tiff` & skannade PDF:er):**
  - Vitmålar ritningshuvud i nedre högra hörnet med valfri bredd/höjd i procent samt inbyggd **förhandsgranskning** (röd ram).
  - Stöd för OCR-textutbyte via Tesseract OCR (valfritt).
- **Extra automatisk maskering (envägs vid anonymisering):**
  - E-postadresser → `[E-POST]`
  - Svenska telefonnummer (`+46` / `07x`) → `[TELEFON]`
  - Koordinater (6- eller 7-siffriga N/E-format) → `[KOORDINAT]`
- **Läckageskanning:** Kontrollerar automatiskt textlagren i alla skapade filer efter anonymisering och varnar om något nyckelord finns kvar.
- **Flik för text & AI-svar:** Klistra in text eller AI-svar för att snabbt anonymisera eller återställa pseudonymiserade namn med ett klick.
- **Återställning (av-anonymisering):** Växla läge till *Återställ* för att översätta tillbaka pseudonymer till originalnamn med samma nyckellista.

---

## Snabbstart – Ladda ner färdig `.exe` (Windows)

Du behöver inte ha Python installerat för att använda programmet.

1. Gå till [**Releases**](https://github.com/Kransengangster1000/Dokumentanonymiserare/releases/latest) och ladda ner **`Anonymiserare.exe`**.
2. Placera `Anonymiserare.exe` i en valfri mapp på din dator.
3. Dubbelklicka på `Anonymiserare.exe` för att starta programmet.
   > **Obs (Windows SmartScreen):** Eftersom `.exe`-filen inte är kodsignerad kan Windows visa en varning första gången. Klicka då på *Mer information* → *Kör ändå*.
4. Vid första starten skapas automatiskt en tom/exempelbaserad `nyckel.json` i samma mapp som programmet.

---

## Användarguide

### 1. Skapa din nyckellista (Flik 3: `Nyckellista`)
1. Öppna fliken **3. Nyckellista (nyckel.json)**.
2. Fyll i **Original** (det riktiga namnet, t.ex. projektnamn, företag eller person) och **Pseudonym** (ersättningsnamn) och klicka på **Lägg till / Uppdatera**.
3. Dubbelklicka på en befintlig rad i tabellen för att redigera den, eller markera och klicka **Ta bort markerad**.
4. Du kan även hantera separata nyckelfiler per projekt via **Öppna annan nyckelfil...**.

> **Viktigt om sekretess:** Din `nyckel.json` innehåller kopplingen mellan riktiga namn och pseudonymer. Förvara den lokalt och dela/ladda aldrig upp den på internet.

### 2. Bearbeta filer (Flik 1: `Bearbeta filer`)
1. Välj **In-mapp** med originaldokumenten. En **Ut-mapp** föreslås automatiskt (`ut_anonymiserat`), men kan ändras fritt.
2. Välj läge:
   - **Anonymisera (inför uppladdning):** Byter ut original → pseudonym, rensar metadata och kör valfri extra maskering + läckageskanning.
   - **Återställ (av-anonymisera):** Byter tillbaka pseudonym → original.
3. Justera vid behov:
   - **Inkludera undermappar** för rekursiv bearbetning.
   - **Vitmåla ritningshuvud** + **Bredd %** / **Höjd %** och klicka **Förhandsgranska TIF...** för att kontrollera den röda rutan innan körning.
   - **OCR/maskering på skannade PDF:er**, **Maskera e-post / telefon / koordinater** och **Läckageskanning**.
4. Klicka på **▶ Kör bearbetning** och granska kontrollrapporten i loggfönstret (grönt = OK, orange = obs/varning, rött = fel eller läckage).

### 3. Bearbeta text & AI-svar (Flik 2: `Text & AI-svar`)
1. Klistra in ett textstycke eller ett AI-svar i det övre fältet.
2. Klicka på **Anonymisera text** eller **Återställ text**.
3. Klicka på **Kopiera resultat** för att kopiera texten till urklipp.

---

## Valfritt: Aktivera OCR på ritningar och skannade PDF:er

Programmet fungerar utan Tesseract OCR (ritningshuvud vitmålas ändå), men om du även vill att programmet ska läsa och ersätta text direkt i `.tif`-bilder och skannade PDF-sidor behöver **Tesseract OCR** vara installerat på Windows:

1. Ladda ner Tesseract OCR för Windows (t.ex. från [UB-Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)).
2. Installera till standardsökvägen:
   ```text
   C:\Program Files\Tesseract-OCR\tesseract.exe
   ```
3. Under installationen, se till att välja språkpaketen **Swedish (`swe`)** och **English (`eng`)** under *Additional language data*.

Programmet känner av installationen automatiskt vid körning.

---

## Kör från källkod (Python)

Kräver **Python 3.10 eller senare**.

```bash
git clone https://github.com/Kransengangster1000/Dokumentanonymiserare.git
cd Dokumentanonymiserare
pip install -r requirements.txt
python AA.py
```

### Bygg en egen `.exe` med PyInstaller

```bash
pip install pyinstaller
python -m PyInstaller --noconfirm Anonymiserare.spec
```

Den färdiga filen skapas i `dist/Anonymiserare.exe`.

---

## Säkerhetsnoteringar & begränsningar

- **Inbäddade bilder i Word-filer (`.docx`):** Text inuti skärmdumpar eller klistrade bilder i Word bearbetas inte automatiskt. Programmet varnar i loggen när bilder upptäcks så att du kan kontrollera dem manuellt.
- **Läckageskanning:** Kontrollerar textlager i `.docx`, `.pdf`, `.xlsx`, `.txt` och `.md`. Kontrollera alltid skannade dokument och ritningar visuellt innan delning.
- **Originalfiler:** Originalfilerna i in-mappen ändras aldrig – alla bearbetade filer skrivs till ut-mappen.
