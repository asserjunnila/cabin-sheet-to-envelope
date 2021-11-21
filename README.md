# cabin-sheet-to-envelope

## Steps 
1. Clone repo
2. Install python and requirements.txt to your machine
3. Create config.json according the schema.json (see example below)
4. Drop .jpeg and make sure that the picture size is relative to the resolution
5. Make sure all the config.json fields are filled according to the need and the spreadsheets
5. Create a folder for cleaner file management and drop all the .xlsx spreadsheets there
6. `python to_envelope.py`
7. Have a look at the results in *envelope_print/* folder and validate the results
8. Enjoy and head for the harbour

## config.json
Example config.json 
```
{
  "spreadsheetconfig": {
    "cabin_class_col": "A",
    "cabin_id_col": "B",
    "last_name_col": "C",
    "first_name_col": "D",
    "DIN1_col": "G",
    "DIN2_col": "H",
    "BRE_col": "I",
    "LUN_col": "J",
    "first_cabin_row": 3
  },
  "picture": "JR21_kirjekuoriin.jpg",
  "heading": "Rannekkeiden kontrolliosat toimivat Avajaisshow'ssa arpalippuina! Pääpalkintona arvonnassa on hulppea Suite-hytti! Katso ohjelmasta lisätiedot!",
  "first": "Laivayhtiö veloittaa asiakkailta alkaen 100€ hyteistä, jotka on sotkettu tai joissa on tupakoitu.",
  "second": "Omien ja Tax Freesta ostettujen alkoholijuomien käyttö laivalla ehdottomasti kielletty.",
  "third": "Älä laita hyttikorttiasi puhelimen suojakoteloon, sillä kotelon magneetti vaurioittaa korttia!",
  "quote": [
    "Jouluristeilyllä on ykkösvaatteet yllä",
    "Jouluristeilyllä juomaa riittää kyllä",
    "On silmänpilkettä, on vinhaa vilkettä",
    "Kun laiva keinuu vaan, kauas häipyy maa",
    "Hytti sata viis, saavu sinne pliis",
    "Tilulilulii, hytti on auki ja se ei mee kii",
    "Siis oothan oikein hyvä mulle jos annan avainkortin sulle",
    "Halutsä nähä mun risteilyohjuksen?",
    "Mikä vuosi ku Estonia uppos? - Olikos se keulavisiiri?",
    "Se oli etukamera",
    "Anna sitä märkää",
    "Jos nain vesisängyssä, tuun merisairaaksi",
    "Älä pidä hyttikorttia puhelimen vieressä",
    "Pitäkää kädet laivan sisäpuolella",
    "Kiinnittäkää turpavyöt",
    "Ei saa maata näkyvissä",
    "Jouluristely, Itämeren ainoa Sirkus",
    "Voi kauhhia ku tytöillä on jäänyt housut kotiin, T. Tuplis",
    "Kuka antais pukille?",
    "#ihanaa",
    "Kauas pilvet karkaavat",
    "Aavan meren tuolla puolen jossakin on maa",
    "Oi jospa kerran sinne satumaahan käydä vois",
    "Kostuta mun tatuointii",
    "Moisturize me!!",
    "Laiva keinuu, minä en",
    "Eiköhän se juna oo jo seilannu",
    "Joulu vailla rissee on ku kesä ilman bissee.",
    "Toiset on korkeakoulus tutkinnon takii. Mä tän risteilyn.",
    "Posti ei kulje, eikä varmaan mekään",
    "Ois kiva saada tietoa asap että risteilläänkö vai ei?",
    "Tulkoot vaikka yleislakko, tän laivan kulkea on pakko.",
    "Minä keinun, laiva ei",
    "Monelta satamassa pitää olla huomenna ettei vaan myöhästy ku se laiva lähtee?",
    "Viime vuonna oli kauhea myrsky, nyt ei satama-altaassa haittaa tyrskyt."
  ]
}
```
