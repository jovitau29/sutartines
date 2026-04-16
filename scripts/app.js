// ── Embedded template v2 (sutartis blank.docx) ───────────────────────────
const TEMPLATE = {
  "sections": [
    { "id": "header",     "label": "Sutarties antraštė" },
    { "id": "landlord",   "label": "Nuomotojo duomenys" },
    { "id": "tenant",     "label": "Nuomininko duomenys" },
    { "id": "property",   "label": "Sutarties objektas" },
    { "id": "financials", "label": "Mokėjimai ir atsiskaitymai" },
    { "id": "dates",      "label": "Sutarties terminas" },
    { "id": "act",        "label": "Perdavimo–priėmimo aktas" }
  ],
  "fields": [
    // ── Header ──────────────────────────────────────────────────────────
    {
      "id": "contractDate", "section": "header", "type": "date",
      "label": "Sutarties data", "required": true,
      "ph": "_____ m. _____ mėn. _____ d."
    },
    {
      "id": "contractCity", "section": "header", "type": "text",
      "label": "Miestas", "required": true,
      "ph": "_____________", "placeholder": "Vilnius"
    },
    // ── Landlord ────────────────────────────────────────────────────────
    {
      "id": "landlordName", "section": "landlord", "type": "text",
      "label": "Nuomotojo vardas, pavardė", "required": true,
      "ph": "_________________________", "placeholder": "Vardas Pavardė"
    },
    {
      "id": "landlordAddress", "section": "landlord", "type": "text",
      "label": "Nuomotojo adresas", "required": true,
      "ph": "_________________________", "placeholder": "Gatvė, Nr., Miestas"
    },
    {
      "id": "landlordIdCard", "section": "landlord", "type": "text",
      "label": "Nuomotojo tapatybės kortelės nr.", "required": true,
      "ph": "____________", "placeholder": "AB123456",
      "hint": "Asmens tapatybės kortelės numeris"
    },
    {
      "id": "landlordPhone", "section": "landlord", "type": "text",
      "label": "Nuomotojo telefono numeris", "required": true,
      "ph": "____________", "placeholder": "+37060000000"
    },
    {
      "id": "landlordEmail", "section": "landlord", "type": "text",
      "label": "Nuomotojo el. paštas", "required": true,
      "ph": "____________", "placeholder": "vardas@gmail.com",
      "hint": "Naudojamas susisiekimui ir komunalinių mokesčių ataskaitoms"
    },
    {
      "id": "landlordIban", "section": "landlord", "type": "text",
      "label": "Nuomotojo IBAN (nuomos mokesčiui)", "required": true,
      "ph": "LT__ ____ ____ ____ ____", "placeholder": "LT12 4010 0400 0123 4567",
      "hint": "Nuomos mokestis pervedamas į šią sąskaitą"
    },
    // ── Tenant ──────────────────────────────────────────────────────────
    {
      "id": "tenantName", "section": "tenant", "type": "text",
      "label": "Nuomininko vardas, pavardė", "required": true,
      "ph": "_________________________", "placeholder": "Vardas Pavardė"
    },
    {
      "id": "tenantId", "section": "tenant", "type": "text",
      "label": "Nuomininko asmens kodas", "required": true,
      "ph": "___________", "placeholder": "XXXXXXXXXXX", "hint": "11 skaitmenų"
    },
    {
      "id": "tenantAddress", "section": "tenant", "type": "text",
      "label": "Nuomininko adresas", "required": true,
      "ph": "_________________________", "placeholder": "Gatvė, Nr., Miestas"
    },
    {
      "id": "tenantPhone", "section": "tenant", "type": "text",
      "label": "Nuomininko telefono numeris", "required": true,
      "ph": "____________", "placeholder": "+37060000000"
    },
    {
      "id": "tenantEmail", "section": "tenant", "type": "text",
      "label": "Nuomininko el. paštas", "required": true,
      "ph": "____________", "placeholder": "vardas@gmail.com",
      "hint": "Komunalinių paslaugų ataskaitos siunčiamos šiuo adresu"
    },
    // ── Property ────────────────────────────────────────────────────────
    {
      "id": "propertyRooms", "section": "property", "type": "text",
      "label": "Kambarių skaičius (žodžiais)", "required": true,
      "ph": "____________", "placeholder": "Dviejų kambarių",
      "hint": "Pvz.: Vieno kambario, dviejų kambarių, trijų kambarių"
    },
    {
      "id": "propertyArea", "section": "property", "type": "number",
      "label": "Plotas (kv. m.)", "required": true,
      "ph": "_____ kv. m.", "placeholder": "0", "min": 1, "max": 9999
    },
    {
      "id": "propertyAddress", "section": "property", "type": "text",
      "label": "Buto adresas", "required": true,
      "ph": "_________________________", "placeholder": "Gatvė, Nr.-Butas, Miestas",
      "hint": "Pilnas adresas įskaitant buto numerį"
    },
    // ── Financials ──────────────────────────────────────────────────────
    {
      "id": "rentAmount", "section": "financials", "type": "number",
      "label": "Mėnesinis nuomos mokestis (EUR)", "required": true,
      "ph": "_____", "placeholder": "0.00", "step": 0.01
    },
    {
      "id": "rentPaymentDay", "section": "financials", "type": "number",
      "label": "Nuomos mokėjimo diena (kiekvieno mėnesio)", "required": true,
      "ph": "_____", "placeholder": "10", "min": 1, "max": 28,
      "hint": "Iki šios dienos turi būti sumokėtas nuomos mokestis"
    },
    {
      "id": "depositAmount", "section": "financials", "type": "number",
      "label": "Užstatas (EUR)", "required": true,
      "ph": "_____", "placeholder": "0.00", "step": 0.01,
      "hint": "Grąžinamas pasibaigus sutarčiai, jei nėra pretenzijų"
    },
    // ── Dates ───────────────────────────────────────────────────────────
    {
      "id": "transferDate", "section": "dates", "type": "date",
      "label": "Buto perdavimo terminas", "required": true,
      "ph": "_____________________",
      "hint": "Nuomotojas privalo perduoti butą ne vėliau kaip iki šios datos"
    },
    {
      "id": "leaseEndDate", "section": "dates", "type": "date",
      "label": "Nuomos termino pabaiga", "required": true,
      "ph": "_____________________",
      "hint": "Pasibaigus šiam terminui, sutartis automatiškai pasibaigia"
    },
    // ── Act ─────────────────────────────────────────────────────────────
    {
      "id": "actDate", "section": "act", "type": "date",
      "label": "Akto sudarymo data", "required": true,
      "ph": "_____ m. _____ mėn. _____ d.",
      "hint": "Data, kada faktiškai pasirašomas perdavimo–priėmimo aktas"
    },
    {
      "id": "meterElectricity", "section": "act", "type": "text",
      "label": "Elektros energijos skaitiklio rodmuo", "required": false,
      "ph": "_________________________", "placeholder": "0000.00 kWh"
    },
    {
      "id": "meterColdWater", "section": "act", "type": "text",
      "label": "Šalto vandens skaitiklio rodmuo", "required": false,
      "ph": "_________________________", "placeholder": "0000.000 m³"
    },
    {
      "id": "meterHotWater", "section": "act", "type": "text",
      "label": "Karšto vandens skaitiklio rodmuo", "required": false,
      "ph": "_________________________", "placeholder": "0000.000 m³"
    },
    {
      "id": "meterGas", "section": "act", "type": "text",
      "label": "Dujų skaitiklio rodmuo", "required": false,
      "ph": "_________________________", "placeholder": "0000.000 m³"
    }
  ],
  "document": [
    "NEKILNOJAMOJO TURTO (BUTO) NUOMOS SUTARTIS",
    "",
    "{{year:contractDate}} m. {{monthDay:contractDate}}",
    "{{contractCity}}",
    "",
    "{{landlordName}}, {{landlordAddress}}, asmens tapatybės kortelės Nr. {{landlordIdCard}} (toliau \u201eNuomotojas\u201c),",
    "",
    "ir",
    "",
    "{{tenantName}}, asmens kodas {{tenantId}} (toliau \u201eNuomininkas\u201c), toliau kartu vadinami \u201eŠalimis\u201c,",
    "",
    "sudaro šią Nekilnojamojo turto nuomos sutartį (toliau \u201eSutartis\u201c):",
    "",
    "1. Sutarties dalykas.",
    "",
    "1.1. Šia Sutartimi Nuomotojas įsipareigoja išnuomoti Nuomininkui Nekilnojamąjį turtą, kaip jis apibrėžtas Sutarties 1.2. punkte, už Sutartyje nustatytą mokestį, o Nuomininkas įsipareigoja mokėti nuomos mokestį.",
    "",
    "1.2. Nuomojamas Nekilnojamasis turtas pagal šią Sutartį yra:",
    "{{propertyRooms}}, {{propertyArea}} kv. metrų bendro ploto butas adresu {{propertyAddress}}.",
    "",
    "1.3. Vieno mėnesio nuomos mokestis yra {{rentAmount}} Eur. Mokestis mokamas avansu, iki einamojo mėnesio {{rentPaymentDay}} d., į 5.2 punkte nurodytą sąskaitą.",
    "",
    "1.4. Nekilnojamojo turto išnuomojimo tikslas yra gyventi.",
    "",
    "1.5. Nuomotojas garantuoja, kad turi visas reikiamas teises sudaryti šiai Sutarčiai ir išnuomoti 1.2 punkte nurodytą Nekilnojamąjį turtą Nuomininkui, taip pat kad turtas nėra niekam perleistas, išnuomotas, dėl jo nėra jokių įsipareigojimų, jis nėra apsunkintas areštu, nėra jokių disponavimo, valdymo ar naudojimo apribojimų, dėl nekilnojamojo turto nevyksta jokie ginčai, dėl ko Nuomininkas negalėtų tinkamai įgyvendinti pagal šią Sutartį suteikiamų teisių.",
    "",
    "2. Nekilnojamojo turto įrengimo ir perdavimo sąlygos.",
    "",
    "2.1. Nuomotojas iki {{date:transferDate}} privalo perduoti Nekilnojamąjį turtą Nuomininkui tvarkingą bei tinkamą naudoti pagal jo paskirtį.",
    "",
    "2.2. Nekilnojamasis turtas perduodamas pagal Perdavimo–priėmimo aktą, kurį pasirašo šalių įgalioti asmenys. Akte įvertinama Nekilnojamojo turto būklė (Perdavimo–priėmimo aktas yra šios Sutarties Priedas Nr. 1).",
    "",
    "2.3. Nuomos mokestis pradedamas skaičiuoti nuo perdavimo–priėmimo akto pasirašymo, nebent susitariama kitaip.",
    "",
    "3. Nuomos terminas.",
    "",
    "3.1. Nekilnojamasis turtas išnuomojamas fiksuotam laikotarpiui iki {{date:leaseEndDate}}. Pasibaigus šiam terminui, Sutartis pasibaigia, jeigu šalys raštu nesusitaria dėl jos pratęsimo. Nuomininkas įsipareigoja sutarties pabaigos dieną perduoti butą Nuomotojui laisvą nuo asmeninių daiktų ir tinkamos būklės (tvarkingą, nereikalaujantį kapitalinio valymo ar remonto, su nesugadintais nuomininko baldais ir technika).",
    "",
    "3.2. Nuomininkas, norėdamas pratęsti nuomos terminą, privalo ne vėliau kaip prieš 30 (trisdešimt) kalendorinių dienų iki sutarties termino pabaigos pateikti Nuomotojui rašytinį prašymą. Sutartis pratęsiama tik raštišku šalių susitarimu.",
    "",
    "3.3. Nuomininkas turi teisę nutraukti sutartį anksčiau termino, apie tai raštu įspėjęs Nuomotoją prieš 60 (šešiasdešimt) kalendorinių dienų. Tokiu atveju Nuomotojui lieka Nuomininko sumokėtas užstatas kaip kompensacija už sutarties nutraukimą anksčiau laiko, nebent šalys raštu susitaria kitaip.",
    "",
    "3.4. Nuomotojas turi teisę vienašališkai nutraukti sutartį prieš terminą, jeigu Nuomininkas esmingai pažeidžia sutarties sąlygas (daugiau nei 14 (keturiolika) kalendorinių dienų vėluoja mokėti nuomą ar komunalinius mokesčius, sugadina turtą, naudoja butą ne pagal paskirtį). Apie nutraukimą Nuomotojas raštu įspėja Nuomininką prieš 30 (trisdešimt) kalendorinių dienų. Tokiu atveju Nuomotojui taip pat lieka užstatas.",
    "",
    "3.5. Sutartis gali būti nutraukta ir kitais Lietuvos Respublikos civiliniame kodekse numatytais pagrindais.",
    "",
    "3.6. Nuomos pasibaigimo dieną Nuomininkas privalo perduoti Nekilnojamąjį turtą Nuomotojui pagal Perdavimo–priėmimo aktą, pasirašomą abiejų šalių, kuriame turi būti įvertinta išnuomotų patalpų būklė.",
    "",
    "4. Šalių pareigos.",
    "",
    "4.1. Nuomotojas įsipareigoja:",
    "",
    "    a) Užtikrinti, kad Nuomininkui būtų tinkamai tiekiamos paslaugos (elektra, vanduo, patalpų šildymas, kanalizacija, bendro naudojimo patalpų priežiūra ir kt.), reikalingos tinkamam Nekilnojamojo turto naudojimui pagal Sutartį, nebent susitariama kitaip ir už tai atsako Nuomininkas.",
    "    b) Savo sąskaita pašalinti sistemų gedimus tuo atveju, jei gedimai atsirado ne dėl Nuomininko kaltės. Tuo atveju, jei gedimai įvyko dėl netinkamo Nuomininko naudojimosi Nekilnojamuoju turtu, išlaidas kompensuoja Nuomininkas.",
    "    c) Nuomotojas turi teisę apžiūrėti Nekilnojamąjį turtą, prieš tai įspėjęs Nuomininką.",
    "    d) Vykdyti kitus šia Sutartimi prisiimtus įsipareigojimus.",
    "",
    "4.2. Nuomininkas įsipareigoja:",
    "",
    "    a) Naudoti Nekilnojamąjį turtą pagal paskirtį, prižiūrėti, naudotis juo nepakeičiant kokybės, išskyrus normalų nusidėvėjimą, griežtai laikytis taikomų priešgaisrinės apsaugos, sanitarinių bei kitų su Nekilnojamojo turto eksploatavimu susijusių taisyklių.",
    "    b) Nekilnojamojo turto neperleisti tretiesiems asmenims.",
    "    c) Be išankstinio rašytinio suderinimo su Nuomotoju neatlikti jokių Patalpų pertvarkymų, rekonstrukcijos, remonto, pagerinimo ir kitokių darbų. Darbus ir kompensaciją už juos derinti el. paštu {{landlordEmail}}.",
    "    d) Nedelsiant informuoti Nuomotoją apie įvykusias avarijas, gaisrą, kitus įvykius, kurie galėjo padaryti žalą Nekilnojamajam turtui ir imtis priemonių patalpų ir jose esančio turto išsaugojimui bei avarijos ar kito įvykio pasekmėms likviduoti.",
    "    e) Tinkamai laiku mokėti nuompinigius ir kitus mokesčius.",
    "    f) Bute nerūkyti.",
    "    g) Nelaikyti bute gyvūnų raštu nesutarus su Nuomotoju.",
    "    h) Nedeklaruoti gyvenamosios vietos buto adresu raštu nesutarus su Nuomotoju.",
    "",
    "5. Mokėjimai ir atsiskaitymai pagal Sutartį.",
    "",
    "5.1. Nuomininkas įsipareigoja mokėti nuomos mokestį, nurodytą šioje Sutartyje aukščiau.",
    "",
    "5.2. Nuomininkas įsipareigoja sumokėti {{depositAmount}} eurų dydžio užstatą. Sutarčiai pasibaigus užstatas bus grąžinamas, jei buto Nuomotojas neturės jokių pretenzijų dėl gyvenamųjų patalpų būklės ir jeigu Nuomininkas bus atsiskaitęs už einamojo mėnesio komunalines paslaugas bei buto nuomą ir neturės jokių įsiskolinimų.",
    "",
    "5.3. Visi nuomininko mokami mokėjimai, išskyrus įmokas už komunalinius mokesčius, turi būti mokami į žemiau nurodytą nuomotojo banko sąskaitą: {{landlordIban}}.",
    "",
    "5.4. Komunaliniai mokesčiai už praėjusio mėnesio komunalines paslaugas mokami atskiru pavedimu. Kas mėnesį, einamojo mėnesio 15 dieną, Nuomininkas įsipareigoja el. paštu ({{landlordEmail}}) informuoti Nuomotoją apie faktinius elektros, dujų bei karšto ir šalto vandens suvartotus kiekius. Nuomotojas kas mėnesį parengia ataskaitą su praėjusio mėnesio komunalinių paslaugų kainomis bei mokėjimo nurodymo duomenimis ir atsiunčia ją Nuomininkui el. paštu ({{tenantEmail}}).",
    "",
    "5.5. Komunaliniai mokesčiai turi būti mokami pagal kas mėnesį suformuotą mokėjimo nurodymą iki einamojo mėnesio 25 d. Taikomas 3 EUR mėnesinis komunalinių paslaugų administravimo mokestis.",
    "",
    "5.6. Nuomotojas įsipareigoja apmokėti komunalinių mokesčių dalį, susidarančią dėl pastato rekonstravimo, renovacijos bei modernizavimo projektų įgyvendinimo.",
    "",
    "6. Sankcijos ir atsakomybė.",
    "",
    "6.1. Nuomininkas atsako už žalą, padarytą Nekilnojamajam turtui, išskyrus natūralų turto nusidėvėjimą.",
    "",
    "6.2. Šalys nėra finansiškai atsakingos už kokių nors įsipareigojimų nevykdymą, jeigu sugeba įrodyti nenugalimos jėgos (force majeure) aplinkybes, vadovaujantis Lietuvos Respublikos civilinio kodekso nuostatomis.",
    "",
    "7. Sutarties galiojimas, pakeitimas ir nutraukimas.",
    "",
    "7.1. Ši Sutartis įsigalioja nuo šios sutarties ir Nekilnojamojo turto priėmimo–perdavimo akto pasirašymo dienos.",
    "",
    "7.2. Visi šios Sutarties pakeitimai, papildymai ir priedai galioja, jei jie yra sudaryti raštu ir pasirašyti abiejų šalių.",
    "",
    "8. Kitos nuostatos.",
    "",
    "8.1. Šalys įsipareigoja neatskleisti jokių šios Sutarties sąlygų jokiems tretiesiems asmenims, išskyrus valstybines institucijas, kurios pagal įstatymus turi teisę gauti tokią informaciją.",
    "",
    "8.2. Bet koks ginčas, kylantis iš šios Sutarties ar susijęs su ja, kuris per 30 dienų nuo vienos šalies pareikšto reikalavimo dėl šios Sutarties įsipareigojimų vykdymo neišsprendžiamas derybų būdu, turi būti sprendžiamas Lietuvos Respublikos įstatymų nustatyta tvarka.",
    "",
    "8.3. Šiai Sutarčiai yra taikomi ir ji turi būti aiškinama pagal Lietuvos Respublikos įstatymus.",
    "",
    "8.4. Šalių pranešimai pagal Sutartį siunčiami elektroniniu pašto adresu ar registruotu paštu adresais, nurodytais rekvizituose. Šalys privalo informuoti viena kitą apie jų duomenų pasikeitimą 5 darbo dienas prieš jiems pasikeičiant.",
    "",
    "8.5. Ši Sutartis sudaryta dviem egzemplioriais po vieną kiekvienai šaliai.",
    "",
    "Sutarties priedai:",
    "    1 Priedas: Priėmimo–perdavimo aktas.",
    "",
    "[keep-together]",
    "ŠALIŲ KONTAKTINIAI DUOMENYS IR PARAŠAI",
    "",
    "Nuomotojas: | Nuomininkas:",
    "{{landlordName}} | {{tenantName}}",
    "Adresas: {{landlordAddress}} | Adresas: {{tenantAddress}}",
    "Tel.: {{landlordPhone}} | Tel.: {{tenantPhone}}",
    "El. paštas: {{landlordEmail}} | El. paštas: {{tenantEmail}}",
    "",
    "Parašas: ____________________ | Parašas: ____________________",
    "[/keep-together]"
  ],
  "parts": [{
    "id": "part-1",
    "label": "Priedas",
    "sections": ["act"],
    "lines": [
    "1 PRIEDAS",
    "PERDAVIMO–PRIĖMIMO AKTAS",
    "prie NEKILNOJAMOJO TURTO (BUTO) NUOMOS SUTARTIES",
    "",
    "{{year:actDate}} m. {{monthDay:actDate}}",
    "{{contractCity}}",
    "",
    "{{landlordName}}, {{landlordAddress}}, asmens tapatybės kortelės Nr. {{landlordIdCard}} (toliau \u201eNuomotojas\u201c),",
    "",
    "ir",
    "",
    "{{tenantName}}, asmens kodas {{tenantId}} (toliau \u201eNuomininkas\u201c), toliau kartu vadinami \u201eŠalimis\u201c,",
    "",
    "sudarė šį Nekilnojamojo turto perdavimo–priėmimo aktą:",
    "",
    "1. Šalys šiuo aktu patvirtina, kad Nuomotojas perduoda Nuomininkui nuomos teise naudotis, o Nuomininkas priima Nekilnojamąjį turtą (butą adresu {{propertyAddress}}) pagal Nekilnojamojo turto (buto) nuomos sutartį.",
    "",
    "2. Nekilnojamasis turtas yra laikomas perduotu {{date:actDate}}.",
    "",
    "3. Nuomininkas patvirtina, kad jis yra susipažinęs su Nekilnojamojo turto būkle ir jį priima tokios būklės, koks yra šio Perdavimo–priėmimo akto sudarymo dieną.",
    "",
    "4. Perduodamose patalpose esančių skaitiklių parodymai šio perdavimo–priėmimo akto pasirašymo metu yra tokie:",
    "",
    "    Elektros energijos: :: {{meterElectricity}}",
    "    Šalto vandens: :: {{meterColdWater}}",
    "    Karšto vandens: :: {{meterHotWater}}",
    "    Dujų: :: {{meterGas}}",
    "",
    "5. Perdavimo–priėmimo aktas sudaromas 2 egzemplioriais: vienas – Nuomotojui, kitas – Nuomininkui, ir įsigalioja nuo jo pasirašymo momento.",
    "",
    "ŠALIŲ KONTAKTINIAI DUOMENYS IR PARAŠAI",
    "",
    "Nuomotojas: | Nuomininkas:",
    "{{landlordName}} | {{tenantName}}",
    "Adresas: {{landlordAddress}} | Adresas: {{tenantAddress}}",
    "Tel.: {{landlordPhone}} | Tel.: {{tenantPhone}}",
    "El. paštas: {{landlordEmail}} | El. paštas: {{tenantEmail}}",
    "",
    "Parašas: ____________________ | Parašas: ____________________"
    ]
  }]
};

const DEFAULT_DATA = {"contractDate":"2026-05-01","contractCity":"Erebonas","landlordName":"Cepelimantas Kalėdauskas","landlordAddress":"Pramanų g. 12-3, Erebonas","landlordIdCard":"ZZ000001","landlordPhone":"+37069900001","landlordEmail":"lordas.landlordas@example.com","landlordIban":"LT00 0000 0000 0000 0000","tenantName":"Kruopa Nuomininkaitė","tenantId":"39901010000","tenantAddress":"Svajonių al. 7-4, Erebonas","tenantPhone":"+37069900002","tenantEmail":"nuominas.nuominaitis@example.com","propertyRooms":"Pi kambarių","propertyArea":52,"propertyAddress":"Debesų g. 14-5, Erebonas","rentAmount":700,"rentPaymentDay":10,"depositAmount":700,"transferDate":"2026-05-15","leaseEndDate":"2027-05-14","actDate":"2026-05-15","meterElectricity":"","meterColdWater":"","meterHotWater":"","meterGas":""};

const SECTION_COLORS = {
  header: '#dbeafe',
  landlord: '#ede9fe',
  tenant: '#dcfce7',
  property: '#ffedd5',
  financials: '#fef9c3',
  dates: '#fce7f3',
  act: '#ccfbf1'
};

const TOKEN_RE = /\{\{([^}]+)\}\}/g;
const DATE_TOKEN_RE = /^(date|year|monthDay):(.+)$/;
const LINE_TYPE = {
  DIRECTIVE: 'directive',
  EMPTY: 'empty',
  COLUMNS: 'columns',
  LABELED: 'labeled',
  TEXT: 'text'
};

const MAX_PARTS = 4;

const state = {
  data: { ...DEFAULT_DATA },
  currentView: 'document',
  liveMode: false,
  liveKey: 'document',
  dragPayload: null,
  pageCounterStyle: null,
  lastLiveRange: null
};

const refs = {
  form: document.getElementById('lease-form'),
  liveFields: document.getElementById('live-fields'),
  panelBody: document.getElementById('panel-body'),
  liveEditor: document.getElementById('live-editor'),
  panelNote: document.getElementById('panel-note'),
  paper: document.querySelector('.paper'),
  documentText: document.getElementById('document-text'),
  previewTabGroup: document.getElementById('preview-tab-group'),
  btnLoad: document.getElementById('btn-load'),
  btnSave: document.getElementById('btn-save'),
  loadInput: document.getElementById('load-input'),
  loadTemplateInput: document.getElementById('load-template-input'),
  btnPrint: document.getElementById('btn-print'),
  btnDocx: document.getElementById('btn-docx'),
  mobileFieldsToggle: document.getElementById('mobile-fields-toggle'),
  mobileFieldsBackdrop: document.getElementById('mobile-fields-backdrop'),
  mobileFieldsSheet: document.getElementById('mobile-fields-sheet'),
  mobileFieldsClose: document.getElementById('mobile-fields-close'),
  mobileSheetTitle: document.getElementById('mobile-sheet-title'),
  mobileSheetContent: document.getElementById('mobile-sheet-content')
};

function normalizeTemplateShape(template) {
  const normalizedParts = Array.isArray(template.parts)
    ? template.parts.map((part, index) => ({
        id: part.id ?? `part-${index + 1}`,
        label: part.label ?? `Priedas ${index + 1}`,
        lines: Array.isArray(part.lines)
          ? part.lines
          : Array.isArray(part.document)
            ? part.document
            : [],
        sections: Array.isArray(part.sections) ? part.sections : ['act']
      }))
    : Array.isArray(template.documentAct)
      ? [{
          id: 'part-1',
          label: 'Priedas',
          lines: template.documentAct,
          sections: ['act']
        }]
      : [];

  template.document = Array.isArray(template.document) ? template.document : [];
  template.parts = normalizedParts.slice(0, MAX_PARTS);
  delete template.documentAct;
  return template;
}

normalizeTemplateShape(TEMPLATE);
const actSection = TEMPLATE.sections.find(section => section.id === 'act');
if (actSection) actSection.label = 'Priedai';

const fieldById = new Map();
const fieldsBySection = {};
for (const section of TEMPLATE.sections) fieldsBySection[section.id] = [];
for (const field of TEMPLATE.fields) {
  fieldById.set(field.id, field);
  (fieldsBySection[field.section] ||= []).push(field);
}

function currentPart() {
  return TEMPLATE.parts.find(part => part.id === state.currentView) ?? null;
}

function isMainDocumentView(view = state.currentView) {
  return view === 'document';
}

function isValidView(view) {
  return view === 'document' || TEMPLATE.parts.some(part => part.id === view);
}

function ensureCurrentView() {
  if (!isValidView(state.currentView)) state.currentView = 'document';
  if (!isValidView(state.liveKey)) state.liveKey = 'document';
}

function currentLines(view = state.currentView) {
  if (view === 'document') return TEMPLATE.document;
  return TEMPLATE.parts.find(part => part.id === view)?.lines ?? [];
}

function currentViewSections() {
  return isMainDocumentView() ? [] : new Set(currentPart()?.sections ?? []);
}

function isMobileLayout() {
  return window.matchMedia('(max-width: 900px)').matches;
}

function syncMobilePanelPlacement(mobileMode) {
  const target = mobileMode ? refs.mobileSheetContent : document.querySelector('.form-panel');
  if (refs.panelBody.parentElement !== target) {
    target.appendChild(refs.panelBody);
  }
}

function closeMobileFieldsSheet() {
  document.body.classList.remove('mobile-sheet-open');
}

function openMobileFieldsSheet() {
  if (!isMobileLayout()) return;
  document.body.classList.add('mobile-sheet-open');
}

function updateMobileTemplateUI() {
  const mobilePanelMode = isMobileLayout();
  document.body.classList.toggle('mobile-panel-mode', mobilePanelMode);
  refs.mobileFieldsToggle.hidden = !mobilePanelMode;
  refs.mobileFieldsBackdrop.hidden = !mobilePanelMode;
  refs.mobileFieldsSheet.hidden = !mobilePanelMode;
  refs.mobileFieldsToggle.textContent = state.liveMode ? 'Laukai' : 'Duomenys';
  refs.mobileSheetTitle.textContent = state.liveMode ? 'Laukai' : 'Duomenys';
  syncMobilePanelPlacement(mobilePanelMode);
  if (!mobilePanelMode) closeMobileFieldsSheet();
}

function partDisplayLabel(part) {
  const index = TEMPLATE.parts.findIndex(item => item.id === part.id);
  if (index === -1) return part.label ?? 'Priedas';
  if (TEMPLATE.parts.length <= 1) return part.label ?? 'Priedas';
  return `${part.label ?? 'Priedas'} ${index + 1}`;
}

function viewLabel(view, templateMode = state.liveMode) {
  if (view === 'document') {
    return templateMode ? 'Sutarties šablonas' : 'Sutartis';
  }

  const part = TEMPLATE.parts.find(item => item.id === view);
  const partLabel = part ? partDisplayLabel(part) : 'Priedas';
  return templateMode ? `${partLabel} šablonas` : partLabel;
}

function renderPreviewTabs() {
  ensureCurrentView();
  refs.previewTabGroup.innerHTML = '';

  const views = ['document', ...TEMPLATE.parts.map(part => part.id)];
  for (const view of views) {
    const button = document.createElement('button');
    button.className = 'seg-btn';
    button.dataset.view = view;
    const isActive = state.currentView === view;
    button.classList.toggle('active', isActive);

    const content = document.createElement('span');
    content.className = 'seg-btn-content';

    const label = document.createElement('span');
    label.textContent = viewLabel(view);
    content.appendChild(label);

    if (isActive && state.liveMode && view === 'document') {
      const addAction = document.createElement('button');
      addAction.type = 'button';
      addAction.className = 'seg-btn-action';
      addAction.setAttribute('aria-label', 'Pridėti priedą');
      addAction.title = 'Pridėti priedą';
      addAction.innerHTML = `
        <span class="btn-icon" aria-hidden="true">
          <svg viewBox="0 0 16 16" focusable="false">
            <path d="M8 3v10" />
            <path d="M3 8h10" />
          </svg>
        </span>
      `;
      addAction.addEventListener('click', event => {
        event.stopPropagation();
        addNewPart();
      });
      addAction.hidden = TEMPLATE.parts.length >= MAX_PARTS;
      content.appendChild(addAction);
    }

    if (isActive && state.liveMode && view !== 'document') {
      const action = document.createElement('button');
      action.type = 'button';
      action.className = 'seg-btn-action';
      action.setAttribute('aria-label', 'Dubliuoti priedą');
      action.title = 'Dubliuoti priedą';
      action.innerHTML = `
        <span class="btn-icon" aria-hidden="true">
          <svg viewBox="0 0 16 16" focusable="false">
            <path d="M5 5.5A1.5 1.5 0 0 1 6.5 4h5A1.5 1.5 0 0 1 13 5.5v5a1.5 1.5 0 0 1-1.5 1.5h-5A1.5 1.5 0 0 1 5 10.5z" />
            <path d="M3 9V4.5A1.5 1.5 0 0 1 4.5 3H9" />
          </svg>
        </span>
      `;
      action.addEventListener('click', event => {
        event.stopPropagation();
        duplicateCurrentPart();
      });
      action.hidden = TEMPLATE.parts.length >= MAX_PARTS;
      content.appendChild(action);

      const deleteAction = document.createElement('button');
      deleteAction.type = 'button';
      deleteAction.className = 'seg-btn-action seg-btn-action-danger';
      deleteAction.setAttribute('aria-label', 'Ištrinti priedą');
      deleteAction.title = 'Ištrinti priedą';
      deleteAction.innerHTML = `
        <span class="btn-icon" aria-hidden="true">
          <svg viewBox="0 0 16 16" focusable="false">
            <path d="M3.5 4.5h9" />
            <path d="M6 4.5V3.8A.8.8 0 0 1 6.8 3h2.4a.8.8 0 0 1 .8.8v.7" />
            <path d="M5 6.5v4.5M8 6.5v4.5M11 6.5v4.5" />
            <path d="M4.5 4.5l.5 8a1 1 0 0 0 1 .9h4a1 1 0 0 0 1-.9l.5-8" />
          </svg>
        </span>
      `;
      deleteAction.addEventListener('click', event => {
        event.stopPropagation();
        deleteCurrentPart();
      });
      content.appendChild(deleteAction);
    }

    button.appendChild(content);
    if (!isActive) {
      button.addEventListener('click', () => setView(view));
    }
    refs.previewTabGroup.appendChild(button);
  }
}

function sectionColor(fieldId) {
  return SECTION_COLORS[fieldById.get(fieldId)?.section] ?? '#f4f4f5';
}

function getLineMeta(line) {
  if (line === '[keep-together]' || line === '[/keep-together]') {
    return { type: LINE_TYPE.DIRECTIVE, value: line };
  }
  if (line === '') {
    return { type: LINE_TYPE.EMPTY };
  }
  if (line.includes(' | ')) {
    const [left, right] = line.split(' | ');
    return { type: LINE_TYPE.COLUMNS, left, right };
  }
  if (line.includes(' :: ')) {
    const [label, value] = line.split(' :: ');
    return { type: LINE_TYPE.LABELED, label, value };
  }

  const trimmed = line.trim();
  return {
    type: LINE_TYPE.TEXT,
    value: line,
    isSection: /^\d+\.\s+\S/.test(line) && !/^\d+\.\d/.test(line),
    isTitle: trimmed.length > 3 && trimmed === trimmed.toUpperCase() && !/[a-z]/.test(trimmed)
  };
}

function parseDateValue(value) {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function fieldPlaceholder(fieldId) {
  return fieldById.get(fieldId)?.ph ?? '___________';
}

function resolveToken(token) {
  const match = token.match(DATE_TOKEN_RE);
  const fieldId = match ? match[2] : token;
  const color = sectionColor(fieldId);
  const value = state.data[fieldId];
  const label = fieldById.get(fieldId)?.label ?? token;

  if (!match) {
    if (value === null || value === undefined || String(value).trim() === '') {
      return { text: fieldPlaceholder(fieldId), isValue: false, color, label };
    }
    return { text: String(value), isValue: true, color, label };
  }

  if (!value) {
    return {
      text: match[1] === 'year' ? '______' : fieldPlaceholder(fieldId),
      isValue: false,
      color,
      label
    };
  }

  const date = parseDateValue(value);
  if (!date) {
    return { text: String(value), isValue: true, color, label };
  }

  if (match[1] === 'date') {
    return {
      text: date.toLocaleDateString('lt-LT', { year: 'numeric', month: 'long', day: 'numeric' }),
      isValue: true,
      color,
      label
    };
  }

  if (match[1] === 'year') {
    return { text: String(date.getFullYear()), isValue: true, color, label };
  }

  return {
    text: date.toLocaleDateString('lt-LT', { month: 'long', day: 'numeric' }),
    isValue: true,
    color,
    label
  };
}

function buildPlaceholderNode(resolved) {
  const placeholder = document.createElement('span');
  placeholder.className = 'doc-placeholder';
  placeholder.style.setProperty('--placeholder-width', `${Math.min(Math.max((resolved.text?.length ?? 12) * 0.34, 4.5), 16)}em`);

  const line = document.createElement('span');
  line.className = 'doc-placeholder-line';

  const label = document.createElement('span');
  label.className = 'doc-placeholder-label';
  label.textContent = resolved.label;

  placeholder.append(line, label);
  return placeholder;
}

function buildLineNodes(rawLine) {
  const fragment = document.createDocumentFragment();
  let lastIndex = 0;

  for (const match of rawLine.matchAll(TOKEN_RE)) {
    if (match.index > lastIndex) {
      fragment.appendChild(document.createTextNode(rawLine.slice(lastIndex, match.index)));
    }

    const resolved = resolveToken(match[1]);
    if (resolved.isValue) {
      const mark = document.createElement('mark');
      mark.textContent = resolved.text;
      mark.style.background = resolved.color;
      fragment.appendChild(mark);
    } else {
      fragment.appendChild(buildPlaceholderNode(resolved));
    }

    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < rawLine.length) {
    fragment.appendChild(document.createTextNode(rawLine.slice(lastIndex)));
  }

  return fragment;
}

function buildForm() {
  refs.form.innerHTML = '';

  for (const section of TEMPLATE.sections) {
    const fields = fieldsBySection[section.id] ?? [];
    if (!fields.length) continue;

    const accent = SECTION_COLORS[section.id] ?? '#f4f4f5';
    const sectionNode = document.createElement('div');
    sectionNode.className = 'section';

    const label = document.createElement('p');
    label.className = 'section-label';
    label.textContent = section.label;
    label.style.cssText = `background:${accent};border-radius:3px;padding:2px 7px;display:inline-block;margin-bottom:14px;`;
    sectionNode.appendChild(label);

    for (const field of fields) {
      const wrapper = document.createElement('div');
      wrapper.className = `field${field.type === 'textarea' ? ' textarea' : ''}`;
      wrapper.style.cssText = `border-left:3px solid ${accent};padding-left:8px;`;

      const fieldLabel = document.createElement('label');
      fieldLabel.setAttribute('for', field.id);
      fieldLabel.innerHTML = field.label + (field.required ? ' <span class="req">*</span>' : '');
      wrapper.appendChild(fieldLabel);

      const input = field.type === 'textarea' ? document.createElement('textarea') : document.createElement('input');
      if (field.type === 'textarea') {
        input.rows = field.rows || 3;
      } else {
        input.type = field.type === 'number' ? 'number' : field.type === 'date' ? 'date' : 'text';
        if (field.min != null) input.min = field.min;
        if (field.max != null) input.max = field.max;
        if (field.step != null) input.step = field.step;
        if (field.type === 'date') {
          input.lang = 'lt-LT';
        }
      }

      input.id = field.id;
      if (field.placeholder) input.placeholder = field.placeholder;
      input.addEventListener('input', () => {
        state.data[field.id] = input.value;
        if (String(input.value).trim()) wrapper.classList.remove('field-invalid');
        render();
      });

      wrapper.appendChild(input);

      if (field.hint) {
        const hint = document.createElement('span');
        hint.className = 'hint';
        hint.textContent = field.hint;
        wrapper.appendChild(hint);
      }

      sectionNode.appendChild(wrapper);
    }

    refs.form.appendChild(sectionNode);
  }
}

function populateForm() {
  for (const field of TEMPLATE.fields) {
    const input = document.getElementById(field.id);
    if (!input) continue;
    input.value = state.data[field.id] ?? '';
  }
  render();
}

function renderDocument() {
  refs.documentText.innerHTML = '';
  let target = refs.documentText;

  for (const line of currentLines()) {
    const meta = getLineMeta(line);

    if (meta.type === LINE_TYPE.DIRECTIVE) {
      if (meta.value === '[keep-together]') {
        const wrapper = document.createElement('div');
        wrapper.className = 'keep-together';
        refs.documentText.appendChild(wrapper);
        target = wrapper;
      } else {
        target = refs.documentText;
      }
      continue;
    }

    if (meta.type === LINE_TYPE.EMPTY) continue;

    if (meta.type === LINE_TYPE.COLUMNS) {
      const row = document.createElement('div');
      row.className = 'doc-col';
      const left = document.createElement('span');
      const right = document.createElement('span');
      left.appendChild(buildLineNodes(meta.left));
      right.appendChild(buildLineNodes(meta.right));
      row.append(left, right);
      target.appendChild(row);
      continue;
    }

    if (meta.type === LINE_TYPE.LABELED) {
      const row = document.createElement('div');
      row.className = 'doc-labeled';
      const label = document.createElement('span');
      label.className = 'doc-lbl';
      label.appendChild(buildLineNodes(meta.label));
      const value = document.createElement('span');
      value.appendChild(buildLineNodes(meta.value));
      row.append(label, value);
      target.appendChild(row);
      continue;
    }

    const lineNode = document.createElement('span');
    lineNode.className = 'doc-line';
    if (meta.isSection) lineNode.classList.add('doc-section');
    if (meta.isTitle) lineNode.classList.add('doc-title');
    lineNode.appendChild(buildLineNodes(meta.value));
    target.appendChild(lineNode);
  }
}

function render() {
  renderDocument();
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function getValidationMessage(invalidFields) {
  const [firstInvalid] = invalidFields;
  if (!firstInvalid) return 'Prašome užpildyti privalomus laukus.';
  if (invalidFields.length === 1) return `Prašome užpildyti lauką „${firstInvalid.field.label}“.`;
  return `Prašome užpildyti lauką „${firstInvalid.field.label}“ ir dar ${invalidFields.length - 1} privalomus laukus.`;
}

function showToast(message, type = 'error') {
  document.querySelectorAll('.validation-toast').forEach(node => node.remove());
  const toast = document.createElement('div');
  toast.className = `validation-toast ${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);
  window.setTimeout(() => toast.remove(), 3500);
}

function validate() {
  const activeSections = currentViewSections();
  document.querySelectorAll('.field-invalid').forEach(node => node.classList.remove('field-invalid'));

  const invalidFields = [];
  for (const field of TEMPLATE.fields) {
    if (!field.required) continue;
    if (!isMainDocumentView() && !activeSections.has(field.section)) continue;
    if (isMainDocumentView() && field.section === 'act') continue;

    const input = document.getElementById(field.id);
    if (!input || String(input.value).trim() !== '') continue;

    input.closest('.field')?.classList.add('field-invalid');
    invalidFields.push({ field, el: input });
  }

  return invalidFields;
}

function ensureValid() {
  const invalid = validate();
  if (!invalid.length) return true;

  showToast(getValidationMessage(invalid));
  invalid[0].el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  invalid[0].el.focus();
  return false;
}

function nextPartId() {
  let index = TEMPLATE.parts.length + 1;
  let id = `part-${index}`;

  while (TEMPLATE.parts.some(part => part.id === id)) {
    index += 1;
    id = `part-${index}`;
  }

  return id;
}

function partSeed() {
  if (TEMPLATE.parts.length) {
    return JSON.parse(JSON.stringify(TEMPLATE.parts[TEMPLATE.parts.length - 1]));
  }

  return {
    label: 'Priedas',
    sections: ['act'],
    lines: ['NAUJAS PRIEDAS', '']
  };
}

function addNewPart() {
  if (TEMPLATE.parts.length >= MAX_PARTS) {
    showToast('Galima turėti daugiausia 4 priedus.');
    return;
  }

  const part = partSeed();
  part.id = nextPartId();
  TEMPLATE.parts.push(part);
  setView(part.id);
  showToast('Priedas pridėtas.', 'success');
}

function duplicateCurrentPart() {
  const part = currentPart();
  if (!part) return;
  if (TEMPLATE.parts.length >= MAX_PARTS) {
    showToast('Galima turėti daugiausia 4 priedus.');
    return;
  }

  const clone = JSON.parse(JSON.stringify(part));
  clone.id = nextPartId();

  const insertAt = TEMPLATE.parts.findIndex(item => item.id === part.id);
  TEMPLATE.parts.splice(insertAt + 1, 0, clone);

  setView(clone.id);
  showToast('Priedas nukopijuotas.', 'success');
}

function deleteCurrentPart() {
  const part = currentPart();
  if (!part) return;

  TEMPLATE.parts = TEMPLATE.parts.filter(item => item.id !== part.id);
  state.liveKey = 'document';
  setView('document');
  showToast('Priedas ištrintas.', 'success');
}

function updateTabActions() {
  renderPreviewTabs();

  if (state.liveMode) {
    refs.btnLoad.title = 'Įkelti šabloną';
    refs.btnLoad.setAttribute('aria-label', 'Įkelti šabloną');
    refs.btnSave.title = 'Išsaugoti šabloną';
    refs.btnSave.setAttribute('aria-label', 'Išsaugoti šabloną');
    refs.btnDocx.disabled = true;
    refs.btnPrint.disabled = true;
    refs.panelNote.textContent = 'Vilkite laukus į šabloną kairėje. Dokumento peržiūra atsinaujina pagal pakeitimus.';
    return;
  }

  refs.btnLoad.title = 'Įkelti duomenis';
  refs.btnLoad.setAttribute('aria-label', 'Įkelti duomenis');
  refs.btnSave.title = 'Išsaugoti duomenis';
  refs.btnSave.setAttribute('aria-label', 'Išsaugoti duomenis');
  refs.btnDocx.disabled = false;
  refs.btnPrint.disabled = false;
  refs.panelNote.innerHTML = 'Užpildykite privalomus laukus, pažymėtus <span class="req">*</span>. Peržiūra atsinaujina automatiškai.';
}

function readJsonFile(input, config) {
  const [file] = input.files;
  if (!file) return;

  const reader = new FileReader();
  reader.onload = event => {
    try {
      const parsed = JSON.parse(event.target.result);
      if (!config.isValid(parsed)) {
        showToast(config.invalidMessage);
        return;
      }
      config.apply(parsed);
      if (config.successMessage) showToast(config.successMessage, 'success');
    } catch {
      showToast(config.parseMessage);
    } finally {
      input.value = '';
    }
  };
  reader.readAsText(file);
}

function collectFormData() {
  const nextData = { ...state.data };
  for (const field of TEMPLATE.fields) {
    const input = document.getElementById(field.id);
    if (!input) continue;
    nextData[field.id] = field.type === 'number' ? (input.value === '' ? null : Number(input.value)) : input.value;
  }
  state.data = nextData;
  return nextData;
}

function loadJson(input) {
  const knownIds = new Set(TEMPLATE.fields.map(field => field.id));
  readJsonFile(input, {
    isValid(parsed) {
      return parsed && typeof parsed === 'object' && !Array.isArray(parsed) && Object.keys(parsed).some(key => knownIds.has(key));
    },
    apply(parsed) {
      state.data = { ...DEFAULT_DATA, ...parsed };
      populateForm();
    },
    invalidMessage: 'Netinkamas duomenų failas.',
    parseMessage: 'Neteisingas JSON failas.',
    successMessage: 'Duomenys įkelti.'
  });
}

function saveJson() {
  const blob = new Blob([JSON.stringify(collectFormData(), null, 2)], { type: 'application/json' });
  downloadBlob('data.json', blob);
  showToast('Duomenys išsaugoti.', 'success');
}

function applyTemplateUpdate(templatePatch) {
  const normalized = normalizeTemplateShape(JSON.parse(JSON.stringify(templatePatch)));
  if (Array.isArray(normalized.document)) TEMPLATE.document = normalized.document;
  TEMPLATE.parts = normalized.parts;
  ensureCurrentView();
  if (state.liveMode) buildLiveEditor();
  renderPreviewTabs();
  render();
}

function loadTemplate(input) {
  readJsonFile(input, {
    isValid(parsed) {
      return Array.isArray(parsed?.document) || Array.isArray(parsed?.documentAct) || Array.isArray(parsed?.parts);
    },
    apply(parsed) {
      applyTemplateUpdate(parsed);
    },
    invalidMessage: 'Netinkamas šablono failas.',
    parseMessage: 'Neteisingas šablono failas.',
    successMessage: 'Šablonas įkeltas.'
  });
}

function saveTemplate() {
  const blob = new Blob([JSON.stringify({ document: TEMPLATE.document, parts: TEMPLATE.parts }, null, 2)], { type: 'application/json' });
  downloadBlob('template.json', blob);
  showToast('Šablonas išsaugotas.', 'success');
}

function setPageCounterEnabled(enabled) {
  if (enabled && !state.pageCounterStyle) {
    state.pageCounterStyle = document.createElement('style');
    state.pageCounterStyle.id = 'page-counter-style';
    state.pageCounterStyle.textContent = [
      '@media print {',
      '  @page {',
      '    @bottom-right {',
      '      content: counter(page) " / " counter(pages);',
      '      font-size: 8pt;',
      '      color: #888;',
      '      font-family: "Segoe UI", system-ui, sans-serif;',
      '    }',
      '  }',
      '}'
    ].join('\n');
    document.head.appendChild(state.pageCounterStyle);
  }

  if (!enabled && state.pageCounterStyle) {
    state.pageCounterStyle.remove();
    state.pageCounterStyle = null;
  }
}

function setView(view) {
  state.currentView = isValidView(view) ? view : 'document';

  setPageCounterEnabled(isMainDocumentView());
  renderPreviewTabs();

  if (state.liveMode) {
    state.liveKey = state.currentView;
    buildLiveEditor();
  }

  updateMobileTemplateUI();
  render();
}

function setLeftTab(tab) {
  state.liveMode = tab === 'live';
  state.liveKey = state.currentView;

  document.querySelectorAll('.left-tab').forEach(button => {
    button.classList.toggle('active', button.dataset.tab === tab);
  });

  refs.form.style.display = state.liveMode ? 'none' : 'block';
  refs.liveFields.style.display = state.liveMode ? 'block' : 'none';
  refs.paper.style.display = state.liveMode ? 'none' : '';
  refs.liveEditor.style.display = state.liveMode ? 'block' : 'none';
  updateTabActions();
  updateMobileTemplateUI();
  refs.btnLoad.onclick = state.liveMode ? () => refs.loadTemplateInput.click() : () => refs.loadInput.click();
  refs.btnSave.onclick = state.liveMode ? saveTemplate : saveJson;

  if (state.liveMode) {
    buildLiveFields();
    buildLiveEditor();
  } else {
    closeMobileFieldsSheet();
  }
}

function getTokenField(token) {
  const match = token.match(DATE_TOKEN_RE);
  return fieldById.get(match ? match[2] : token);
}

function makeLiveChip(token) {
  const field = getTokenField(token);
  const prefix = token.startsWith('date:')
    ? 'data · '
    : token.startsWith('year:')
      ? 'metai · '
      : token.startsWith('monthDay:')
        ? 'mėn. d. · '
        : '';

  const chip = document.createElement('span');
  chip.className = 'tpl-chip';
  chip.setAttribute('contenteditable', 'false');
  chip.setAttribute('draggable', 'true');
  chip.dataset.token = token;
  chip.style.background = sectionColor(field?.id ?? token);
  chip.textContent = prefix + (field?.label ?? token);
  chip.addEventListener('dragstart', event => {
    state.dragPayload = { type: 'editor', token, chip };
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', token);
    event.stopPropagation();
  });
  chip.addEventListener('dragend', () => {
    state.dragPayload = null;
  });
  return chip;
}

function populateLiveSpan(element, text) {
  if (!text) {
    element.appendChild(document.createElement('br'));
    return;
  }

  let lastIndex = 0;
  for (const match of text.matchAll(TOKEN_RE)) {
    if (match.index > lastIndex) {
      element.appendChild(document.createTextNode(text.slice(lastIndex, match.index)));
    }
    element.appendChild(makeLiveChip(match[1]));
    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < text.length) {
    element.appendChild(document.createTextNode(text.slice(lastIndex)));
  }
}

function buildLiveLine(line) {
  const meta = getLineMeta(line);

  if (meta.type === LINE_TYPE.DIRECTIVE) {
    const directive = document.createElement('div');
    directive.className = 'live-directive';
    directive.setAttribute('contenteditable', 'false');
    directive.dataset.directive = meta.value;
    return directive;
  }

  if (meta.type === LINE_TYPE.EMPTY) {
    const emptyLine = document.createElement('div');
    emptyLine.appendChild(document.createElement('br'));
    return emptyLine;
  }

  if (meta.type === LINE_TYPE.COLUMNS) {
    const row = document.createElement('div');
    row.className = 'live-col';
    row.dataset.lineType = 'col';
    const left = document.createElement('span');
    left.dataset.col = 'left';
    const right = document.createElement('span');
    right.dataset.col = 'right';
    populateLiveSpan(left, meta.left);
    populateLiveSpan(right, meta.right);
    row.append(left, right);
    return row;
  }

  if (meta.type === LINE_TYPE.LABELED) {
    const row = document.createElement('div');
    row.className = 'live-labeled';
    row.dataset.lineType = 'labeled';
    const label = document.createElement('span');
    label.className = 'live-lbl';
    const value = document.createElement('span');
    populateLiveSpan(label, meta.label);
    populateLiveSpan(value, meta.value);
    row.append(label, value);
    return row;
  }

  const textLine = document.createElement('div');
  if (meta.isSection || meta.isTitle) textLine.style.fontWeight = 'bold';
  if (meta.isTitle) textLine.style.textAlign = 'center';
  populateLiveSpan(textLine, meta.value);
  return textLine;
}

function buildLiveEditor() {
  refs.liveEditor.innerHTML = '';
  for (const line of currentLines(state.liveKey)) {
    refs.liveEditor.appendChild(buildLiveLine(line));
  }
}

function serializeContent(element) {
  let text = '';
  element.childNodes.forEach(node => {
    if (node.nodeType === Node.TEXT_NODE) {
      text += node.textContent;
      return;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return;
    if (node.dataset?.token) {
      text += `{{${node.dataset.token}}}`;
      return;
    }
    if (node.nodeName !== 'BR') {
      text += serializeContent(node);
    }
  });
  return text;
}

function serializeEditorLine(element) {
  if (element.dataset?.directive) return element.dataset.directive;
  if (element.dataset?.lineType === 'col') {
    const columns = element.querySelectorAll('[data-col]');
    return `${serializeContent(columns[0] ?? element)} | ${serializeContent(columns[1] ?? element)}`;
  }
  if (element.dataset?.lineType === 'labeled') {
    const [label, value] = element.children;
    return `${serializeContent(label)} :: ${serializeContent(value)}`;
  }
  return serializeContent(element);
}

function syncFromEditor() {
  const lines = [];
  refs.liveEditor.childNodes.forEach(node => {
    if (node.nodeType === Node.TEXT_NODE) {
      node.textContent.split('\n').forEach(text => lines.push(text));
      return;
    }
    if (node.nodeName === 'DIV' || node.nodeName === 'P') {
      lines.push(serializeEditorLine(node));
      return;
    }
    if (node.nodeName === 'BR') {
      lines.push('');
    }
  });
  if (state.liveKey === 'document') {
    TEMPLATE.document = lines;
  } else {
    const part = TEMPLATE.parts.find(item => item.id === state.liveKey);
    if (part) part.lines = lines;
  }
  render();
}

function getFieldVariants(field) {
  if (field.type !== 'date') {
    return [{ token: field.id, label: field.label }];
  }

  return [
    { token: `date:${field.id}`, label: `${field.label} (pilna data)` },
    { token: `year:${field.id}`, label: `${field.label} (metai)` },
    { token: `monthDay:${field.id}`, label: `${field.label} (mėn. diena)` }
  ];
}

function insertTokenIntoEditor(token) {
  const chip = makeLiveChip(token);
  const selection = window.getSelection();
  const savedRange = state.lastLiveRange;

  if (savedRange && refs.liveEditor.contains(savedRange.startContainer)) {
    const range = savedRange.cloneRange();
    range.deleteContents();
    range.insertNode(chip);
    const cursor = document.createRange();
    cursor.setStartAfter(chip);
    cursor.collapse(true);
    selection.removeAllRanges();
    selection.addRange(cursor);
    state.lastLiveRange = cursor.cloneRange();
    syncFromEditor();
    return;
  }

  refs.liveEditor.focus();
  refs.liveEditor.appendChild(chip);
  const cursor = document.createRange();
  cursor.setStartAfter(chip);
  cursor.collapse(true);
  selection.removeAllRanges();
  selection.addRange(cursor);
  state.lastLiveRange = cursor.cloneRange();
  syncFromEditor();
}

function buildLiveFields() {
  refs.liveFields.innerHTML = '';

  for (const section of TEMPLATE.sections) {
    const fields = fieldsBySection[section.id] ?? [];
    if (!fields.length) continue;

    const group = document.createElement('div');
    group.className = 'live-group';

    const header = document.createElement('div');
    header.className = 'live-group-label';
    header.textContent = section.label;
    header.style.background = SECTION_COLORS[section.id] ?? '#f4f4f5';
    group.appendChild(header);

    for (const field of fields) {
      for (const variant of getFieldVariants(field)) {
        const chip = document.createElement('div');
        chip.className = 'live-source-chip';
        chip.setAttribute('draggable', 'true');
        chip.dataset.token = variant.token;
        chip.style.background = sectionColor(field.id);
        chip.textContent = variant.label;
        chip.title = `{{${variant.token}}}`;
        chip.addEventListener('dragstart', event => {
          state.dragPayload = { type: 'list', token: variant.token };
          event.dataTransfer.effectAllowed = 'copy';
          event.dataTransfer.setData('text/plain', variant.token);
        });
        chip.addEventListener('dragend', () => {
          state.dragPayload = null;
        });
        chip.addEventListener('click', () => {
          if (!isMobileLayout() || !state.liveMode) return;
          insertTokenIntoEditor(variant.token);
          closeMobileFieldsSheet();
        });
        group.appendChild(chip);
      }
    }

    refs.liveFields.appendChild(group);
  }
}

function createDropRange(event) {
  if (document.caretRangeFromPoint) {
    return document.caretRangeFromPoint(event.clientX, event.clientY);
  }
  if (document.caretPositionFromPoint) {
    const position = document.caretPositionFromPoint(event.clientX, event.clientY);
    if (!position) return null;
    const range = document.createRange();
    range.setStart(position.offsetNode, position.offset);
    range.collapse(true);
    return range;
  }
  return null;
}

function initLiveDragDrop() {
  const saveLiveSelection = () => {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount) return;
    const range = selection.getRangeAt(0);
    if (!refs.liveEditor.contains(range.startContainer)) return;
    state.lastLiveRange = range.cloneRange();
  };

  refs.liveEditor.addEventListener('dragover', event => {
    if (!state.dragPayload) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = state.dragPayload.type === 'editor' ? 'move' : 'copy';
  });

  refs.liveEditor.addEventListener('drop', event => {
    if (!state.dragPayload) return;
    event.preventDefault();

    const payload = state.dragPayload;
    state.dragPayload = null;
    const range = createDropRange(event);
    if (!range) return;

    const chip = makeLiveChip(payload.token);
    range.insertNode(chip);

    const selection = window.getSelection();
    const cursor = document.createRange();
    cursor.setStartAfter(chip);
    cursor.collapse(true);
    selection.removeAllRanges();
    selection.addRange(cursor);
    state.lastLiveRange = cursor.cloneRange();

    if (payload.type === 'editor' && payload.chip && payload.chip !== chip) {
      payload.chip.remove();
    }

    syncFromEditor();
  });

  refs.liveEditor.addEventListener('input', syncFromEditor);
  refs.liveEditor.addEventListener('focus', saveLiveSelection);
  refs.liveEditor.addEventListener('mouseup', saveLiveSelection);
  refs.liveEditor.addEventListener('keyup', saveLiveSelection);
  refs.liveEditor.addEventListener('touchend', saveLiveSelection);
}

function saveDocx() {
  if (!ensureValid()) return;

  const { Document, Paragraph, TextRun, Table, TableRow, TableCell, Packer, WidthType, AlignmentType } = window.docx;
  const cm = value => Math.round(value * 567);
  const NO_BORDER = { style: 'none', size: 0, color: 'FFFFFF' };
  const NO_BORDERS = {
    top: NO_BORDER,
    bottom: NO_BORDER,
    left: NO_BORDER,
    right: NO_BORDER,
    insideHorizontal: NO_BORDER,
    insideVertical: NO_BORDER
  };

  function lineRuns(rawLine, bold = false) {
    const runs = [];
    let lastIndex = 0;

    for (const match of rawLine.matchAll(TOKEN_RE)) {
      if (match.index > lastIndex) {
        runs.push(new TextRun({ text: rawLine.slice(lastIndex, match.index), bold, font: 'Times New Roman', size: 22 }));
      }
      runs.push(new TextRun({ text: resolveToken(match[1]).text, bold, font: 'Times New Roman', size: 22 }));
      lastIndex = match.index + match[0].length;
    }

    if (lastIndex < rawLine.length) {
      runs.push(new TextRun({ text: rawLine.slice(lastIndex), bold, font: 'Times New Roman', size: 22 }));
    }

    return runs.length ? runs : [new TextRun({ text: '', font: 'Times New Roman', size: 22 })];
  }

  function paragraph(rawLine, options = {}) {
    const { bold = false, alignment = AlignmentType.BOTH, spaceBefore = 0, spaceAfter = 60, keepLines = false, indent = 0 } = options;
    return new Paragraph({
      alignment,
      keepLines,
      indent: indent ? { left: indent } : undefined,
      spacing: { before: spaceBefore, after: spaceAfter, line: 240 },
      children: lineRuns(rawLine, bold)
    });
  }

  function rowTable(left, right, leftWidthPct = 50) {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: NO_BORDERS,
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: NO_BORDERS,
            width: { size: leftWidthPct, type: WidthType.PERCENTAGE },
            children: [paragraph(left)]
          }),
          new TableCell({
            borders: NO_BORDERS,
            width: { size: 100 - leftWidthPct, type: WidthType.PERCENTAGE },
            children: [paragraph(right)]
          })
        ]
      })]
    });
  }

  const children = [];
  let keepTogether = false;

  for (const line of currentLines()) {
    const meta = getLineMeta(line);

    if (meta.type === LINE_TYPE.DIRECTIVE) {
      keepTogether = meta.value === '[keep-together]';
      continue;
    }

    if (meta.type === LINE_TYPE.EMPTY) {
      children.push(new Paragraph({ spacing: { before: 0, after: 80, line: 240 } }));
      continue;
    }

    if (meta.type === LINE_TYPE.COLUMNS) {
      children.push(rowTable(meta.left, meta.right));
      continue;
    }

    if (meta.type === LINE_TYPE.LABELED) {
      children.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: NO_BORDERS,
        rows: [new TableRow({
          children: [
            new TableCell({
              borders: NO_BORDERS,
              width: { size: cm(5), type: WidthType.DXA },
              children: [paragraph(meta.label)]
            }),
            new TableCell({
              borders: NO_BORDERS,
              children: [paragraph(meta.value)]
            })
          ]
        })]
      }));
      continue;
    }

    children.push(paragraph(meta.value, {
      bold: meta.isSection || meta.isTitle,
      alignment: meta.isTitle ? AlignmentType.CENTER : AlignmentType.BOTH,
      spaceBefore: meta.isSection ? 180 : 0,
      spaceAfter: 60,
      keepLines: keepTogether
    }));
  }

  const documentFile = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: cm(2), bottom: cm(2), left: cm(3), right: cm(2) }
        }
      },
      children
    }]
  });

  const filename = isMainDocumentView()
    ? 'sutartis.docx'
    : `priedas-${TEMPLATE.parts.findIndex(part => part.id === state.currentView) + 1}.docx`;
  Packer.toBlob(documentFile).then(blob => downloadBlob(filename, blob));
}

function init() {
  buildForm();
  populateForm();
  updateTabActions();
  updateMobileTemplateUI();

  refs.btnLoad.onclick = () => refs.loadInput.click();
  refs.btnSave.onclick = saveJson;
  refs.loadInput.addEventListener('change', () => loadJson(refs.loadInput));
  refs.loadTemplateInput.addEventListener('change', () => loadTemplate(refs.loadTemplateInput));
  refs.btnPrint.addEventListener('click', () => {
    if (!ensureValid()) return;
    window.print();
  });
  refs.btnDocx.addEventListener('click', saveDocx);
  refs.mobileFieldsToggle.addEventListener('click', openMobileFieldsSheet);
  refs.mobileFieldsBackdrop.addEventListener('click', closeMobileFieldsSheet);
  refs.mobileFieldsClose.addEventListener('click', closeMobileFieldsSheet);
  window.addEventListener('resize', updateMobileTemplateUI);

  document.querySelectorAll('.left-tab').forEach(button => {
    button.addEventListener('click', () => setLeftTab(button.dataset.tab));
  });

  initLiveDragDrop();
  setView('document');
}

init();
