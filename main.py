import os
import json
from faster_whisper import WhisperModel
from docxtpl import DocxTemplate
from docx import Document
from dotenv import load_dotenv
load_dotenv()

# =============================
# INIT
# =============================
# Detailniveau kan er nog bij bij voporbeelden in prompt
# Lengte kan ook
import streamlit as st
from openai import OpenAI
# laad .env bestand
load_dotenv()

client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY")
)


def load_whisper():
    return WhisperModel("base")

whisper_model = load_whisper()

# =============================
# HELPERS
# =============================
def read_docx(file_path):
    doc = Document(file_path)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def parse_basisdocument(text):
    items = []
    for line in text.split("\n"):
        line = line.strip()
        if line:
            items.append(line)
    return items

# =============================
# LOAD STATIC FILES
# =============================
BASISDOCUMENT = parse_basisdocument(read_docx("basis.docx"))
BASISDOCUMENT_TEXT = "\n".join(BASISDOCUMENT)

# voorbeelden tijdelijk uitgeschakeld
VOORBEELD1 = read_docx("voorbeeld1.docx")
VOORBEELD2 = read_docx("voorbeeld2.docx")
VOORBEELD3 = read_docx("voorbeeld3.docx")

# =============================
# AUDIO → TEXT
# =============================
def transcribe_audio(file_path):
    segments, _ = whisper_model.transcribe(file_path)
    return " ".join([s.text for s in segments]).strip()

# =============================
# AI → JSON OUTPUT
# =============================
def generate_json(transcript, notes=""):

    SYSTEM = f"""
    DOEL:
    Schrijf een volledig, praktisch en professioneel consultverslag voor de cliënt.
    Het verslag moet voelen als een echt therapeutisch adviesdocument dat direct gebruikt kan worden.

    De cliënt moet alles begrijpen zonder extra uitleg:
    - wat er speelt in het lichaam en de klachten
    - wat de oorzaak en samenhang is
    - wat er concreet moet gebeuren
    - hoe supplementen gebruikt moeten worden (stap voor stap)
    - wat voeding/levensstijl doet met de klachten
    - welke vervolgstappen of onderzoeken nodig zijn

    ====================

    SCHRIJFSTIJL (BELANGRIJKSTE REGEL):

    Schrijf in een mix van:
    - korte alinea’s
    - en bullets

    MAAR:
    - elke bullet is een mini-uitleg (niet 1 losse zin)
    - bullets zijn verhalend, niet droog
    - elke bullet bevat context + uitleg + situatie + eventueel gevolg

    Dus NIET:
    • pijn in knie

    MAAR WEL:
    • Je ervaart pijn in de knie bij langere belasting zoals lopen of traplopen, waarbij de belasting duidelijk toeneemt na 10–15 minuten en herstel nodig is na rust.

    ====================

    STRUCTUUR (MOET HIER OP LIJKEN):

    - Bovenaan: datum + naam + volgende consult
    - Daarna: klachten en huidige situatie (meest uitgebreid)
    - Daarna: voeding
    - Daarna: supplementen uitleg (uitgebreid + stap-voor-stap)
    - Daarna: onderzoeken / therapeuten
    - Daarna: overzicht supplementen schema

    ====================

    SCHRIJFREGELS:

    - Concreet, duidelijk en professioneel
    - Geen algemene adviezen zonder context
    - Geen vage woorden zoals “gezond”, “beter”, “goed bezig”
    - Altijd oorzaak → gevolg → impact op dagelijks leven
    - Nooit samenvatten als details beschikbaar zijn
    - Werk elk onderwerp volledig uit alsof een nieuwe therapeut alles moet begrijpen

    - Spreek de cliënt aan met “u”
    - Nooit “je” of “jij”
    - Formeel Nederlands

    - Elke klacht bevat:
    frequentie, ernst, triggers, belasting, en herstel

    - Elke supplementbeschrijving bevat:
    werking + reden + timing + opbouw + wat te verwachten

    ====================

    BELANGRIJKE REGELS (HARD):

    - Geef ALTIJD geldige JSON
    - Geen tekst buiten JSON
    - Geen uitleg
    - Geen hallucinaties
    - Geen bronverwijzingen
    - Geen woorden zoals: “volgens”, “uit transcript”, “advies”

    - Gebruik transcript + notities als enige bron
    - Gebruik BASISDOCUMENT alleen voor supplementdetails
    - Voeg nooit nieuwe supplementen toe

    ====================

    BULLETS REGEL (ZEER BELANGRIJK):

    - "huidige_situatie" = verhalende bullets (1–3 zinnen per bullet)
    - elke bullet begint met "•"
    - elke bullet = 1 onderwerp met context
    - geen losse woorden of korte zinnen

    Andere secties:
    - voeding_verminderen
    - voeding_verhogen
    - onderzoeken
    - therapeuten

    ====================

    SUPPLEMENT REGELS:

    - Werk supplementen stap voor stap uit
    - Elke stap = aparte bullet
    - Geen dubbele info in 1 bullet
    - Opbouwschema is verplicht als het in transcript staat

    - Als “opbouw” of dosering verandert:
    → maak dit een logische volgorde (week 1 → week 2 → etc.)

    - Alleen relevante details uit BASISDOCUMENT gebruiken
    - Altijd prijs vermelden indien aanwezig

    ====================

    MOMENTEN LOGICA:

    - "voor ontbijt" → voor_ontbijt = true
    - "bij ontbijt" → ontbijt = true
    - nooit beide tegelijk tenzij expliciet genoemd

    ====================

    ALS INFORMATIE ONTBREEKT:

    - gebruik "Onbekend"
    - NOOIT "NVT"
    - ontbrekende supplementvelden weglaten (niet invullen)

    ====================

    CONTROLE:

    Controleer vóór output:
    - alle klachten verwerkt
    - alle supplementen verwerkt
    - alle acties verwerkt
    - geen samenvattingen waar details zijn

    ====================

    VOORBEELD STIJL (ZEER BELANGRIJK):

    Gebruik deze stijl als leidraad:

    • Je ervaart toename van klachten bij belasting zoals wandelen en traplopen, waarbij de pijn vooral optreedt na enkele minuten en daarna langzaam afneemt in rust.
    • De energie is redelijk stabiel, maar wordt duidelijk beïnvloed door werkdruk en mentale belasting gedurende de week.
    • De nachtrust varieert in duur en kwaliteit, waarbij sommige nachten korter zijn maar wel als herstellend worden ervaren.

    ====================

    JSON OUTPUT:

    {{
    "datum": "",
    "naam": "",
    "volgende_consult": "",
    "huidige_situatie": "",
    "voeding_verminderen": "",
    "voeding_verhogen": "",
    "onderzoeken": "",
    "therapeuten": "",
    "supplementen": [
        {{
        "naam": "",
        "details": [],
        "voor_ontbijt": false,
        "ontbijt": false,
        "tussen_1": false,
        "lunch": false,
        "tussen_2": false,
        "diner": false,
        "voor_slapen": false
        }}
    ]
    }}

    ====================

    BELANGRIJK:
    Kopieer NIET de voorbeelden letterlijk.
    Kopieer alleen de stijl, structuur en diepgang.
    """
    
    USER = f"""
    TRANSCRIPT:
    {transcript}

    NOTITIES:
    {notes}
    """

    response = client.chat.completions.create(
        model="gpt-5-mini",
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": USER}
        ],
        temperature=1,
        response_format={"type": "json_object"}
    )

    return json.loads(response.choices[0].message.content)

def clean_supplements(data):
    for s in data.get("supplementen", []):
        # verwijder lege of "Onbekend" details
        if "details" in s:
            s["details"] = [
                d for d in s["details"]
                if d and d.strip() and d != "Onbekend"
            ]

            # als leeg → volledig weghalen
            if not s["details"]:
                s["details"] = []
    return data

def strip_bullets(data):
    for s in data.get("supplementen", []):
        if "details" in s:
            s["details"] = [
                d.replace("•", "").strip() for d in s["details"]
                if d
            ]
    return data

def kruis(val):
    return "✖" if val is True else ""
    
def safe_get(s, key):
    return s.get(key, False)    
# =============================
# WORD GENERATOR
# =============================
def generate_word(data, output="verslag.docx"):
    for s in data["supplementen"]:
        s["kruis_voor_ontbijt"] = kruis(safe_get(s, "voor_ontbijt"))
        s["kruis_ontbijt"] = kruis(safe_get(s, "ontbijt"))
        s["kruis_tussen_1"] = kruis(safe_get(s, "tussen_1"))
        s["kruis_lunch"] = kruis(safe_get(s, "lunch"))
        s["kruis_tussen_2"] = kruis(safe_get(s, "tussen_2"))
        s["kruis_diner"] = kruis(safe_get(s, "diner"))
        s["kruis_voor_slapen"] = kruis(safe_get(s, "voor_slapen"))

    doc = DocxTemplate("template.docx")
    
    doc.render(data)
    doc.save(output)

    print(f"Word opgeslagen: {output}")

# =============================
# MAIN
# =============================
if __name__ == "__main__":

    print("=== AI Consult Generator ===")

    keuze = input("1 = tekst | 2 = audio: ")

    if keuze == "1":
        transcript = input("Plak transcript:\n")
    elif keuze == "2":
        pad = input("Audio pad: ")
        transcript = transcribe_audio(pad)
        print("\nTranscript geladen.\n")
    else:
        print("Ongeldige keuze")
        exit()

    notes = input("Notities:\n")

    print("\nAI bezig...\n")

    data = generate_json(transcript, notes)
    data = clean_supplements(data)
    data = strip_bullets(data)

    print("\nWord genereren...\n")

    generate_word(data)

    print("\nKLAAR")