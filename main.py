import io
import os
import json
from datetime import datetime
from pathlib import Path

import httpx
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), ".env"))

app = FastAPI()

GROQ_API_KEY  = (os.getenv("GROQ_API_KEY") or "").strip()
DATA_FILE     = Path(__file__).parent / "data" / "dossiers.json"
SETTINGS_FILE = Path(__file__).parent / "data" / "settings.json"

SETTINGS_DEFAULT = {
    "profil": {"prenom": "", "nom": "", "titre": "", "cabinet": ""},
    "setup": {"context": "", "folders": [], "next_id": 1}
}


def load_data():
    if not DATA_FILE.exists():
        DATA_FILE.parent.mkdir(exist_ok=True)
        DATA_FILE.write_text(json.dumps({"dossiers": [], "next_id": 1001}, ensure_ascii=False))
    return json.loads(DATA_FILE.read_text(encoding="utf-8"))


def load_settings():
    if not SETTINGS_FILE.exists():
        SETTINGS_FILE.parent.mkdir(exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(SETTINGS_DEFAULT, ensure_ascii=False))
    s = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    # Migration anciens formats
    s["setup"].setdefault("folders", [])
    s["setup"].setdefault("next_id", 1)
    return s


def save_settings(s):
    SETTINGS_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")


def save_data(data):
    DATA_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


app.mount("/static", StaticFiles(directory="public"), name="static")


@app.get("/")
def index():
    return FileResponse("public/index.html")


# ─── Settings ────────────────────────────────────────────────────────────────

@app.get("/settings")
def get_settings():
    return load_settings()

@app.post("/settings")
async def post_settings(request: Request):
    body = await request.json()
    s = load_settings()
    if "profil" in body:
        s["profil"].update(body["profil"])
    if "setup" in body:
        for k, v in body["setup"].items():
            if k not in ("folders", "next_id"):   # ne pas écraser via cette route
                s["setup"][k] = v
    save_settings(s)
    return s


# ─── Setup : dossiers & items ─────────────────────────────────────────────────

@app.post("/settings/folders")
async def create_folder(request: Request):
    body = await request.json()
    name = body.get("name", "").strip()
    if not name:
        raise HTTPException(400, "Nom requis")
    s = load_settings()
    folder = {"id": str(s["setup"]["next_id"]), "name": name, "items": []}
    s["setup"]["next_id"] += 1
    s["setup"]["folders"].append(folder)
    save_settings(s)
    return folder

@app.put("/settings/folders/{fid}")
async def update_folder(fid: str, request: Request):
    body = await request.json()
    s = load_settings()
    folder = next((f for f in s["setup"]["folders"] if f["id"] == fid), None)
    if not folder:
        raise HTTPException(404, "Dossier introuvable")
    folder["name"] = body.get("name", folder["name"]).strip()
    save_settings(s)
    return folder

@app.delete("/settings/folders/{fid}")
async def delete_folder(fid: str):
    s = load_settings()
    s["setup"]["folders"] = [f for f in s["setup"]["folders"] if f["id"] != fid]
    save_settings(s)
    return {"ok": True}

@app.post("/settings/folders/{fid}/items")
async def create_item(fid: str, request: Request):
    body = await request.json()
    name = body.get("name", "").strip()
    if not name:
        raise HTTPException(400, "Nom requis")
    s = load_settings()
    folder = next((f for f in s["setup"]["folders"] if f["id"] == fid), None)
    if not folder:
        raise HTTPException(404, "Dossier introuvable")
    item = {"id": str(s["setup"]["next_id"]), "name": name, "detail": body.get("detail", "").strip()}
    s["setup"]["next_id"] += 1
    folder["items"].append(item)
    save_settings(s)
    return item

@app.put("/settings/folders/{fid}/items/{iid}")
async def update_item(fid: str, iid: str, request: Request):
    body = await request.json()
    s = load_settings()
    folder = next((f for f in s["setup"]["folders"] if f["id"] == fid), None)
    if not folder:
        raise HTTPException(404, "Dossier introuvable")
    item = next((i for i in folder["items"] if i["id"] == iid), None)
    if not item:
        raise HTTPException(404, "Item introuvable")
    if "name"   in body: item["name"]   = body["name"].strip()
    if "detail" in body: item["detail"] = body["detail"].strip()
    save_settings(s)
    return item

@app.delete("/settings/folders/{fid}/items/{iid}")
async def delete_item(fid: str, iid: str):
    s = load_settings()
    folder = next((f for f in s["setup"]["folders"] if f["id"] == fid), None)
    if folder:
        folder["items"] = [i for i in folder["items"] if i["id"] != iid]
        save_settings(s)
    return {"ok": True}


# ─── Dossiers ─────────────────────────────────────────────────────────────────

@app.get("/dossiers")
def list_dossiers():
    data = load_data()
    return {"dossiers": data["dossiers"], "next_id": data["next_id"]}

@app.post("/dossiers")
async def create_dossier(request: Request):
    body = await request.json()
    nom    = body.get("nom",    "").strip().upper()
    prenom = body.get("prenom", "").strip().capitalize()
    if not nom or not prenom:
        raise HTTPException(400, "Nom et prénom requis")
    data = load_data()
    dossier = {
        "id": data["next_id"],
        "nom": nom,
        "prenom": prenom,
        "created_at": datetime.now().isoformat(),
        "setup_items": [],
        "meetings": []
    }
    data["dossiers"].append(dossier)
    data["next_id"] += 1
    save_data(data)
    return dossier

@app.put("/dossiers/{dossier_id}/setup_items")
async def update_dossier_setup_items(dossier_id: int, request: Request):
    body = await request.json()
    data = load_data()
    dossier = next((d for d in data["dossiers"] if d["id"] == dossier_id), None)
    if not dossier:
        raise HTTPException(404, "Dossier introuvable")
    dossier["setup_items"] = body.get("setup_items", [])
    save_data(data)
    return {"ok": True}

@app.post("/dossiers/{dossier_id}/meetings")
async def add_meeting(dossier_id: int, request: Request):
    body = await request.json()
    data = load_data()
    dossier = next((d for d in data["dossiers"] if d["id"] == dossier_id), None)
    if not dossier:
        raise HTTPException(404, "Dossier introuvable")
    meeting = {
        "date": datetime.now().isoformat(),
        "transcription": body.get("transcription", ""),
        "compte_rendu":  body.get("compte_rendu",  "")
    }
    dossier["meetings"].append(meeting)
    save_data(data)
    return meeting

@app.patch("/dossiers/{dossier_id}/meetings/{meeting_index}")
async def update_meeting(dossier_id: int, meeting_index: int, request: Request):
    body = await request.json()
    data = load_data()
    dossier = next((d for d in data["dossiers"] if d["id"] == dossier_id), None)
    if not dossier:
        raise HTTPException(404, "Dossier introuvable")
    if meeting_index < 0 or meeting_index >= len(dossier["meetings"]):
        raise HTTPException(404, "Séance introuvable")
    dossier["meetings"][meeting_index]["compte_rendu"] = body.get("compte_rendu", "")
    save_data(data)
    return dossier["meetings"][meeting_index]

@app.delete("/dossiers/{dossier_id}")
async def delete_dossier(dossier_id: int):
    data = load_data()
    data["dossiers"] = [d for d in data["dossiers"] if d["id"] != dossier_id]
    save_data(data)
    return {"ok": True}


# ─── Export Word ─────────────────────────────────────────────────────────────

@app.post("/export-docx")
async def export_docx(request: Request):
    body = await request.json()
    text     = body.get("text", "").strip()
    patiente = body.get("patiente", "").strip()
    numero   = body.get("numero")
    date_str = body.get("date", "").strip()
    if not text:
        raise HTTPException(400, "Texte manquant")

    doc = Document()

    # Styles de base
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # Titre
    title_line = f"Compte rendu — {patiente}" if patiente else "Compte rendu"
    if patiente and numero:
        title_line += f" (N°{numero})"
    title = doc.add_heading(title_line, level=1)
    title.runs[0].font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    if date_str:
        p = doc.add_paragraph(f"Séance du {date_str}")
        p.runs[0].font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        p.runs[0].font.size = Pt(10)
        doc.add_paragraph("")

    # Parse les sections ## ...
    sections = text.split("\n")
    current_title = None
    current_lines = []

    def flush_section(stitle, slines):
        if stitle:
            doc.add_heading(stitle, level=2)
        for line in slines:
            line = line.strip()
            if not line:
                continue
            if line.startswith("- ") or line.startswith("• "):
                p = doc.add_paragraph(line[2:], style="List Bullet")
            else:
                doc.add_paragraph(line)

    for line in sections:
        if line.startswith("## "):
            flush_section(current_title, current_lines)
            current_title = line[3:].strip()
            current_lines = []
        else:
            current_lines.append(line)
    flush_section(current_title, current_lines)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    safe_name = patiente.replace(" ", "_") if patiente else "patiente"
    date_slug = datetime.now().strftime("%Y-%m-%d")
    filename  = f"CR_{safe_name}_{date_slug}.docx"

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ─── Q&A Neurosciences ───────────────────────────────────────────────────────

@app.post("/ask")
async def ask(request: Request):
    data = await request.json()
    question = data.get("question", "").strip()
    context  = data.get("context",  "").strip()
    if not question:
        raise HTTPException(400, "Question vide")

    system_msg = (
        "Tu es un assistant expert en neurosciences cliniques, neuropsychologie et psychiatrie. "
        "Tu aides un praticien de santé pendant la rédaction de comptes rendus de consultations. "
        "Réponds de manière concise, précise et professionnelle en français, "
        "en t'appuyant sur les données scientifiques actuelles et les meilleures pratiques cliniques."
    )

    user_msg = question
    if context:
        user_msg = f"Contexte — compte rendu de la séance en cours :\n{context}\n\nQuestion : {question}"

    async with httpx.AsyncClient() as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": "llama-3.3-70b-versatile",
                "messages": [
                    {"role": "system", "content": system_msg},
                    {"role": "user",   "content": user_msg}
                ],
                "temperature": 0.3
            },
            timeout=30.0,
        )
    result = resp.json()
    return {"answer": result["choices"][0]["message"]["content"]}


# ─── Transcription ────────────────────────────────────────────────────────────

@app.post("/transcribe")
async def transcribe(audio: UploadFile = File(...)):
    file_bytes = await audio.read()
    filename = audio.filename or "audio.mp3"
    content_type = audio.content_type or "audio/mpeg"

    async with httpx.AsyncClient() as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/audio/transcriptions",
            headers={"Authorization": f"Bearer {GROQ_API_KEY}"},
            files={"file": (filename, file_bytes, content_type)},
            data={"model": "whisper-large-v3", "language": "fr"},
            timeout=120.0,
        )

    if resp.status_code != 200:
        raise HTTPException(500, f"Erreur Groq transcription : {resp.text}")

    return {"transcription": resp.json()["text"].strip()}


# ─── Structuration ────────────────────────────────────────────────────────────

@app.post("/structure")
async def structure(request: Request):
    data = await request.json()
    transcription = data.get("transcription", "").strip()
    prenom        = data.get("prenom", "").strip()
    nom           = data.get("nom",    "").strip()
    dossier_id    = data.get("dossier_id")
    if not transcription:
        raise HTTPException(400, "Transcription vide")

    patient_ref = f"{prenom} {nom}".strip() if prenom else "la patiente"

    # Profil praticien + instructions générales
    settings = load_settings()
    profil       = settings.get("profil", {})
    setup_ctx    = settings.get("setup", {}).get("context", "").strip()
    praticien    = " ".join(p for p in [profil.get("titre",""), profil.get("prenom",""), profil.get("nom","")] if p).strip()
    cabinet      = profil.get("cabinet", "").strip()

    context_block = ""
    if praticien or cabinet or setup_ctx:
        context_block  = "Contexte du praticien :\n"
        if praticien: context_block += f"- Praticien : {praticien}\n"
        if cabinet:   context_block += f"- Cabinet : {cabinet}\n"
        if setup_ctx: context_block += f"- Instructions générales : {setup_ctx}\n"
        context_block += "\n"

    # Items setup associés au dossier patient
    items_block = ""
    if dossier_id:
        try:
            doss_data = load_data()
            doss = next((d for d in doss_data["dossiers"] if d["id"] == int(dossier_id)), None)
            if doss and doss.get("setup_items"):
                items_map = {
                    i["id"]: (f["name"], i["name"], i.get("detail", ""))
                    for f in settings.get("setup", {}).get("folders", [])
                    for i in f.get("items", [])
                }
                lines = []
                for iid in doss["setup_items"]:
                    if iid in items_map:
                        fn, iname, detail = items_map[iid]
                        line = f"- {fn} / {iname}"
                        if detail:
                            line += f" : {detail}"
                        lines.append(line)
                if lines:
                    items_block = "Informations spécifiques à " + (prenom or "la patiente") + " :\n" + "\n".join(lines) + "\n\n"
        except Exception:
            pass

    prompt = f"""Tu es un assistant expert en comptes rendus de consultation médicale ou de suivi patient.

{context_block}{items_block}La patiente s'appelle {patient_ref}. Dans tout le compte rendu, réfère-toi à elle uniquement par son prénom "{prenom or patient_ref}".

Voici la transcription brute de la séance :
\"\"\"{transcription}\"\"\"

Génère un compte rendu structuré, clair et professionnel en français avec exactement ces sections :

## Résumé de la séance
(2-3 phrases résumant l'essentiel, en utilisant le prénom de la patiente)

## Points abordés
(liste à puces)

## Observations et décisions
(liste à puces, ou "Aucune" si applicable)

## Actions à suivre
(liste à puces)

## Prochaine séance
(liste à puces ou date/objectif si mentionné)

Sois concis, factuel et professionnel. Respecte les instructions spécifiques. Utilise le prénom {prenom or "de la patiente"} dès que tu parles d'elle."""

    async with httpx.AsyncClient() as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"},
            json={"model": "llama-3.3-70b-versatile", "messages": [{"role": "user", "content": prompt}], "temperature": 0.3},
            timeout=30.0,
        )
    result = resp.json()
    return {"structured": result["choices"][0]["message"]["content"]}
