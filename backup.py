"""
backup.py — esporta tutto il DB in JSON e salva su GitHub
Da eseguire come cron job su Render (o manualmente)
Oppure aggiungere la route /backup_manuale per scaricarlo dal browser
"""
import os
import json
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime

DATABASE_URL = os.environ.get("DATABASE_URL")

def db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

def esporta_backup():
    conn = db()
    cur = conn.cursor()

    tabelle = ["stock", "produzione", "storico", "note", "materie_prime", "storico_mp"]
    backup = {"data_backup": datetime.now().isoformat(), "tabelle": {}}

    for tabella in tabelle:
        cur.execute(f"SELECT * FROM {tabella} ORDER BY id")
        righe = cur.fetchall()
        # Converti datetime in stringa
        righe_serializzabili = []
        for r in righe:
            riga = dict(r)
            for k, v in riga.items():
                if hasattr(v, "isoformat"):
                    riga[k] = v.isoformat()
            righe_serializzabili.append(riga)
        backup["tabelle"][tabella] = righe_serializzabili

    cur.close()
    conn.close()
    return backup

def salva_backup_locale(path="backup_db.json"):
    backup = esporta_backup()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(backup, f, ensure_ascii=False, indent=2)
    print(f"Backup salvato: {path} — {sum(len(v) for v in backup['tabelle'].values())} righe totali")
    return path

if __name__ == "__main__":
    salva_backup_locale()
