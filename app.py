from flask import Flask, render_template, request, redirect, send_file, session
import os
import psycopg2
from psycopg2.extras import RealDictCursor
from openpyxl import load_workbook
from datetime import datetime
import json

app = Flask(__name__)
app.secret_key = "lc_wine_secret_2026"

DATABASE_URL = os.environ.get("DATABASE_URL")
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
MODELLI_DIR  = os.path.join(BASE_DIR, "modelli")


def db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


def init_db():
    conn = db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS stock (
            id SERIAL PRIMARY KEY, cliente TEXT, prodotto TEXT, qty INTEGER
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS produzione (
            id SERIAL PRIMARY KEY, cliente TEXT, prodotto TEXT,
            qty INTEGER, done INTEGER DEFAULT 0
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS storico (
            id SERIAL PRIMARY KEY, cliente TEXT, prodotto TEXT,
            qty INTEGER, tipo TEXT, data TIMESTAMP DEFAULT NOW()
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS note (
            id SERIAL PRIMARY KEY,
            testo TEXT,
            data TIMESTAMP DEFAULT NOW()
        )
    """)
    conn.commit()
    cur.close()
    conn.close()

init_db()

# ==================================================
# CONFIGURAZIONE CLIENTI
# L'utente inserisce FARDELLI — internamente
# convertiamo in bottiglie (fardelli x moltiplicatore)
# ==================================================

clients = {
    "Roberto": [
        "Catarratto 2L", "Rosato 2L", "Merlot 2L", "Il Nero 2L",
        "Bianco E.N. 2L", "Rosato E.N. 2L", "Rosso E.N. 2L",
        "Catarratto 1L", "Rosato 1L", "Merlot 1L",
        "Bianco S.E. 1L", "Rosato S.E. 1L", "Rosso S.E. 1L",
    ],
    "Francesco": [
        "Catarratto 2L", "Chardonnay 2L", "Rosato 2L", "Merlot 2L", "Syrah 2L",
        "Catarratto 1L", "Syrah 1L",
    ],
    "Emanuele": [
        "Catarratto 2L", "Rosato 2L", "Il Nero 2L", "Merlot 2L",
        "Vino Rosso 2L", "Syrah 2L",
        "Bianco S.E. 2L", "Rosato S.E. 2L", "Rosso S.E. 2L",
        "Catarratto R.B. 2L", "Rosato R.B. 2L", "Il Nero R.B. 2L", "Vino Rosso R.B. 2L",
        "Catarratto 1L", "Rosato 1L", "Il Nero 1L",
        "Bianco S.E. 1L", "Rosato S.E. 1L", "Rosso S.E. 1L",
    ],
    "Mazzarrone": [
        "Divino Bianco 2L", "Divino Rosato 2L", "Divino Rosso 2L", "Divino Syrah 2L",
        "Divino Bianco 1L", "Divino Rosato 1L", "Divino Rosso 1L", "Divino Syrah 1L",
        "Pachinos Bianco 2L", "Pachinos Rosato 2L", "Pachinos Rosso 2L", "Pachinos Syrah 2L",
        "Pachinos Bianco 1L", "Pachinos Rosato 1L", "Pachinos Rosso 1L", "Pachinos Syrah 1L",
    ],
    "Sisa": [
        "Bianco 2L", "Rosso 2L",
    ],
}

# Moltiplicatore fardello->bottiglie per ogni prodotto di ogni cliente
MOLTIPLICATORI = {
    "Roberto": {
        "Catarratto 2L": 6,  "Rosato 2L": 6,      "Merlot 2L": 6,
        "Il Nero 2L": 6,     "Bianco E.N. 2L": 6,  "Rosato E.N. 2L": 6,
        "Rosso E.N. 2L": 6,  "Catarratto 1L": 12,  "Rosato 1L": 12,
        "Merlot 1L": 12,     "Bianco S.E. 1L": 16, "Rosato S.E. 1L": 16,
        "Rosso S.E. 1L": 16,
    },
    "Francesco": {
        "Catarratto 2L": 6,  "Chardonnay 2L": 6,  "Rosato 2L": 6,
        "Merlot 2L": 6,      "Syrah 2L": 6,
        "Catarratto 1L": 12, "Syrah 1L": 12,
    },
    "Emanuele": {
        "Catarratto 2L": 9,      "Rosato 2L": 9,          "Il Nero 2L": 9,
        "Merlot 2L": 9,          "Vino Rosso 2L": 9,       "Syrah 2L": 9,
        "Bianco S.E. 2L": 9,     "Rosato S.E. 2L": 9,      "Rosso S.E. 2L": 9,
        "Catarratto R.B. 2L": 9, "Rosato R.B. 2L": 9,      "Il Nero R.B. 2L": 9,
        "Vino Rosso R.B. 2L": 9, "Catarratto 1L": 16,      "Rosato 1L": 16,
        "Il Nero 1L": 16,        "Bianco S.E. 1L": 16,     "Rosato S.E. 1L": 16,
        "Rosso S.E. 1L": 16,
    },
    "Mazzarrone": {
        "Divino Bianco 2L": 6,   "Divino Rosato 2L": 6,
        "Divino Rosso 2L": 6,    "Divino Syrah 2L": 6,
        "Divino Bianco 1L": 12,  "Divino Rosato 1L": 12,
        "Divino Rosso 1L": 12,   "Divino Syrah 1L": 12,
        "Pachinos Bianco 2L": 6, "Pachinos Rosato 2L": 6,
        "Pachinos Rosso 2L": 6,  "Pachinos Syrah 2L": 6,
        "Pachinos Bianco 1L": 12,"Pachinos Rosato 1L": 12,
        "Pachinos Rosso 1L": 12, "Pachinos Syrah 1L": 12,
    },
    "Sisa": {
        "Bianco 2L": 6, "Rosso 2L": 6,
    },
}

CLIENTI_CONFIG = {
    "Roberto": {
        "bolla": "bolla_roberto.xlsx", "conteggio": "conteggio_roberto.xlsx",
        "prodotti": {
            "Catarratto 2L": 23,  "Rosato 2L": 25,      "Merlot 2L": 27,
            "Il Nero 2L": 29,     "Bianco E.N. 2L": 31,  "Rosato E.N. 2L": 33,
            "Rosso E.N. 2L": 35,  "Catarratto 1L": 37,   "Rosato 1L": 39,
            "Merlot 1L": 41,      "Bianco S.E. 1L": 43,  "Rosato S.E. 1L": 45,
            "Rosso S.E. 1L": 47,
        },
        "righe_az": [23,25,27,29,31,33,35,37,39,41,43,45,47],
        "cella_titolo_bolla": "H2",     "cella_data_bolla": "F55",
        "cella_titolo_conteggio": "G2", "cella_data_conteggio": "F53",
    },
    "Francesco": {
        "bolla": "bolla_francesco.xlsx", "conteggio": "conteggio_francesco.xlsx",
        "prodotti": {
            "Catarratto 2L": 21, "Chardonnay 2L": 23, "Rosato 2L": 25,
            "Merlot 2L": 27,     "Syrah 2L": 29,
            "Catarratto 1L": 43, "Syrah 1L": 45,
        },
        "righe_az": [21,23,25,27,29,35,37,39,43,45],
        "cella_titolo_bolla": "H2",     "cella_data_bolla": "F55",
        "cella_titolo_conteggio": "H2", "cella_data_conteggio": "F65",
    },
    "Emanuele": {
        "bolla": "bolla_emanuele.xlsx", "conteggio": "conteggio_emanuele.xlsx",
        "prodotti_bolla": {
            "Catarratto 2L": 21,      "Rosato 2L": 23,         "Il Nero 2L": 25,
            "Vino Rosso 2L": 27,      "Syrah 2L": 29,           "Merlot 2L": 31,
            "Catarratto 1L": 33,      "Rosato 1L": 35,          "Il Nero 1L": 37,
            "Catarratto R.B. 2L": 41, "Rosato R.B. 2L": 43,    "Il Nero R.B. 2L": 45,
            "Vino Rosso R.B. 2L": 47, "Bianco S.E. 2L": 51,    "Rosato S.E. 2L": 53,
            "Rosso S.E. 2L": 55,      "Bianco S.E. 1L": 57,    "Rosato S.E. 1L": 59,
            "Rosso S.E. 1L": 61,
        },
        "prodotti_conteggio": {
            "Catarratto 2L": 25,      "Rosato 2L": 27,          "Il Nero 2L": 29,
            "Vino Rosso 2L": 31,      "Syrah 2L": 33,           "Merlot 2L": 35,
            "Catarratto 1L": 37,      "Rosato 1L": 39,          "Il Nero 1L": 41,
            "Catarratto R.B. 2L": 45, "Rosato R.B. 2L": 47,    "Il Nero R.B. 2L": 49,
            "Vino Rosso R.B. 2L": 51, "Bianco S.E. 2L": 55,    "Rosato S.E. 2L": 57,
            "Rosso S.E. 2L": 59,      "Bianco S.E. 1L": 61,    "Rosato S.E. 1L": 63,
            "Rosso S.E. 1L": 65,
        },
        "righe_az_bolla":     [21,23,25,27,29,31,33,35,37,41,43,45,47,51,53,55,57,59,61],
        "righe_az_conteggio": [25,27,29,31,33,35,37,39,41,45,47,49,51,55,57,59,61,63,65],
        "cella_titolo_bolla": "H2",     "cella_data_bolla": "F67",
        "cella_titolo_conteggio": "H2", "cella_data_conteggio": "F71",
    },
    "Mazzarrone": {
        "bolla": None, "conteggio": None, "prodotti": {},
        "righe_az": [],
        "cella_titolo_bolla": None, "cella_data_bolla": None,
        "cella_titolo_conteggio": None, "cella_data_conteggio": None,
    },
    "Sisa": {
        "bolla": None, "conteggio": None, "prodotti": {},
        "righe_az": [],
        "cella_titolo_bolla": None, "cella_data_bolla": None,
        "cella_titolo_conteggio": None, "cella_data_conteggio": None,
    },
}


# ==================================================
# FUNZIONI EXCEL
# ==================================================

def _aggiorna_excel(ws, cliente, richieste_fardelli, tipo="bolla"):
    """richieste_fardelli: lista di (prodotto, fardelli)"""
    cfg = CLIENTI_CONFIG[cliente]

    if cliente == "Emanuele":
        mappa    = cfg["prodotti_bolla"] if tipo == "bolla" else cfg["prodotti_conteggio"]
        righe_az = cfg["righe_az_bolla"] if tipo == "bolla" else cfg["righe_az_conteggio"]
    else:
        mappa    = cfg.get("prodotti", {})
        righe_az = cfg.get("righe_az", [])

    for riga in righe_az:
        ws[f"G{riga}"] = 0

    for prodotto, fardelli in richieste_fardelli:
        if prodotto in mappa:
            ws[f"G{mappa[prodotto]}"] = fardelli


def _genera_file(cliente, richieste_fardelli, tipo):
    cfg = CLIENTI_CONFIG[cliente]
    nome_modello = cfg["bolla"] if tipo == "bolla" else cfg["conteggio"]

    if not nome_modello:
        return None, "Modello non ancora disponibile per questo cliente"

    file_modello = os.path.join(MODELLI_DIR, nome_modello)
    if not os.path.exists(file_modello):
        return None, f"File modello mancante: {nome_modello}"

    wb = load_workbook(file_modello)
    ws = wb.active

    cella_titolo = cfg[f"cella_titolo_{tipo}"]
    cella_data   = cfg[f"cella_data_{tipo}"]
    if cella_titolo:
        ws[cella_titolo] = "DOCUMENTO DI TRASPORTO\nN.          DEL\n"
    if cella_data:
        ws[cella_data] = "DATA RITIRO\n\n\n"

    _aggiorna_excel(ws, cliente, richieste_fardelli, tipo)

    nome_out = f"{tipo}_generato_{cliente.lower()}.xlsx"
    output   = os.path.join(BASE_DIR, nome_out)
    wb.save(output)
    return output, None


def _leggi_richieste_fardelli(cliente, form):
    """
    Legge dal form i fardelli inseriti dall'utente.
    Ritorna lista di (prodotto, fardelli) — fardelli è già l'unità giusta per Excel.
    Salva anche in sessione i fardelli (non le bottiglie).
    """
    richieste = []
    molt = MOLTIPLICATORI.get(cliente, {})
    for i, prodotto in enumerate(clients[cliente]):
        val = form.get(f"qty_{i}")
        if val and val.isdigit():
            f = int(val)
            if f > 0:
                richieste.append((prodotto, f))
    return richieste


def _fardelli_a_bottiglie(cliente, richieste_fardelli):
    """Converte lista (prodotto, fardelli) in (prodotto, bottiglie) per il magazzino."""
    molt = MOLTIPLICATORI.get(cliente, {})
    return [(p, f * molt.get(p, 1)) for p, f in richieste_fardelli]


# ==================================================
# HOME
# ==================================================
@app.route("/")
def home():
    return render_template("home.html")


# ==================================================
# STORICO
# ==================================================
@app.route("/storico")
def storico():
    conn = db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM storico ORDER BY data DESC LIMIT 300")
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return render_template("storico.html", rows=rows)


# ==================================================
# PRODUZIONE
# ==================================================
@app.route("/produzione")
def produzione():
    conn = db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM produzione ORDER BY id")
    righe = cur.fetchall()
    cur.execute("SELECT * FROM note ORDER BY data DESC")
    note = cur.fetchall()
    cur.close()
    conn.close()
    return render_template("produzione.html", clients=clients, rows=righe, note=note)


@app.route("/nuova_produzione", methods=["POST"])
def nuova_produzione():
    cliente = request.form["client"]
    conn = db()
    cur = conn.cursor()
    molt = MOLTIPLICATORI.get(cliente, {})
    for i, prodotto in enumerate(clients[cliente]):
        val = request.form.get(f"qty_{i}")
        if val and val.isdigit():
            f = int(val)
            if f > 0:
                bottiglie = f * molt.get(prodotto, 1)
                cur.execute(
                    "INSERT INTO produzione(cliente, prodotto, qty) VALUES(%s,%s,%s)",
                    (cliente, prodotto, bottiglie)
                )
                cur.execute(
                    "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
                    (cliente, prodotto, bottiglie, "Produzione Inserita")
                )
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/produzione")


@app.route("/toggle/<int:id>")
def toggle(id):
    conn = db()
    cur = conn.cursor()
    cur.execute("SELECT done FROM produzione WHERE id=%s", (id,))
    row = cur.fetchone()
    nuovo = 0 if row["done"] == 1 else 1
    cur.execute("UPDATE produzione SET done=%s WHERE id=%s", (nuovo, id))
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/produzione")


@app.route("/passa_magazzino")
def passa_magazzino():
    conn = db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM produzione WHERE done=1")
    finiti = cur.fetchall()
    for r in finiti:
        cur.execute(
            "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
            (r["cliente"], r["prodotto"])
        )
        ex = cur.fetchone()
        if ex:
            cur.execute("UPDATE stock SET qty=%s WHERE id=%s",
                        (ex["qty"] + r["qty"], ex["id"]))
        else:
            cur.execute("INSERT INTO stock(cliente, prodotto, qty) VALUES(%s,%s,%s)",
                        (r["cliente"], r["prodotto"], r["qty"]))
        cur.execute(
            "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
            (r["cliente"], r["prodotto"], r["qty"], "Passato a Magazzino")
        )
        cur.execute("DELETE FROM produzione WHERE id=%s", (r["id"],))
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/produzione")


# NOTE PRODUZIONE
@app.route("/aggiungi_nota", methods=["POST"])
def aggiungi_nota():
    testo = request.form.get("testo", "").strip()
    if testo:
        conn = db()
        cur = conn.cursor()
        cur.execute("INSERT INTO note(testo) VALUES(%s)", (testo,))
        conn.commit()
        cur.close()
        conn.close()
    return redirect("/produzione")


@app.route("/elimina_nota/<int:id>")
def elimina_nota(id):
    conn = db()
    cur = conn.cursor()
    cur.execute("DELETE FROM note WHERE id=%s", (id,))
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/produzione")


# ==================================================
# MAGAZZINO
# ==================================================
@app.route("/magazzino")
def magazzino():
    msg = request.args.get("msg", "")
    cliente_sel = request.args.get("cliente", "")
    conn = db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM stock WHERE qty > 0 ORDER BY cliente, prodotto")
    rows = cur.fetchall()
    cur.close()
    conn.close()
    grouped = {}
    for r in rows:
        grouped.setdefault(r["cliente"], []).append(r)
    return render_template("magazzino.html", grouped=grouped, clients=clients,
                           msg=msg, cliente_sel=cliente_sel)


@app.route("/scarica", methods=["POST"])
def scarica():
    cliente = request.form["client"]
    molt = MOLTIPLICATORI.get(cliente, {})
    richieste_bt = []
    for i, prodotto in enumerate(clients[cliente]):
        val = request.form.get(f"qty_{i}")
        if val and val.isdigit():
            f = int(val)
            if f > 0:
                richieste_bt.append((prodotto, f * molt.get(prodotto, 1)))
    if not richieste_bt:
        return redirect("/magazzino?msg=Nessun prodotto selezionato&cliente=" + cliente)
    conn = db()
    cur = conn.cursor()
    for prodotto, q in richieste_bt:
        cur.execute("SELECT * FROM stock WHERE cliente=%s AND prodotto=%s", (cliente, prodotto))
        row = cur.fetchone()
        if not row:
            cur.close(); conn.close()
            return redirect("/magazzino?msg=" + prodotto + " non presente&cliente=" + cliente)
        if row["qty"] < q:
            cur.close(); conn.close()
            return redirect("/magazzino?msg=" + prodotto + " quantita insufficiente&cliente=" + cliente)
    for prodotto, q in richieste_bt:
        cur.execute("SELECT * FROM stock WHERE cliente=%s AND prodotto=%s", (cliente, prodotto))
        row = cur.fetchone()
        nuova = row["qty"] - q
        if nuova == 0:
            cur.execute("DELETE FROM stock WHERE id=%s", (row["id"],))
        else:
            cur.execute("UPDATE stock SET qty=%s WHERE id=%s", (nuova, row["id"]))
        cur.execute(
            "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
            (cliente, prodotto, q, "Scarico Magazzino")
        )
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/magazzino?msg=Scarico completato&cliente=" + cliente)


# ==================================================
# CONSEGNE
# ==================================================
@app.route("/consegne")
def consegne():
    msg = request.args.get("msg", "")
    cliente_sel = request.args.get("cliente", "")
    conn = db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM stock WHERE qty > 0 ORDER BY cliente, prodotto")
    rows = cur.fetchall()
    cur.close()
    conn.close()
    grouped = {}
    for r in rows:
        grouped.setdefault(r["cliente"], []).append(r)
    return render_template("consegne.html", clients=clients, grouped=grouped,
                           cliente_sel=cliente_sel, msg=msg,
                           moltiplicatori=MOLTIPLICATORI)


@app.route("/esegui_consegna", methods=["POST"])
def esegui_consegna():
    cliente = request.form["client"]
    richieste_f = _leggi_richieste_fardelli(cliente, request.form)
    if not richieste_f:
        return redirect("/consegne?msg=Nessun prodotto selezionato&cliente=" + cliente)
    richieste_bt = _fardelli_a_bottiglie(cliente, richieste_f)
    conn = db()
    cur = conn.cursor()
    for prodotto, q in richieste_bt:
        cur.execute("SELECT * FROM stock WHERE cliente=%s AND prodotto=%s", (cliente, prodotto))
        row = cur.fetchone()
        if not row:
            cur.close(); conn.close()
            return redirect("/consegne?msg=" + prodotto + " non presente&cliente=" + cliente)
        if row["qty"] < q:
            cur.close(); conn.close()
            return redirect("/consegne?msg=" + prodotto + " quantita insufficiente&cliente=" + cliente)
    for prodotto, q in richieste_bt:
        cur.execute("SELECT * FROM stock WHERE cliente=%s AND prodotto=%s", (cliente, prodotto))
        row = cur.fetchone()
        nuova = row["qty"] - q
        if nuova == 0:
            cur.execute("DELETE FROM stock WHERE id=%s", (row["id"],))
        else:
            cur.execute("UPDATE stock SET qty=%s WHERE id=%s", (nuova, row["id"]))
        cur.execute(
            "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
            (cliente, prodotto, q, "Consegna")
        )
    conn.commit()
    cur.close()
    conn.close()
    # Salva fardelli in sessione (per Excel) e bottiglie per riepilogo
    session["consegna_cliente"]   = cliente
    session["consegna_fardelli"]  = json.dumps(richieste_f)
    session["consegna_bottiglie"] = json.dumps(richieste_bt)
    return redirect("/conferma_consegna")


@app.route("/conferma_consegna")
def conferma_consegna():
    cliente    = session.get("consegna_cliente", "")
    richieste_f  = json.loads(session.get("consegna_fardelli", "[]"))
    richieste_bt = json.loads(session.get("consegna_bottiglie", "[]"))
    if not cliente or not richieste_f:
        return redirect("/consegne?msg=Nessuna consegna attiva")
    molt = MOLTIPLICATORI.get(cliente, {})
    return render_template("conferma_consegna.html", cliente=cliente,
                           richieste_f=richieste_f, richieste_bt=richieste_bt,
                           molt=molt)


@app.route("/download_bolla")
def download_bolla():
    cliente    = session.get("consegna_cliente", "")
    richieste_f = json.loads(session.get("consegna_fardelli", "[]"))
    if not cliente or not richieste_f:
        return redirect("/consegne?msg=Nessuna consegna attiva")
    output, errore = _genera_file(cliente, richieste_f, "bolla")
    if errore:
        return redirect("/conferma_consegna?msg=" + errore)
    return send_file(output, as_attachment=True, download_name=f"Bolla_{cliente}.xlsx")


@app.route("/download_conteggio")
def download_conteggio():
    cliente    = session.get("consegna_cliente", "")
    richieste_f = json.loads(session.get("consegna_fardelli", "[]"))
    if not cliente or not richieste_f:
        return redirect("/consegne?msg=Nessuna consegna attiva")
    output, errore = _genera_file(cliente, richieste_f, "conteggio")
    if errore:
        return redirect("/conferma_consegna?msg=" + errore)
    return send_file(output, as_attachment=True, download_name=f"Conteggio_{cliente}.xlsx")


@app.route("/solo_bolla", methods=["POST"])
def solo_bolla():
    cliente = request.form["client"]
    richieste_f = _leggi_richieste_fardelli(cliente, request.form)
    if not richieste_f:
        return redirect("/consegne?msg=Nessun prodotto selezionato&cliente=" + cliente)
    output, errore = _genera_file(cliente, richieste_f, "bolla")
    if errore:
        return redirect("/consegne?msg=" + errore + "&cliente=" + cliente)
    return send_file(output, as_attachment=True, download_name=f"Bolla_{cliente}.xlsx")


@app.route("/solo_conteggio", methods=["POST"])
def solo_conteggio():
    cliente = request.form["client"]
    richieste_f = _leggi_richieste_fardelli(cliente, request.form)
    if not richieste_f:
        return redirect("/consegne?msg=Nessun prodotto selezionato&cliente=" + cliente)
    output, errore = _genera_file(cliente, richieste_f, "conteggio")
    if errore:
        return redirect("/consegne?msg=" + errore + "&cliente=" + cliente)
    return send_file(output, as_attachment=True, download_name=f"Conteggio_{cliente}.xlsx")


# ==================================================
# START
# ==================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
