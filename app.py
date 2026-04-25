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
    conn.commit()
    cur.close()
    conn.close()

init_db()

# ==================================================
# CONFIGURAZIONE CLIENTI
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

# --------------------------------------------------
# MAPPA PRODOTTI -> RIGA TEMPLATE + MOLTIPLICATORE
# (riga colonna G nel file Excel, bt per fardello)
# --------------------------------------------------

CLIENTI_CONFIG = {
    "Roberto": {
        "bolla":     "bolla_roberto.xlsx",
        "conteggio": "conteggio_roberto.xlsx",
        "prodotti": {
            "Catarratto 2L":  (23, 6),
            "Rosato 2L":      (25, 6),
            "Merlot 2L":      (27, 6),
            "Il Nero 2L":     (29, 6),
            "Bianco E.N. 2L": (31, 6),
            "Rosato E.N. 2L": (33, 6),
            "Rosso E.N. 2L":  (35, 6),
            "Catarratto 1L":  (37, 12),
            "Rosato 1L":      (39, 12),
            "Merlot 1L":      (41, 12),
            "Bianco S.E. 1L": (43, 16),
            "Rosato S.E. 1L": (45, 16),
            "Rosso S.E. 1L":  (47, 16),
        },
        "cella_titolo_bolla":     "H2",
        "cella_data_bolla":       "F55",
        "cella_titolo_conteggio": "G2",
        "cella_data_conteggio":   "F53",
    },
    "Francesco": {
        "bolla":     "bolla_francesco.xlsx",
        "conteggio": "conteggio_francesco.xlsx",
        "prodotti": {
            "Catarratto 2L":  (21, 6),
            "Chardonnay 2L":  (23, 6),
            "Rosato 2L":      (25, 6),
            "Merlot 2L":      (27, 6),
            "Syrah 2L":       (29, 6),
            "Catarratto 1L":  (43, 12),
            "Syrah 1L":       (45, 12),
        },
        "cella_titolo_bolla":     "H2",
        "cella_data_bolla":       "F55",
        "cella_titolo_conteggio": "H2",
        "cella_data_conteggio":   "F65",
    },
    "Emanuele": {
        "bolla":     "bolla_emanuele.xlsx",
        "conteggio": "conteggio_emanuele.xlsx",
        "prodotti": {
            # bolla: righe colonna G
            # conteggio: stesse righe (verificate)
            "Catarratto 2L":     (21, 9),   # bolla G21 / conteggio G25
            "Rosato 2L":         (23, 9),
            "Il Nero 2L":        (25, 9),
            "Vino Rosso 2L":     (27, 9),
            "Syrah 2L":          (29, 9),
            "Merlot 2L":         (31, 9),
            "Catarratto 1L":     (33, 16),
            "Rosato 1L":         (35, 16),
            "Il Nero 1L":        (37, 16),
            "Catarratto R.B. 2L":(41, 9),
            "Rosato R.B. 2L":    (43, 9),
            "Il Nero R.B. 2L":   (45, 9),
            "Vino Rosso R.B. 2L":(47, 9),
            "Bianco S.E. 2L":    (51, 9),
            "Rosato S.E. 2L":    (53, 9),
            "Rosso S.E. 2L":     (55, 9),
            "Bianco S.E. 1L":    (57, 16),
            "Rosato S.E. 1L":    (59, 16),
            "Rosso S.E. 1L":     (61, 16),
        },
        "cella_titolo_bolla":     "H2",
        "cella_data_bolla":       "F67",
        "cella_titolo_conteggio": "H2",
        "cella_data_conteggio":   "F71",
    },
    "Mazzarrone": {
        "bolla":     None,
        "conteggio": None,
        "prodotti":  {},
        "cella_titolo_bolla":     None,
        "cella_data_bolla":       None,
        "cella_titolo_conteggio": None,
        "cella_data_conteggio":   None,
    },
    "Sisa": {
        "bolla":     None,
        "conteggio": None,
        "prodotti":  {},
        "cella_titolo_bolla":     None,
        "cella_data_bolla":       None,
        "cella_titolo_conteggio": None,
        "cella_data_conteggio":   None,
    },
}

# Righe da azzerare per ogni cliente (tutte le righe prodotto del template)
RIGHE_AZZERAMENTO = {
    "Roberto":   [23,25,27,29,31,33,35,37,39,41,43,45,47],
    "Francesco": [21,23,25,27,29,35,37,39,43,45],
    "Emanuele":  [21,23,25,27,29,31,33,35,37,41,43,45,47,51,53,55,57,59,61],
    "Mazzarrone": [],
    "Sisa":       [],
}

# Per il conteggio Emanuele le righe G sono diverse dalla bolla
EMANUELE_RIGHE_CONTEGGIO = {
    "Catarratto 2L":     25,
    "Rosato 2L":         27,
    "Il Nero 2L":        29,
    "Vino Rosso 2L":     31,
    "Syrah 2L":          33,
    "Merlot 2L":         35,
    "Catarratto 1L":     37,
    "Rosato 1L":         39,
    "Il Nero 1L":        41,
    "Catarratto R.B. 2L":45,
    "Rosato R.B. 2L":    47,
    "Il Nero R.B. 2L":   49,
    "Vino Rosso R.B. 2L":51,
    "Bianco S.E. 2L":    55,
    "Rosato S.E. 2L":    57,
    "Rosso S.E. 2L":     59,
    "Bianco S.E. 1L":    61,
    "Rosato S.E. 1L":    63,
    "Rosso S.E. 1L":     65,
}
RIGHE_AZZERAMENTO_CONTEGGIO_EMANUELE = [25,27,29,31,33,35,37,39,41,45,47,49,51,55,57,59,61,63,65]


# ==================================================
# FUNZIONI EXCEL
# ==================================================

def _aggiorna_excel(ws, cliente, richieste, tipo="bolla"):
    cfg = CLIENTI_CONFIG[cliente]

    if tipo == "conteggio" and cliente == "Emanuele":
        righe_az = RIGHE_AZZERAMENTO_CONTEGGIO_EMANUELE
        for riga in righe_az:
            ws[f"G{riga}"] = 0
        for prodotto, qty_bt in richieste:
            if prodotto in EMANUELE_RIGHE_CONTEGGIO:
                riga = EMANUELE_RIGHE_CONTEGGIO[prodotto]
                molt = cfg["prodotti"][prodotto][1]
                ws[f"G{riga}"] = qty_bt / molt
    else:
        for riga in RIGHE_AZZERAMENTO.get(cliente, []):
            ws[f"G{riga}"] = 0
        for prodotto, qty_bt in richieste:
            if prodotto in cfg["prodotti"]:
                riga, molt = cfg["prodotti"][prodotto]
                ws[f"G{riga}"] = qty_bt / molt


def _genera_file(cliente, richieste, tipo):
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

    _aggiorna_excel(ws, cliente, richieste, tipo)

    nome_out = f"{tipo}_generato_{cliente.lower()}.xlsx"
    output   = os.path.join(BASE_DIR, nome_out)
    wb.save(output)
    return output, None


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
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return render_template("produzione.html", clients=clients, rows=rows)


@app.route("/nuova_produzione", methods=["POST"])
def nuova_produzione():
    cliente = request.form["client"]
    conn = db()
    cur = conn.cursor()
    for i, prodotto in enumerate(clients[cliente]):
        qty = request.form.get(f"qty_{i}")
        if qty and qty.isdigit():
            q = int(qty)
            if q > 0:
                cur.execute(
                    "INSERT INTO produzione(cliente, prodotto, qty) VALUES(%s,%s,%s)",
                    (cliente, prodotto, q)
                )
                cur.execute(
                    "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
                    (cliente, prodotto, q, "Produzione Inserita")
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
    richieste = []
    for i, prodotto in enumerate(clients[cliente]):
        qty = request.form.get(f"qty_{i}")
        if qty and qty.isdigit():
            q = int(qty)
            if q > 0:
                richieste.append((prodotto, q))
    if not richieste:
        return redirect("/magazzino?msg=Nessun prodotto selezionato&cliente=" + cliente)
    conn = db()
    cur = conn.cursor()
    for prodotto, q in richieste:
        cur.execute("SELECT * FROM stock WHERE cliente=%s AND prodotto=%s", (cliente, prodotto))
        row = cur.fetchone()
        if not row:
            cur.close(); conn.close()
            return redirect("/magazzino?msg=" + prodotto + " non presente&cliente=" + cliente)
        if row["qty"] < q:
            cur.close(); conn.close()
            return redirect("/magazzino?msg=" + prodotto + " quantita insufficiente&cliente=" + cliente)
    for prodotto, q in richieste:
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
                           cliente_sel=cliente_sel, msg=msg)


@app.route("/esegui_consegna", methods=["POST"])
def esegui_consegna():
    cliente = request.form["client"]
    richieste = []
    for i, prodotto in enumerate(clients[cliente]):
        qty = request.form.get(f"qty_{i}")
        if qty and qty.isdigit():
            q = int(qty)
            if q > 0:
                richieste.append((prodotto, q))
    if not richieste:
        return redirect("/consegne?msg=Nessun prodotto selezionato&cliente=" + cliente)
    conn = db()
    cur = conn.cursor()
    for prodotto, q in richieste:
        cur.execute("SELECT * FROM stock WHERE cliente=%s AND prodotto=%s", (cliente, prodotto))
        row = cur.fetchone()
        if not row:
            cur.close(); conn.close()
            return redirect("/consegne?msg=" + prodotto + " non presente&cliente=" + cliente)
        if row["qty"] < q:
            cur.close(); conn.close()
            return redirect("/consegne?msg=" + prodotto + " quantita insufficiente&cliente=" + cliente)
    for prodotto, q in richieste:
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
    session["consegna_cliente"]   = cliente
    session["consegna_richieste"] = json.dumps(richieste)
    return redirect("/conferma_consegna")


@app.route("/conferma_consegna")
def conferma_consegna():
    cliente   = session.get("consegna_cliente", "")
    richieste = json.loads(session.get("consegna_richieste", "[]"))
    if not cliente or not richieste:
        return redirect("/consegne?msg=Nessuna consegna attiva")
    return render_template("conferma_consegna.html", cliente=cliente, richieste=richieste)


@app.route("/download_bolla")
def download_bolla():
    cliente   = session.get("consegna_cliente", "")
    richieste = json.loads(session.get("consegna_richieste", "[]"))
    if not cliente or not richieste:
        return redirect("/consegne?msg=Nessuna consegna attiva")
    output, errore = _genera_file(cliente, richieste, "bolla")
    if errore:
        return redirect("/conferma_consegna?msg=" + errore)
    return send_file(output, as_attachment=True,
                     download_name=f"Bolla_{cliente}.xlsx")


@app.route("/download_conteggio")
def download_conteggio():
    cliente   = session.get("consegna_cliente", "")
    richieste = json.loads(session.get("consegna_richieste", "[]"))
    if not cliente or not richieste:
        return redirect("/consegne?msg=Nessuna consegna attiva")
    output, errore = _genera_file(cliente, richieste, "conteggio")
    if errore:
        return redirect("/conferma_consegna?msg=" + errore)
    return send_file(output, as_attachment=True,
                     download_name=f"Conteggio_{cliente}.xlsx")


@app.route("/solo_bolla", methods=["POST"])
def solo_bolla():
    cliente = request.form["client"]
    richieste = []
    for i, prodotto in enumerate(clients[cliente]):
        qty = request.form.get(f"qty_{i}")
        if qty and qty.isdigit():
            q = int(qty)
            if q > 0:
                richieste.append((prodotto, q))
    if not richieste:
        return redirect("/consegne?msg=Nessun prodotto selezionato&cliente=" + cliente)
    output, errore = _genera_file(cliente, richieste, "bolla")
    if errore:
        return redirect("/consegne?msg=" + errore + "&cliente=" + cliente)
    return send_file(output, as_attachment=True,
                     download_name=f"Bolla_{cliente}.xlsx")


@app.route("/solo_conteggio", methods=["POST"])
def solo_conteggio():
    cliente = request.form["client"]
    richieste = []
    for i, prodotto in enumerate(clients[cliente]):
        qty = request.form.get(f"qty_{i}")
        if qty and qty.isdigit():
            q = int(qty)
            if q > 0:
                richieste.append((prodotto, q))
    if not richieste:
        return redirect("/consegne?msg=Nessun prodotto selezionato&cliente=" + cliente)
    output, errore = _genera_file(cliente, richieste, "conteggio")
    if errore:
        return redirect("/consegne?msg=" + errore + "&cliente=" + cliente)
    return send_file(output, as_attachment=True,
                     download_name=f"Conteggio_{cliente}.xlsx")


# ==================================================
# START
# ==================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
