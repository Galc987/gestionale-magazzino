from flask import Flask, render_template, request, redirect, send_file
import os
import psycopg2
from psycopg2.extras import RealDictCursor
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)

DATABASE_URL = os.environ.get("DATABASE_URL")

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
MODELLI_DIR = os.path.join(BASE_DIR, "modelli")


def db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


def init_db():
    conn = db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS stock (
            id SERIAL PRIMARY KEY,
            cliente TEXT,
            prodotto TEXT,
            qty INTEGER
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS produzione (
            id SERIAL PRIMARY KEY,
            cliente TEXT,
            prodotto TEXT,
            qty INTEGER,
            done INTEGER DEFAULT 0
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS storico (
            id SERIAL PRIMARY KEY,
            cliente TEXT,
            prodotto TEXT,
            qty INTEGER,
            tipo TEXT,
            data TIMESTAMP DEFAULT NOW()
        )
    """)
    conn.commit()
    cur.close()
    conn.close()


init_db()

# --------------------------------------------------
# DATI CLIENTI
# --------------------------------------------------
clients = {
    "Roberto":   ["Catarratto 2L", "Rosato 2L", "Merlot 2L"],
    "Francesco": ["Catarratto 2L", "Chardonnay 2L", "Merlot 2L"]
}

# Mappa prodotto app -> riga nel template Excel di Roberto
# La colonna G di quella riga contiene i fardelli (bottiglie / moltiplicatore)
# Il moltiplicatore (bt per fardello) è già nelle formule del template
ROBERTO_RIGHE = {
    "Catarratto 2L": 23,
    "Rosato 2L":     25,
    "Merlot 2L":     27,
}
ROBERTO_MOLT = {
    "Catarratto 2L": 6,
    "Rosato 2L":     6,
    "Merlot 2L":     6,
}


# --------------------------------------------------
# HOME
# --------------------------------------------------
@app.route("/")
def home():
    return render_template("home.html")


# --------------------------------------------------
# STORICO
# --------------------------------------------------
@app.route("/storico")
def storico():
    conn = db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM storico ORDER BY data DESC LIMIT 300")
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return render_template("storico.html", rows=rows)


# --------------------------------------------------
# PRODUZIONE
# --------------------------------------------------
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
            cur.execute(
                "UPDATE stock SET qty=%s WHERE id=%s",
                (ex["qty"] + r["qty"], ex["id"])
            )
        else:
            cur.execute(
                "INSERT INTO stock(cliente, prodotto, qty) VALUES(%s,%s,%s)",
                (r["cliente"], r["prodotto"], r["qty"])
            )
        cur.execute(
            "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
            (r["cliente"], r["prodotto"], r["qty"], "Passato a Magazzino")
        )
        cur.execute("DELETE FROM produzione WHERE id=%s", (r["id"],))
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/produzione")


# --------------------------------------------------
# MAGAZZINO
# --------------------------------------------------
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
    return render_template(
        "magazzino.html",
        grouped=grouped,
        clients=clients,
        msg=msg,
        cliente_sel=cliente_sel
    )


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
        cur.execute(
            "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
            (cliente, prodotto)
        )
        row = cur.fetchone()
        if not row:
            cur.close()
            conn.close()
            return redirect("/magazzino?msg=" + prodotto + " non presente&cliente=" + cliente)
        if row["qty"] < q:
            cur.close()
            conn.close()
            return redirect("/magazzino?msg=" + prodotto + " quantita insufficiente&cliente=" + cliente)
    for prodotto, q in richieste:
        cur.execute(
            "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
            (cliente, prodotto)
        )
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


# --------------------------------------------------
# CONSEGNE — pagina principale
# --------------------------------------------------
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

    return render_template(
        "consegne.html",
        clients=clients,
        grouped=grouped,
        cliente_sel=cliente_sel,
        msg=msg
    )


# --------------------------------------------------
# FUNZIONE INTERNA: scarica magazzino e genera Excel
# --------------------------------------------------
def _esegui_consegna_db(cliente, richieste, tipo_storico, cur):
    """
    Scala il magazzino e registra nello storico.
    richieste: lista di (prodotto, qty_bottiglie)
    Ritorna errore stringa o None se ok.
    """
    # Controllo disponibilita
    for prodotto, q in richieste:
        cur.execute(
            "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
            (cliente, prodotto)
        )
        row = cur.fetchone()
        if not row:
            return prodotto + " non presente in magazzino"
        if row["qty"] < q:
            return prodotto + " quantita insufficiente"

    # Scarico
    for prodotto, q in richieste:
        cur.execute(
            "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
            (cliente, prodotto)
        )
        row = cur.fetchone()
        nuova = row["qty"] - q
        if nuova == 0:
            cur.execute("DELETE FROM stock WHERE id=%s", (row["id"],))
        else:
            cur.execute("UPDATE stock SET qty=%s WHERE id=%s", (nuova, row["id"]))
        cur.execute(
            "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
            (cliente, prodotto, q, tipo_storico)
        )
    return None


def _aggiorna_template_roberto(ws, richieste):
    """
    Azzera tutti i fardelli nel template e inserisce quelli della consegna.
    IMPORTANTE: scrive solo nella colonna G che e' la cella top-left dei merge.
    """
    # Azzera tutte le righe prodotto
    for riga in [23, 25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45, 47]:
        ws[f"G{riga}"] = 0

    # Inserisce fardelli per i prodotti della consegna
    for prodotto, qty_bt in richieste:
        if prodotto in ROBERTO_RIGHE:
            riga = ROBERTO_RIGHE[prodotto]
            molt = ROBERTO_MOLT[prodotto]
            ws[f"G{riga}"] = qty_bt / molt


# --------------------------------------------------
# GENERA BOLLA
# --------------------------------------------------
@app.route("/genera_bolla", methods=["POST"])
def genera_bolla():
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

    if cliente != "Roberto":
        return redirect("/consegne?msg=Bolla disponibile solo per Roberto&cliente=" + cliente)

    file_modello = os.path.join(MODELLI_DIR, "bolla_roberto.xlsx")
    if not os.path.exists(file_modello):
        return redirect("/consegne?msg=File modello bolla mancante&cliente=" + cliente)

    conn = db()
    cur = conn.cursor()
    errore = _esegui_consegna_db(cliente, richieste, "Bolla Generata", cur)
    if errore:
        cur.close()
        conn.close()
        return redirect("/consegne?msg=" + errore + "&cliente=" + cliente)
    conn.commit()
    cur.close()
    conn.close()

    # Genera file Excel
    wb = load_workbook(file_modello)
    ws = wb.active

    # Lascia vuoti numero e data (scritti a penna)
    # H2 e' la cella top-left del merge H2:K5
    ws["H2"] = "DOCUMENTO DI TRASPORTO\nN.          DEL\n"

    # Data ritiro vuota — F55 e' top-left di F55:G58
    ws["F55"] = "DATA RITIRO\n\n\n"

    # Aggiorna fardelli
    _aggiorna_template_roberto(ws, richieste)

    output = os.path.join(BASE_DIR, "bolla_generata.xlsx")
    wb.save(output)

    return send_file(output, as_attachment=True, download_name="Bolla_Roberto.xlsx")


# --------------------------------------------------
# GENERA CONTEGGIO
# --------------------------------------------------
@app.route("/genera_conteggio", methods=["POST"])
def genera_conteggio():
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

    if cliente != "Roberto":
        return redirect("/consegne?msg=Conteggio disponibile solo per Roberto&cliente=" + cliente)

    file_modello = os.path.join(MODELLI_DIR, "conteggio_roberto.xlsx")
    if not os.path.exists(file_modello):
        return redirect("/consegne?msg=File modello conteggio mancante&cliente=" + cliente)

    conn = db()
    cur = conn.cursor()
    errore = _esegui_consegna_db(cliente, richieste, "Conteggio Generato", cur)
    if errore:
        cur.close()
        conn.close()
        return redirect("/consegne?msg=" + errore + "&cliente=" + cliente)
    conn.commit()
    cur.close()
    conn.close()

    # Genera file Excel
    wb = load_workbook(file_modello)
    ws = wb.active

    # G2 e' top-left del merge nel conteggio
    ws["G2"] = "DOCUMENTO DI TRASPORTO\nN.          DEL\n"

    # Data ritiro vuota — F53 e' top-left di F53:G56 nel conteggio
    ws["F53"] = "DATA RITIRO\n\n\n"

    # Aggiorna fardelli (stessa logica, stesse righe)
    _aggiorna_template_roberto(ws, richieste)

    output = os.path.join(BASE_DIR, "conteggio_generato.xlsx")
    wb.save(output)

    return send_file(output, as_attachment=True, download_name="Conteggio_Roberto.xlsx")


# --------------------------------------------------
# START
# --------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
