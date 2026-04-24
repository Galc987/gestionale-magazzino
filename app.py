from flask import Flask, render_template, request, redirect, send_file
import os
import psycopg2
import shutil
from psycopg2.extras import RealDictCursor
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)

DATABASE_URL = os.environ.get("DATABASE_URL")

# --------------------------------------------------
# PATH BASE SICURO PER RENDER / LINUX
# --------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODELLI_DIR = os.path.join(BASE_DIR, "modelli")


def db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


# --------------------------------------------------
# INIT DB
# --------------------------------------------------
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
    "Roberto": ["Catarratto 2L", "Rosato 2L", "Merlot 2L"],
    "Francesco": ["Catarratto 2L", "Chardonnay 2L", "Merlot 2L"]
}

nomi_roberto = {
    "Catarratto 2L": "PET VINO CATARRATTO LITRI 2",
    "Rosato 2L": "PET VINO ROSATO LITRI 2",
    "Merlot 2L": "PET VINO MERLOT LITRI 2"
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

    cur.execute(
        "UPDATE produzione SET done=%s WHERE id=%s",
        (nuovo, id)
    )

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

        cur.execute(
            "DELETE FROM produzione WHERE id=%s",
            (r["id"],)
        )

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
        return redirect("/magazzino?msg=Nessun prodotto selezionato")

    conn = db()
    cur = conn.cursor()

    # controllo prima
    for prodotto, q in richieste:
        cur.execute(
            "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
            (cliente, prodotto)
        )

        row = cur.fetchone()

        if not row or row["qty"] < q:
            cur.close()
            conn.close()
            return redirect("/magazzino?msg=Quantita insufficiente")

    # scarico
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
            cur.execute(
                "UPDATE stock SET qty=%s WHERE id=%s",
                (nuova, row["id"])
            )

        cur.execute(
            "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
            (cliente, prodotto, q, "Scarico Magazzino")
        )

    conn.commit()
    cur.close()
    conn.close()

    return redirect("/magazzino?msg=Scarico completato")


# --------------------------------------------------
# CONSEGNE
# --------------------------------------------------
@app.route("/consegne")
def consegne():
    return render_template(
        "consegne.html",
        products=clients["Roberto"],
        msg=request.args.get("msg", "")
    )


# --------------------------------------------------
# BOLLA
# --------------------------------------------------
@app.route("/genera_bolla", methods=["POST"])
def genera_bolla():

    cliente = request.form["client"]

    if cliente != "Roberto":
        return redirect("/consegne?msg=Bolla attiva solo per Roberto")

    conn = db()
    cur = conn.cursor()

    righe = []

    for i, prodotto in enumerate(clients["Roberto"]):
        qty = request.form.get(f"qty_{i}")

        if qty and qty.isdigit():
            q = int(qty)

            if q > 0:
                cur.execute(
                    "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
                    ("Roberto", prodotto)
                )

                row = cur.fetchone()

                if not row or row["qty"] < q:
                    cur.close()
                    conn.close()
                    return redirect("/consegne?msg=Stock insufficiente")

                righe.append((prodotto, q, row["id"], row["qty"]))

    if not righe:
        cur.close()
        conn.close()
        return redirect("/consegne?msg=Nessun prodotto selezionato")

    # FILE MODELLO
    file_path = os.path.join(MODELLI_DIR, "bolla_roberto.xlsx")

    if not os.path.exists(file_path):
        cur.close()
        conn.close()
        return redirect("/consegne?msg=File bolla mancante")

    wb = load_workbook(file_path)
    ws = wb.active

    # numero documento
    ws["I3"] = "N. 40/2026 DEL " + datetime.now().strftime("%d/%m/%Y")

    riga = 24
    totale_colli = 0

    for prodotto, q, stock_id, old_qty in righe:

        ws[f"B{riga}"] = nomi_roberto[prodotto]
        ws[f"G{riga}"] = q
        ws[f"K{riga}"] = q * 6

        totale_colli += q

        nuova = old_qty - q

        if nuova == 0:
            cur.execute("DELETE FROM stock WHERE id=%s", (stock_id,))
        else:
            cur.execute(
                "UPDATE stock SET qty=%s WHERE id=%s",
                (nuova, stock_id)
            )

        cur.execute(
            "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
            ("Roberto", prodotto, q, "Bolla Generata")
        )

        riga += 1

    ws["G50"] = totale_colli
    ws["H59"] = totale_colli

    conn.commit()
    cur.close()
    conn.close()

    output = os.path.join(BASE_DIR, "bolla_generata.xlsx")
    wb.save(output)

    return send_file(output, as_attachment=True)


# --------------------------------------------------
# CONTEGGIO
# --------------------------------------------------
@app.route("/genera_conteggio", methods=["POST"])
def genera_conteggio():

    richieste = []

    for i, prodotto in enumerate(clients["Roberto"]):
        qty = request.form.get(f"qty_{i}")

        if qty and qty.isdigit():
            q = int(qty)

            if q > 0:
                richieste.append((prodotto, q))

    if not richieste:
        return redirect("/consegne?msg=Nessun prodotto selezionato")

    conn = db()
    cur = conn.cursor()

    # controllo stock
    for prodotto, q in richieste:
        cur.execute(
            "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
            ("Roberto", prodotto)
        )

        row = cur.fetchone()

        if not row or row["qty"] < q:
            cur.close()
            conn.close()
            return redirect("/consegne?msg=Stock insufficiente")

    # modello
    file_path = os.path.join(MODELLI_DIR, "conteggio_roberto.xlsx")

    if not os.path.exists(file_path):
        cur.close()
        conn.close()
        return redirect("/consegne?msg=File conteggio mancante")

    wb = load_workbook(file_path)
    ws = wb.active

    riga = 15

    for prodotto, q in richieste:

        ws[f"A{riga}"] = nomi_roberto[prodotto]
        ws[f"D{riga}"] = q
        ws[f"E{riga}"] = q * 6

        cur.execute(
            "SELECT * FROM stock WHERE cliente=%s AND prodotto=%s",
            ("Roberto", prodotto)
        )

        row = cur.fetchone()

        nuova = row["qty"] - q

        if nuova == 0:
            cur.execute("DELETE FROM stock WHERE id=%s", (row["id"],))
        else:
            cur.execute(
                "UPDATE stock SET qty=%s WHERE id=%s",
                (nuova, row["id"])
            )

        cur.execute(
            "INSERT INTO storico(cliente, prodotto, qty, tipo) VALUES(%s,%s,%s,%s)",
            ("Roberto", prodotto, q, "Conteggio Generato")
        )

        riga += 1

    conn.commit()
    cur.close()
    conn.close()

    output = os.path.join(BASE_DIR, "conteggio_generato.xlsx")
    wb.save(output)

    return send_file(output, as_attachment=True)


# --------------------------------------------------
# START
# --------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
