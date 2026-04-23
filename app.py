from flask import Flask, render_template, request, redirect
import os
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)

DATABASE_URL = os.environ.get("DATABASE_URL")

def db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

# -----------------------
# INIT DATABASE
# -----------------------

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

# -----------------------
# CLIENTI
# -----------------------

clients = {
    "Roberto": ["Catarratto 2L", "Rosato 2L", "Merlot 2L"],
    "Francesco": ["Catarratto 2L", "Chardonnay 2L", "Merlot 2L"]
}

# -----------------------
# HOME
# -----------------------

@app.route("/")
def home():
    return render_template("home.html")

# -----------------------
# STORICO
# -----------------------

@app.route("/storico")
def storico():

    conn = db()
    cur = conn.cursor()

    cur.execute("""
        SELECT * FROM storico
        ORDER BY data DESC
        LIMIT 300
    """)

    rows = cur.fetchall()

    cur.close()
    conn.close()

    return render_template("storico.html", rows=rows)

# -----------------------
# PRODUZIONE
# -----------------------

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
            nuova = ex["qty"] + r["qty"]

            cur.execute(
                "UPDATE stock SET qty=%s WHERE id=%s",
                (nuova, ex["id"])
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

# -----------------------
# MAGAZZINO
# -----------------------

@app.route("/magazzino")
def magazzino():

    msg = request.args.get("msg", "")
    cliente_sel = request.args.get("cliente", "")

    conn = db()
    cur = conn.cursor()

    cur.execute("""
        SELECT * FROM stock
        WHERE qty > 0
        ORDER BY cliente, prodotto
    """)

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

    # CONTROLLO TOTALE PRIMA DI SCARICARE
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
            return redirect("/magazzino?msg=" + prodotto + " quantità insufficiente&cliente=" + cliente)

    # SOLO SE TUTTO OK SCARICA
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

    return redirect("/magazzino?msg=Scarico completato&cliente=" + cliente)

# -----------------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
    return redirect("/magazzino?msg=Scarico completato&cliente=" + cliente)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
