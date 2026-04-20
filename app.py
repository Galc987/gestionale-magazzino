from flask import Flask, render_template, request, redirect
import sqlite3

app = Flask(__name__)

# --------------------
# DATABASE
# --------------------

def db():
    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = db()
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS stock (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente TEXT,
        prodotto TEXT,
        qty INTEGER
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS produzione (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente TEXT,
        prodotto TEXT,
        qty INTEGER,
        done INTEGER DEFAULT 0
    )
    """)

    conn.commit()
    conn.close()

init_db()

# --------------------
# CLIENTI
# --------------------

clients = {
    "Roberto": [
        "Catarratto 2L",
        "Rosato 2L",
        "Merlot 2L"
    ],
    "Francesco": [
        "Catarratto 2L",
        "Chardonnay 2L",
        "Merlot 2L"
    ]
}

# --------------------
# HOME
# --------------------

@app.route("/")
def home():
    return render_template("home.html")

# --------------------
# PRODUZIONE
# --------------------

@app.route("/produzione")
def produzione():

    conn = db()
    rows = conn.execute("SELECT * FROM produzione").fetchall()
    conn.close()

    return render_template(
        "produzione.html",
        clients=clients,
        rows=rows
    )

@app.route("/nuova_produzione", methods=["POST"])
def nuova_produzione():

    cliente = request.form["client"]

    conn = db()

    for i, prodotto in enumerate(clients[cliente]):

        qty = request.form.get(f"qty_{i}")

        if qty and qty.isdigit():

            q = int(qty)

            if q > 0:
                conn.execute(
                    "INSERT INTO produzione(cliente, prodotto, qty) VALUES (?,?,?)",
                    (cliente, prodotto, q)
                )

    conn.commit()
    conn.close()

    return redirect("/produzione")

@app.route("/toggle/<int:id>")
def toggle(id):

    conn = db()

    row = conn.execute(
        "SELECT done FROM produzione WHERE id=?",
        (id,)
    ).fetchone()

    nuovo = 0 if row["done"] == 1 else 1

    conn.execute(
        "UPDATE produzione SET done=? WHERE id=?",
        (nuovo, id)
    )

    conn.commit()
    conn.close()

    return redirect("/produzione")

@app.route("/passa_magazzino")
def passa_magazzino():

    conn = db()

    finiti = conn.execute(
        "SELECT * FROM produzione WHERE done=1"
    ).fetchall()

    for r in finiti:

        trovato = conn.execute(
            "SELECT * FROM stock WHERE cliente=? AND prodotto=?",
            (r["cliente"], r["prodotto"])
        ).fetchone()

        if trovato:

            nuova = trovato["qty"] + r["qty"]

            conn.execute(
                "UPDATE stock SET qty=? WHERE id=?",
                (nuova, trovato["id"])
            )

        else:

            conn.execute(
                "INSERT INTO stock(cliente, prodotto, qty) VALUES (?,?,?)",
                (r["cliente"], r["prodotto"], r["qty"])
            )

        conn.execute(
            "DELETE FROM produzione WHERE id=?",
            (r["id"],)
        )

    conn.commit()
    conn.close()

    return redirect("/produzione")

# --------------------
# MAGAZZINO
# --------------------

@app.route("/magazzino")
def magazzino():

    conn = db()

    rows = conn.execute(
        "SELECT * FROM stock ORDER BY cliente"
    ).fetchall()

    conn.close()

    return render_template(
        "magazzino.html",
        rows=rows
    )

# --------------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
