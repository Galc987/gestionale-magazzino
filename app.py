from flask import Flask, render_template, request, redirect, send_file, session
import os
import psycopg2
from psycopg2.extras import RealDictCursor
from openpyxl import load_workbook
import json

app = Flask(__name__)
app.secret_key = "lc_wine_secret_2026"

DATABASE_URL = os.environ.get("DATABASE_URL")
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
MODELLI_DIR  = os.path.join(BASE_DIR, "modelli")

# Bottiglie per pedana
BT_PER_PEDANA_2L = 889
BT_PER_PEDANA_1L = 1344


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
            id SERIAL PRIMARY KEY, testo TEXT, data TIMESTAMP DEFAULT NOW()
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS materie_prime (
            id SERIAL PRIMARY KEY, cliente TEXT, materiale TEXT,
            qty INTEGER DEFAULT 0, soglia_minima INTEGER DEFAULT 0
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS storico_mp (
            id SERIAL PRIMARY KEY, cliente TEXT, materiale TEXT,
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

MOLTIPLICATORI = {
    "Roberto": {
        "Catarratto 2L": 6,  "Rosato 2L": 6,      "Merlot 2L": 6,
        "Il Nero 2L": 6,     "Bianco E.N. 2L": 6,  "Rosato E.N. 2L": 6,
        "Rosso E.N. 2L": 6,  "Catarratto 1L": 12,  "Rosato 1L": 12,
        "Merlot 1L": 12,     "Bianco S.E. 1L": 16, "Rosato S.E. 1L": 16,
        "Rosso S.E. 1L": 16,
    },
    "Francesco": {
        "Catarratto 2L": 6, "Chardonnay 2L": 6, "Rosato 2L": 6,
        "Merlot 2L": 6,     "Syrah 2L": 6,
        "Catarratto 1L": 12, "Syrah 1L": 12,
    },
    "Emanuele": {
        "Catarratto 2L": 9,      "Rosato 2L": 9,      "Il Nero 2L": 9,
        "Merlot 2L": 9,          "Vino Rosso 2L": 9,  "Syrah 2L": 9,
        "Bianco S.E. 2L": 9,     "Rosato S.E. 2L": 9, "Rosso S.E. 2L": 9,
        "Catarratto R.B. 2L": 9, "Rosato R.B. 2L": 9, "Il Nero R.B. 2L": 9,
        "Vino Rosso R.B. 2L": 9, "Catarratto 1L": 16, "Rosato 1L": 16,
        "Il Nero 1L": 16,        "Bianco S.E. 1L": 16,"Rosato S.E. 1L": 16,
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

# Materie prime per cliente
# Bottiglie vuote: globali (non per cliente)
BOTTIGLIE_GLOBALI = ["Bottiglie 2L vuote", "Bottiglie 1L vuote"]

# Soglie minime fisse
SOGLIE_MP = {
    "Bottiglie 2L vuote": 8000,
    "Bottiglie 1L vuote": 1800,
}
SOGLIA_ETICHETTE = 1000

# Etichette per cliente
def _build_etichette(cliente):
    return [f"Etichetta {p}" for p in clients[cliente]]

ETICHETTE_CLIENTI = {c: _build_etichette(c) for c in clients}

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
            "Catarratto 2L": 21,      "Rosato 2L": 23,      "Il Nero 2L": 25,
            "Vino Rosso 2L": 27,      "Syrah 2L": 29,       "Merlot 2L": 31,
            "Catarratto 1L": 33,      "Rosato 1L": 35,      "Il Nero 1L": 37,
            "Catarratto R.B. 2L": 41, "Rosato R.B. 2L": 43, "Il Nero R.B. 2L": 45,
            "Vino Rosso R.B. 2L": 47, "Bianco S.E. 2L": 51, "Rosato S.E. 2L": 53,
            "Rosso S.E. 2L": 55,      "Bianco S.E. 1L": 57, "Rosato S.E. 1L": 59,
            "Rosso S.E. 1L": 61,
        },
        "prodotti_conteggio": {
            "Catarratto 2L": 25,      "Rosato 2L": 27,      "Il Nero 2L": 29,
            "Vino Rosso 2L": 31,      "Syrah 2L": 33,       "Merlot 2L": 35,
            "Catarratto 1L": 37,      "Rosato 1L": 39,      "Il Nero 1L": 41,
            "Catarratto R.B. 2L": 45, "Rosato R.B. 2L": 47, "Il Nero R.B. 2L": 49,
            "Vino Rosso R.B. 2L": 51, "Bianco S.E. 2L": 55, "Rosato S.E. 2L": 57,
            "Rosso S.E. 2L": 59,      "Bianco S.E. 1L": 61, "Rosato S.E. 1L": 63,
            "Rosso S.E. 1L": 65,
        },
        "righe_az_bolla":     [21,23,25,27,29,31,33,35,37,41,43,45,47,51,53,55,57,59,61],
        "righe_az_conteggio": [25,27,29,31,33,35,37,39,41,45,47,49,51,55,57,59,61,63,65],
        "cella_titolo_bolla": "H2",     "cella_data_bolla": "F67",
        "cella_titolo_conteggio": "H2", "cella_data_conteggio": "F71",
    },
    "Mazzarrone": {
        "bolla": None, "conteggio": None, "prodotti": {}, "righe_az": [],
        "cella_titolo_bolla": None, "cella_data_bolla": None,
        "cella_titolo_conteggio": None, "cella_data_conteggio": None,
    },
    "Sisa": {
        "bolla": None, "conteggio": None, "prodotti": {}, "righe_az": [],
        "cella_titolo_bolla": None, "cella_data_bolla": None,
        "cella_titolo_conteggio": None, "cella_data_conteggio": None,
    },
}


# ==================================================
# FUNZIONI EXCEL
# ==================================================

def _aggiorna_excel(ws, cliente, richieste_fardelli, tipo="bolla"):
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
    output = os.path.join(BASE_DIR, f"{tipo}_generato_{cliente.lower()}.xlsx")
    wb.save(output)
    return output, None


def _leggi_richieste_fardelli(cliente, form):
    richieste = []
    for i, prodotto in enumerate(clients[cliente]):
        val = form.get(f"qty_{i}")
        if val and val.isdigit():
            f = int(val)
            if f > 0:
                richieste.append((prodotto, f))
    return richieste


def _fardelli_a_bottiglie(cliente, richieste_fardelli):
    molt = MOLTIPLICATORI.get(cliente, {})
    return [(p, f * molt.get(p, 1)) for p, f in richieste_fardelli]


def _is_2L(prodotto):
    return "2L" in prodotto or "2l" in prodotto

def _is_1L(prodotto):
    return "1L" in prodotto or "1l" in prodotto


def _scarico_automatico_bottiglie(cur, cliente, prodotti_bottiglie):
    """
    Scarica automaticamente:
    - Bottiglie vuote: globali (cliente="GLOBALE"), aggregate per formato
    - Etichette: per cliente, una per bottiglia per ogni prodotto
    prodotti_bottiglie = [(prodotto, qty_bt), ...]
    """
    # --- Bottiglie vuote GLOBALI ---
    bt2L = sum(q for p, q in prodotti_bottiglie if _is_2L(p))
    bt1L = sum(q for p, q in prodotti_bottiglie if _is_1L(p))

    for materiale, qty_usate in [("Bottiglie 2L vuote", bt2L), ("Bottiglie 1L vuote", bt1L)]:
        if qty_usate <= 0:
            continue
        cur.execute(
            "SELECT * FROM materie_prime WHERE cliente=%s AND materiale=%s",
            ("GLOBALE", materiale)
        )
        row = cur.fetchone()
        if row and row["qty"] >= qty_usate:
            cur.execute("UPDATE materie_prime SET qty=%s WHERE id=%s",
                        (row["qty"] - qty_usate, row["id"]))
            cur.execute(
                "INSERT INTO storico_mp(cliente, materiale, qty, tipo) VALUES(%s,%s,%s,%s)",
                ("GLOBALE", materiale, qty_usate, "Scarico automatico produzione")
            )

    # --- Etichette per CLIENTE ---
    for prodotto, qty_bt in prodotti_bottiglie:
        materiale = f"Etichetta {prodotto}"
        cur.execute(
            "SELECT * FROM materie_prime WHERE cliente=%s AND materiale=%s",
            (cliente, materiale)
        )
        row = cur.fetchone()
        if row and row["qty"] >= qty_bt:
            cur.execute("UPDATE materie_prime SET qty=%s WHERE id=%s",
                        (row["qty"] - qty_bt, row["id"]))
            cur.execute(
                "INSERT INTO storico_mp(cliente, materiale, qty, tipo) VALUES(%s,%s,%s,%s)",
                (cliente, materiale, qty_bt, "Scarico automatico produzione")
            )


def _get_alert_mp():
    """Ritorna lista di dict con tutti gli alert sotto soglia (soglie fisse)."""
    try:
        conn = db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM materie_prime ORDER BY cliente, materiale")
        rows = cur.fetchall()
        cur.close()
        conn.close()
        alert = []
        for r in rows:
            soglia = SOGLIE_MP.get(r["materiale"], SOGLIA_ETICHETTE)
            if r["qty"] <= soglia:
                alert.append({
                    "cliente":   r["cliente"],
                    "materiale": r["materiale"],
                    "qty":       r["qty"],
                    "soglia":    soglia,
                })
        return alert
    except:
        return []


def _conta_alert_mp():
    return len(_get_alert_mp())


# ==================================================
# HOME
# ==================================================
@app.route("/")
def home():
    return render_template("home.html", alert_mp=_get_alert_mp())


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
    return render_template("produzione.html", clients=clients, rows=righe,
                           note=note, moltiplicatori=MOLTIPLICATORI)


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

    # Raggruppa per cliente per lo scarico bottiglie
    per_cliente = {}
    for r in finiti:
        per_cliente.setdefault(r["cliente"], []).append((r["prodotto"], r["qty"]))

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

    # Scarico automatico bottiglie vuote per cliente
    for cliente, prodotti_bt in per_cliente.items():
        _scarico_automatico_bottiglie(cur, cliente, prodotti_bt)

    conn.commit()
    cur.close()
    conn.close()
    return redirect("/produzione")


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
# MAGAZZINO PRODOTTI FINITI
# ==================================================
@app.route("/magazzino")
def magazzino():
    msg = request.args.get("msg", "")
    cliente_sel = request.args.get("cliente", "")
    conn = db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM stock WHERE qty > 0 ORDER BY cliente, prodotto")
    stock_rows = cur.fetchall()
    cur.execute("SELECT * FROM materie_prime ORDER BY cliente, materiale")
    mp_rows = cur.fetchall()
    cur.close()
    conn.close()

    # Stock prodotti finiti per cliente
    grouped = {}
    for r in stock_rows:
        grouped.setdefault(r["cliente"], []).append(r)

    # Materie prime per cliente + alert
    grouped_mp = {}
    alert_mp = []
    for r in mp_rows:
        grouped_mp.setdefault(r["cliente"], []).append(r)
        if r["soglia_minima"] > 0 and r["qty"] <= r["soglia_minima"]:
            alert_mp.append(r)

    # Separa giacenze bottiglie globali da etichette per cliente
    bottiglie_globali = grouped_mp.get("GLOBALE", [])
    grouped_etichette = {k: v for k, v in grouped_mp.items() if k != "GLOBALE"}

    return render_template(
        "magazzino.html",
        grouped=grouped,
        bottiglie_globali=bottiglie_globali,
        grouped_etichette=grouped_etichette,
        alert_mp=_get_alert_mp(),
        clients=clients,
        etichette_clienti=ETICHETTE_CLIENTI,
        msg=msg,
        cliente_sel=cliente_sel,
        moltiplicatori=MOLTIPLICATORI,
        bt_pedana_2l=BT_PER_PEDANA_2L,
        bt_pedana_1l=BT_PER_PEDANA_1L,
    )


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


# Materie prime: inizializza, carico, scarico, soglia
@app.route("/init_materie_prime")
def init_materie_prime():
    conn = db()
    cur = conn.cursor()
    # Bottiglie vuote: globali
    for materiale in BOTTIGLIE_GLOBALI:
        cur.execute(
            "SELECT id FROM materie_prime WHERE cliente=%s AND materiale=%s",
            ("GLOBALE", materiale)
        )
        if not cur.fetchone():
            cur.execute(
                "INSERT INTO materie_prime(cliente, materiale, qty, soglia_minima) VALUES(%s,%s,0,0)",
                ("GLOBALE", materiale)
            )
    # Etichette: per cliente
    for cliente, etichette in ETICHETTE_CLIENTI.items():
        for materiale in etichette:
            cur.execute(
                "SELECT id FROM materie_prime WHERE cliente=%s AND materiale=%s",
                (cliente, materiale)
            )
            if not cur.fetchone():
                cur.execute(
                    "INSERT INTO materie_prime(cliente, materiale, qty, soglia_minima) VALUES(%s,%s,0,0)",
                    (cliente, materiale)
                )
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/magazzino?msg=Materie prime inizializzate")


@app.route("/carico_mp", methods=["POST"])
def carico_mp():
    materiale = request.form["materiale"]
    # Bottiglie vuote sono globali
    cliente = "GLOBALE" if materiale in BOTTIGLIE_GLOBALI else request.form["cliente"]
    val       = request.form.get("qty", "0")
    unita     = request.form.get("unita", "pezzi")  # "pedane" o "pezzi"
    if not val.isdigit() or int(val) <= 0:
        return redirect("/magazzino?msg=Quantita non valida&cliente=" + cliente)
    q = int(val)
    # Converti pedane in bottiglie se necessario
    if unita == "pedane":
        if "2L" in materiale:
            q = q * BT_PER_PEDANA_2L
        elif "1L" in materiale:
            q = q * BT_PER_PEDANA_1L
    conn = db()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM materie_prime WHERE cliente=%s AND materiale=%s",
        (cliente, materiale)
    )
    row = cur.fetchone()
    if row:
        cur.execute("UPDATE materie_prime SET qty=%s WHERE id=%s", (row["qty"] + q, row["id"]))
    else:
        cur.execute(
            "INSERT INTO materie_prime(cliente, materiale, qty, soglia_minima) VALUES(%s,%s,%s,0)",
            (cliente, materiale, q)
        )
    cur.execute(
        "INSERT INTO storico_mp(cliente, materiale, qty, tipo) VALUES(%s,%s,%s,%s)",
        (cliente, materiale, q, "Carico")
    )
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/magazzino?msg=Carico registrato&cliente=" + cliente)


@app.route("/scarico_mp", methods=["POST"])
def scarico_mp():
    materiale = request.form["materiale"]
    cliente = "GLOBALE" if materiale in BOTTIGLIE_GLOBALI else request.form["cliente"]
    val       = request.form.get("qty", "0")
    if not val.isdigit() or int(val) <= 0:
        return redirect("/magazzino?msg=Quantita non valida&cliente=" + cliente)
    q = int(val)
    conn = db()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM materie_prime WHERE cliente=%s AND materiale=%s",
        (cliente, materiale)
    )
    row = cur.fetchone()
    if not row or row["qty"] < q:
        cur.close(); conn.close()
        return redirect("/magazzino?msg=" + materiale + " quantita insufficiente&cliente=" + cliente)
    cur.execute("UPDATE materie_prime SET qty=%s WHERE id=%s", (row["qty"] - q, row["id"]))
    cur.execute(
        "INSERT INTO storico_mp(cliente, materiale, qty, tipo) VALUES(%s,%s,%s,%s)",
        (cliente, materiale, q, "Scarico")
    )
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/magazzino?msg=Scarico registrato&cliente=" + cliente)


@app.route("/set_soglia_mp", methods=["POST"])
def set_soglia_mp():
    cliente   = request.form["cliente"]
    materiale = request.form["materiale"]
    val       = request.form.get("soglia", "0")
    if not val.isdigit():
        return redirect("/magazzino?msg=Soglia non valida&cliente=" + cliente)
    conn = db()
    cur = conn.cursor()
    cur.execute(
        "UPDATE materie_prime SET soglia_minima=%s WHERE cliente=%s AND materiale=%s",
        (int(val), cliente, materiale)
    )
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/magazzino?msg=Soglia aggiornata&cliente=" + cliente)


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
    session["consegna_cliente"]   = cliente
    session["consegna_fardelli"]  = json.dumps(richieste_f)
    session["consegna_bottiglie"] = json.dumps(richieste_bt)
    return redirect("/conferma_consegna")


@app.route("/conferma_consegna")
def conferma_consegna():
    cliente      = session.get("consegna_cliente", "")
    richieste_f  = json.loads(session.get("consegna_fardelli", "[]"))
    richieste_bt = json.loads(session.get("consegna_bottiglie", "[]"))
    if not cliente or not richieste_f:
        return redirect("/consegne?msg=Nessuna consegna attiva")
    return render_template("conferma_consegna.html", cliente=cliente,
                           richieste_f=richieste_f, richieste_bt=richieste_bt)


@app.route("/download_bolla")
def download_bolla():
    cliente     = session.get("consegna_cliente", "")
    richieste_f = json.loads(session.get("consegna_fardelli", "[]"))
    if not cliente or not richieste_f:
        return redirect("/consegne?msg=Nessuna consegna attiva")
    output, errore = _genera_file(cliente, richieste_f, "bolla")
    if errore:
        return redirect("/conferma_consegna?msg=" + errore)
    return send_file(output, as_attachment=True, download_name=f"Bolla_{cliente}.xlsx")


@app.route("/download_conteggio")
def download_conteggio():
    cliente     = session.get("consegna_cliente", "")
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
