from flask import Flask, render_template, request, redirect

app = Flask(__name__)

# -----------------------
# DATI CLIENTI / PRODOTTI
# -----------------------
clients = {
    "Roberto": [
        "Catarratto 2L",
        "Rosato 2L",
        "Merlot 2L",
        "Il Nero 2L",
        "Bianco E.N. 2L",
        "Rosato E.N. 2L",
        "Rosso E.N. 2L",
        "Catarratto 1L",
        "Rosato 1L",
        "Merlot 1L"
    ],
    "Francesco": [
        "Catarratto 2L",
        "Chardonnay 2L",
        "Rosato 2L",
        "Merlot 2L",
        "Syrah 2L",
        "Catarratto 1L",
        "Syrah 1L"
    ],
    "Emanuele": [
        "Catarratto 2L",
        "Rosato 2L",
        "Il Nero 2L",
        "Merlot 2L",
        "Vino Rosso 2L"
    ],
    "Divino": [
        "Bianco 2L",
        "Rosato 2L",
        "Rosso 2L",
        "Syrah 2L"
    ],
    "Pachinos": [
        "Bianco 2L",
        "Rosato 2L",
        "Rosso 2L",
        "Syrah 2L"
    ],
    "Sisa": [
        "Bianco 2L",
        "Rosso 2L"
    ]
}

orders = []

# -----------------------
# HOME
# -----------------------
@app.route("/")
def index():
    return render_template("index.html", clients=clients, orders=orders)

# -----------------------
# CREA ORDINE
# -----------------------
@app.route("/add_order", methods=["POST"])
def add_order():
    client = request.form.get("client")
    product = request.form.get("product")
    qty = request.form.get("qty")

    orders.append({
        "client": client,
        "product": product,
        "qty": qty
    })

    return redirect("/")

# -----------------------
# RUN
# -----------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
