# app.py

from flask import Flask, render_template, request, redirect, send_file
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

app = Flask(__name__)

# =========================
# CLIENTI / PRODOTTI
# =========================

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

# =========================
# MAGAZZINO
# =========================

stock = {}

# =========================
# PRODUZIONE
# =========================

orders = []

# =========================
# MAPPATURA TEMPLATE EXCEL
# Riga prodotto nel template
# =========================

row_map = {
    "Catarratto 2L": 23,
    "Rosato 2L": 25,
    "Merlot 2L": 27,
    "Chardonnay 2L": 29
}

# =========================
# HOME
# =========================

@app.route("/")
def index():

    grouped_stock = {}

    for key, qty in stock.items():

        client, product = key.split(" - ", 1)

        if client not in grouped_stock:
            grouped_stock[client] = {}

        grouped_stock[client][product] = qty

    return render_template(
        "index.html",
        clients=clients,
        stock=grouped_stock,
        orders=orders
    )

# =========================
# NUOVA PRODUZIONE
# =========================

@app.route("/add_order", methods=["POST"])
def add_order():

    client = request.form["client"]
    items = []

    for i, product in enumerate(clients[client]):

        qty = request.form.get(f"qty_{i}")

        if qty and qty.isdigit():

            q = int(qty)

            if q > 0:
                items.append({
                    "product": product,
                    "qty": q,
                    "done": False
                })

    if items:
        orders.append({
            "client": client,
            "items": items
        })

    return redirect("/")

# =========================
# FLAG PRODOTTO FINITO
# =========================

@app.route("/toggle/<int:o>/<int:i>")
def toggle(o, i):

    orders[o]["items"][i]["done"] = not orders[o]["items"][i]["done"]

    return redirect("/")

# =========================
# PASSA A MAGAZZINO
# =========================

@app.route("/complete/<int:o>")
def complete(o):

    client = orders[o]["client"]
    remaining = []

    for item in orders[o]["items"]:

        if item["done"]:

            key = client + " - " + item["product"]

            if key not in stock:
                stock[key] = 0

            stock[key] += item["qty"]

        else:
            remaining.append(item)

    if remaining:
        orders[o]["items"] = remaining
    else:
        orders.pop(o)

    return redirect("/")

# =========================
# CONSEGNA
# =========================

@app.route("/delivery", methods=["POST"])
def delivery():

    client = request.form["client"]

    delivered = {}

    for product in clients[client]:

        key = client + " - " + product

        qty = request.form.get(product)

        if qty and qty.isdigit():

            q = int(qty)

            if q > 0 and key in stock and stock[key] >= q:

                stock[key] -= q
                delivered[product] = q

    app.config["LAST_CLIENT"] = client
    app.config["LAST_ITEMS"] = delivered

    return render_template(
        "download.html",
        client=client
    )

# =========================
# CREA BOLLA
# =========================

@app.route("/download_bolla")
def download_bolla():

    client = app.config["LAST_CLIENT"]
    items = app.config["LAST_ITEMS"]

    wb = load_workbook("BOLLA.xlsx")
    ws = wb.active

    ws["B2"] = client

    for product, qty in items.items():

        row = row_map[product]

        ws[f"B{row}"] = product
        ws[f"G{row}"] = qty

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"BOLLA_{client}_{datetime.today().date()}.xlsx"

    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================
# CREA CONTEGGIO
# =========================

@app.route("/download_conteggio")
def download_conteggio():

    client = app.config["LAST_CLIENT"]
    items = app.config["LAST_ITEMS"]

    wb = load_workbook("CONTEGGIO.xlsx")
    ws = wb.active

    ws["B2"] = client

    for product, qty in items.items():

        row = row_map[product]

        ws[f"B{row}"] = product
        ws[f"G{row}"] = qty

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"CONTEGGIO_{client}_{datetime.today().date()}.xlsx"

    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
