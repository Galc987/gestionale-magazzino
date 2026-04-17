from flask import Flask, render_template, request, redirect

app = Flask(__name__)

clients = {
    "Roberto": ["Catarratto 2L", "Rosato 2L", "Merlot 2L"],
    "Francesco": ["Catarratto 2L", "Chardonnay 2L"]
}

stock = {
    "Catarratto 2L": 100,
    "Rosato 2L": 80,
    "Merlot 2L": 50,
    "Chardonnay 2L": 30
}

orders = []

deliveries = []

# ---------------- ORDINI / PRODUZIONE ----------------
@app.route("/")
def index():
    return render_template(
        "index.html",
        clients=clients,
        stock=stock,
        orders=orders,
        deliveries=deliveries
    )

@app.route("/add_order", methods=["POST"])
def add_order():
    client = request.form.get("client")

    products = request.form.getlist("product")
    qtys = request.form.getlist("qty")

    items = []

    for p, q in zip(products, qtys):
        if q and int(q) > 0:
            items.append({
                "product": p,
                "qty": int(q),
                "done": False
            })

    orders.append({
        "client": client,
        "items": items,
        "status": "IN_PRODUZIONE"
    })

    return redirect("/")

# ---------------- TOGGLE PRODOTTO SINGOLO ----------------
@app.route("/toggle_item/<int:o>/<int:i>")
def toggle_item(o, i):
    orders[o]["items"][i]["done"] = not orders[o]["items"][i]["done"]
    return redirect("/")

# ---------------- COMPLETA ORDINE ----------------
@app.route("/complete_order/<int:index>")
def complete_order(index):
    orders[index]["status"] = "PRONTO_MAGAZZINO"
    return redirect("/")

# ---------------- CONSEGNA ----------------
@app.route("/deliver/<int:index>")
def deliver(index):
    deliveries.append(orders[index])
    orders[index]["status"] = "CONSEGNATO"
    return redirect("/")

# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
