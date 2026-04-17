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

    items = []

    for product in clients[client]:
        qty = request.form.get(product)

        if qty and qty.strip() != "":
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
            "items": items,
            "status": "IN_PRODUZIONE"
        })

    return redirect("/")

@app.route("/toggle/<int:o>/<int:i>")
def toggle(o, i):
    orders[o]["items"][i]["done"] = not orders[o]["items"][i]["done"]
    return redirect("/")

@app.route("/complete/<int:index>")
def complete(index):
    orders[index]["status"] = "PRONTO"
    return redirect("/")

@app.route("/deliver/<int:index>")
def deliver(index):
    deliveries.append(orders[index])
    orders[index]["status"] = "CONSEGNATO"
    return redirect("/")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
