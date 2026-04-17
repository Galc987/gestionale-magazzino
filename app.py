from flask import Flask, render_template, request, redirect

app = Flask(__name__)

clients = {
    "Roberto": ["Catarratto 2L", "Rosato 2L", "Merlot 2L"],
    "Francesco": ["Catarratto 2L", "Chardonnay 2L"]
}

stock = {}

orders = []


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


@app.route("/add_order", methods=["POST"])
def add_order():

    client = request.form.get("client")
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


@app.route("/toggle/<int:o>/<int:i>")
def toggle(o, i):

    orders[o]["items"][i]["done"] = not orders[o]["items"][i]["done"]

    return redirect("/")


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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
