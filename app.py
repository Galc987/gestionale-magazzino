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

@app.route("/")
def index():
    return render_template(
        "index.html",
        clients=clients,
        stock=stock,
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

    for item in orders[o]["items"]:
        if item["done"]:
            stock[item["product"]] += item["qty"]

    orders.pop(o)

    return redirect("/")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
