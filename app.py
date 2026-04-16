from flask import Flask, render_template, request, redirect

app = Flask(__name__)

clients = {
    "Roberto": ["Catarratto 2L", "Rosato 2L", "Merlot 2L"],
    "Francesco": ["Catarratto 2L", "Chardonnay 2L"],
}

# magazzino base (per farlo tornare visibile)
stock = {
    "Catarratto 2L": 100,
    "Rosato 2L": 80,
    "Merlot 2L": 50,
    "Chardonnay 2L": 30
}

orders = []

@app.route("/")
def index():
    return render_template("index.html", clients=clients, stock=stock, orders=orders)

@app.route("/add_order", methods=["POST"])
def add_order():
    client = request.form.get("client")

    products = request.form.getlist("product")
    qtys = request.form.getlist("qty")

    order_items = []

    for p, q in zip(products, qtys):
        if q and int(q) > 0:
            order_items.append({"product": p, "qty": int(q)})

    orders.append({
        "client": client,
        "items": order_items,
        "done": False
    })

    return redirect("/")

@app.route("/toggle_order/<int:index>")
def toggle_order(index):
    orders[index]["done"] = not orders[index]["done"]
    return redirect("/")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
