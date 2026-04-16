from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

products = [
    {"name": "Catarratto 2L Roberto", "stock": 0},
    {"name": "Rosato 2L Roberto", "stock": 0},
    {"name": "Catarratto 2L Francesco", "stock": 0},
]

orders = []
production = []

@app.route("/")
def home():
    return render_template("index.html", products=products, orders=orders, production=production)

@app.route("/add_order", methods=["POST"])
def add_order():
    client = request.form["client"]
    product = request.form["product"]
    qty = int(request.form["qty"])

    orders.append({
        "client": client,
        "product": product,
        "qty": qty,
        "status": "da produrre"
    })

    production.append({
        "product": product,
        "qty": qty
    })

    return redirect(url_for("home"))

@app.route("/add_stock", methods=["POST"])
def add_stock():
    product_name = request.form["product"]
    qty = int(request.form["qty"])

    for p in products:
        if p["name"] == product_name:
            p["stock"] += qty

    return redirect(url_for("home"))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
