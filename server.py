from flask import Flask, request, jsonify
from datetime import datetime
import json
import os

app = Flask(__name__)

DATA_FILE = "inventario_render.json"

def guardar_json(data):
    if not os.path.exists(DATA_FILE):
        with open(DATA_FILE, "w") as f:
            json.dump([], f, indent=4)

    with open(DATA_FILE, "r") as f:
        contenido = json.load(f)

    data["fecha_recepcion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    contenido.append(data)

    with open(DATA_FILE, "w") as f:
        json.dump(contenido, f, indent=4)


@app.route("/inventario", methods=["POST"])
def recibir_inventario():
    data = request.json
    guardar_json(data)
    return jsonify({"status": "ok", "mensaje": "recibido"})


@app.route("/inventario", methods=["GET"])
def enviar_inventario():
    if not os.path.exists(DATA_FILE):
        return jsonify([])
    with open(DATA_FILE, "r") as f:
        return jsonify(json.load(f))


@app.route("/")
def home():
    return jsonify({"mensaje": "Servidor funcionando :)"})
