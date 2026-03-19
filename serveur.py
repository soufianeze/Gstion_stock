"""
CVE — Serveur cloud (Railway)
Stockage en mémoire + export Excel téléchargeable
"""
import os, json, io
from datetime import datetime
from flask import Flask, jsonify, request, send_from_directory, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__, static_folder='.')

# Stockage en mémoire
DATA = {"produits": [], "mouvements": []}

def fill(c): return PatternFill("solid", fgColor=c)
def fnt(bold=False, color="1C2B3A", size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)
def aln(h="center", v="center"):
    return Alignment(horizontal=h, vertical=v, wrap_text=True)
def bdr():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def now_str():
    return datetime.now().strftime("%d/%m/%Y %H:%M")

def generer_excel(produits, mouvements):
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    HDRS = [
        ("Reference", 11), ("Designation", 28), ("Categorie", 15),
        ("Unite", 9), ("Quantite", 12), ("Min", 9), ("Max", 9),
        ("Prix", 11), ("Valeur (DH)", 13), ("Statut", 13),
        ("Date Creation", 20), ("Derniere MAJ", 20), ("Description", 25),
    ]
    for i, (label, width) in enumerate(HDRS, 1):
        c = ws.cell(1, i, label)
        c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c.fill = fill("1A4A6B"); c.alignment = aln(); c.border = bdr()
        ws.column_dimensions[c.column_letter].width = width
    ws.row_dimensions[1].height = 24
    for i, p in enumerate(produits):
        r = i + 2
        statut = "Critique" if float(p.get("qte",0)) <= 0 or float(p.get("qte",0)) < float(p.get("min",0))*.5 \
                 else "Faible" if float(p.get("qte",0)) < float(p.get("min",0)) else "Normal"
        valeur = float(p.get("qte",0)) * float(p.get("prix",0))
        vals = [p.get("ref",""), p.get("nom",""), p.get("cat",""), p.get("unite",""),
                float(p.get("qte",0)), float(p.get("min",0)), float(p.get("max",0)),
                float(p.get("prix",0)), valeur, statut,
                p.get("createdAt", now_str()), p.get("updatedAt", now_str()), p.get("desc","")]
        bg = "E6F5F4" if i % 2 == 0 else "FDFAF5"
        for j, v in enumerate(vals, 1):
            cell = ws.cell(r, j, v)
            cell.fill = fill(bg); cell.border = bdr()
            cell.font = fnt(size=10); cell.alignment = aln("left")
        ws.row_dimensions[r].height = 18
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

@app.route('/')
def index():
    return send_from_directory('.', 'interface.html')

@app.route('/api/produits', methods=['GET'])
def get_produits():
    return jsonify({"produits": DATA["produits"], "mouvements": DATA["mouvements"], "fichier": "Cloud"})

@app.route('/api/produits', methods=['POST'])
def save_produits():
    data = request.json
    DATA["produits"]   = data.get("produits", [])
    DATA["mouvements"] = data.get("mouvements", [])
    return jsonify({"ok": True, "message": f"OK {len(DATA['produits'])} produits", "fichier": "Cloud", "date": now_str()})

@app.route('/api/export', methods=['GET'])
def export_excel():
    buf = generer_excel(DATA["produits"], DATA["mouvements"])
    return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f"CVE_Stock_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

@app.route('/api/statut', methods=['GET'])
def statut():
    return jsonify({"fichier": "Cloud", "existe": True, "taille_ko": 0, "modifie": now_str()})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
