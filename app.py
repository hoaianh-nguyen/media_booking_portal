#!/usr/bin/env python3
"""
ShopeeFood Media Ops — Flask server
Run: python3 app.py
Open: http://127.0.0.1:8080
"""

import os, json
from flask import Flask, jsonify, render_template, abort

app = Flask(__name__)
DATA_DIR = 'data'

def load(filename):
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        abort(503, f"Data not found: {filename}. Run prepare_data.py first.")
    with open(path, encoding='utf-8') as f:
        return json.load(f)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/bookings')
def api_bookings():
    return jsonify(load('bookings.json'))

@app.route('/api/visibility')
def api_visibility():
    return jsonify(load('visibility.json'))

@app.route('/api/caps')
def api_caps():
    return jsonify(load('caps.json'))

@app.route('/api/status')
def api_status():
    files = ['bookings.json', 'visibility.json', 'caps.json']
    status = {}
    for f in files:
        path = os.path.join(DATA_DIR, f)
        if os.path.exists(path):
            status[f] = f"{os.path.getsize(path)//1024} KB"
        else:
            status[f] = "missing — run prepare_data.py"
    return jsonify(status)

if __name__ == '__main__':
    print("\n" + "="*50)
    print("  ShopeeFood Media Ops")
    print("  Open: http://127.0.0.1:8080")
    print("="*50 + "\n")

    missing = [f for f in ['bookings.json','visibility.json','caps.json']
               if not os.path.exists(os.path.join(DATA_DIR, f))]
    if missing:
        print("  ⚠️  Data files missing. Run first:")
        print("      python3 prepare_data.py\n")

    app.run(debug=True, port=8080)
