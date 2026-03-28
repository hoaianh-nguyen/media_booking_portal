#!/usr/bin/env python3
"""
ShopeeFood Media Ops — Flask server
Run: python3 app.py
Open: http://127.0.0.1:8080
"""

import os, json, gzip
from urllib.parse import quote
from flask import Flask, jsonify, render_template, abort, Response, send_from_directory

app = Flask(__name__)
DATA_DIR = 'data'
PNG_DIR  = 'PNG'   # folder containing date subfolders with creative images

def load_raw(filename):
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        abort(503, f"Data not found: {filename}. Run prepare_data.py first.")
    with open(path, 'rb') as f:
        return f.read()

def load(filename):
    return json.loads(load_raw(filename))

def gzip_json(data_bytes):
    compressed = gzip.compress(data_bytes, compresslevel=6)
    resp = Response(compressed, mimetype='application/json')
    resp.headers['Content-Encoding'] = 'gzip'
    resp.headers['Cache-Control'] = 'no-store'
    resp.headers['Vary'] = 'Accept-Encoding'
    return resp

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/bookings')
def api_bookings():
    return gzip_json(load_raw('bookings.json'))

@app.route('/api/visibility')
def api_visibility():
    return gzip_json(load_raw('visibility.json'))

@app.route('/api/caps')
def api_caps():
    return gzip_json(load_raw('caps.json'))

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

@app.route('/png/<path:filename>')
def serve_png(filename):
    """Serve creative images from the PNG folder."""
    return send_from_directory(PNG_DIR, filename)

@app.route('/api/png-index')
def api_png_index():
    """Return nested dict: {date: {bannerType: ['/png/date/Banner/file.png', ...]}}
    Structure: PNG/1.4/Hometop/file.png  or  PNG/1.4/file.png
    """
    index = {}
    if not os.path.isdir(PNG_DIR):
        return jsonify(index)
    EXT = ('.png','.jpg','.jpeg','.webp','.gif')
    for date_folder in sorted(os.listdir(PNG_DIR)):
        date_path = os.path.join(PNG_DIR, date_folder)
        if not os.path.isdir(date_path):
            continue
        index[date_folder] = {}
        for item in sorted(os.listdir(date_path)):
            item_path = os.path.join(date_path, item)
            if os.path.isdir(item_path):
                # PNG/1.4/Hometop/file1.png, file2.png …
                banner = item  # folder name IS the banner type
                files = [f for f in sorted(os.listdir(item_path))
                         if f.lower().endswith(EXT)]
                if files:
                    # Store list of URLs
                    index[date_folder][banner] = [
                        f'/png/{quote(date_folder)}/{quote(item)}/{quote(f)}' for f in files
                    ]
            elif item.lower().endswith(EXT):
                # PNG/1.4/Hometop.png  — flat inside date folder
                banner = item.rsplit('.',1)[0]
                index[date_folder][banner] = [f'/png/{quote(date_folder)}/{quote(item)}']
    return jsonify(index)

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

    if os.path.isdir(PNG_DIR):
        dates = [d for d in os.listdir(PNG_DIR) if os.path.isdir(os.path.join(PNG_DIR,d))]
        print(f"  📁 PNG folder found: {len(dates)} date folder(s): {sorted(dates)}")
    else:
        print(f"  ℹ️  No PNG/ folder found yet. Add creative images to PNG/<date>/")

    app.run(debug=True, port=8080)
