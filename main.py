from flask import Flask, request, jsonify, render_template
import pandas as pd
from openpyxl import load_workbook
import threading
from flask_cors import CORS  # Import flask_cors to handle cross-origin requests

app = Flask(__name__)
CORS(app)  # Enable CORS

# File Excel tempat data akan disimpan
excel_file = 'sensor_data.xlsx'

# Lock untuk penanganan konkruensi
lock = threading.Lock()


# Inisialisasi file Excel dengan kolom jika belum ada
def init_excel_file():
    df = pd.DataFrame(columns=['Timestamp', 'Temperature', 'Humidity', 'Gas'])
    df.to_excel(excel_file, index=False)


# Cek apakah file Excel ada, jika tidak buat baru
try:
    load_workbook(excel_file)
except FileNotFoundError:
    init_excel_file()

@app.route('/')
def index():
    return render_template('index.html')

#ambil data

@app.route('/ambil-data', methods=['GET'])
def get_data():
    # Baca data dari file Excel
    try:
        with lock:
            df = pd.read_excel(excel_file)

        # Konversi DataFrame menjadi list of dicts
        data = df.to_dict(orient='records')

    #     return jsonify(data), 200
    # except Exception as e:
    #     print(f"Error reading Excel file: {e}")
    #     return jsonify({"status": "Failed to read data"}), 500

    #return ke halaman

        return render_template('data.html', data=data)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
    return jsonify({"status": "Failed to read data"}), 500

@app.route('/grafik', methods=['GET'])
def grafik():
    # Baca data dari file Excel
    try:
        with lock:
            df = pd.read_excel(excel_file)
        # Konversi DataFrame menjadi list of dicts
        data = df.to_dict(orient='records')
        # Extract values for the chart
        timestamps = df['Timestamp'].astype(str).tolist()
        temperatures = df['Temperature'].tolist()
        humidity = df['Humidity'].tolist()
        gas = df['Gas'].tolist()
        # Render HTML with table and chart
        return render_template('data_with_graph.html', data=data, timestamps=timestamps, temperatures=temperatures, humidity=humidity, gas=gas)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return jsonify({"status": "Failed to read data"}), 500

@app.route('/sensor-data', methods=['POST'])
def receive_data():
    # Ambil data dari request
    data = request.json
    print("Received data:", data)  # Log the received data for debugging

    temperature = data.get('t')
    humidity = data.get('h')
    gas = data.get('sensorValue')

    # Validasi input
    if temperature is None or humidity is None or gas is None:
        return jsonify({"status": "Invalid data"}), 400

    timestamp = pd.Timestamp.now()
    # Tambahkan data ke file Excel menggunakan lock
    with lock:
        # Baca file Excel
        df = pd.read_excel(excel_file)
        # Tambahkan baris baru
        new_row = {'Timestamp': timestamp, 'Temperature': temperature, 'Humidity': humidity, 'Gas': gas}
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)  # Use concat instead of append
        # Simpan kembali ke file Excel
        df.to_excel(excel_file, index=False)

    return jsonify({"status": "Data Berhasil Disimpan"}), 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
