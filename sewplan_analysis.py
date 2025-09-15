from flask import Flask, jsonify, send_from_directory, render_template
from sqlalchemy import create_engine, text
import logging
import socket
import waitress
import sys
from dotenv import load_dotenv
import os

app = Flask(__name__)

# Load environment variables from .env file
load_dotenv()

username = os.getenv("DB_USERNAME")
password = os.getenv("DB_PASSWORD")
host = os.getenv("DB_HOST")
port = os.getenv("DB_PORT")
service_name = os.getenv("DB_SERVICE_NAME")

nytg_connection_string = f"oracle+oracledb://{username}:{password}@{host}:{port}/{service_name}"
conn = create_engine(nytg_connection_string)

@app.route('/api/plan_data')
def get_plan_data():
    try:
        with conn.connect() as connection:
            result = connection.execute(text("""
                SELECT 
                    d.HEAD_ID,
                    d.SEW_DATE,
                    d.PLAN_PCS,
                    h.FACTORY,
                    h.PROD_LINE,
                    h.CUSTOMER_NAME,
                    h.STYLE_REF,
                    h.DELIVERY_DATE,
                    h.SO_NO_DOC,
                    h.PRODUCT_TYPE,
                    h.SUB_NO,
                    h.COLOR_FC,
                    h.SAM_IE,
                    h.EMBROIDERY,
                    h.HEAT,
                    h.PAD_PRINT,
                    h.PRINT,
                    h.BOND,
                    h.LASER,
                    h.ORDER_PLAN,
                    h.START_SEW
                FROM FR_IMPORT_PLAN_DETAIL d
                LEFT JOIN FR_IMPORT_PLAN_HEAD h ON d.HEAD_ID = h.HEAD_ID
            """))
            columns = result.keys()
            data = [dict(zip(columns, row)) for row in result.fetchall()]
            logging.info(f"Fetched {len(data)} rows from database.")
        return jsonify(data)
    except Exception as e:
        logging.error(f"Error fetching data: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/order_qty_map')
def get_order_qty_map():
    try:
        with conn.connect() as connection:
            result = connection.execute(text("""
                SELECT SO_NO_DOC, SUB_NO, COLOR_FC, ORDER_QTY
                FROM FR_IMPORT_DATA_CENTER
            """))
            data = {}
            for row in result.fetchall():
                so_no_doc = str(row[0]).strip() if row[0] is not None else ''
                sub_no = str(row[1]).strip() if row[1] is not None else ''
                color_fc = str(row[2]).strip() if row[2] is not None else ''
                order_qty = row[3]
                key = f"{so_no_doc}||{sub_no}||{color_fc}"
                data[key] = order_qty
            logging.info(f"Fetched {len(data)} order_qty rows from FR_IMPORT_DATA_CENTER.")
        return jsonify(data)
    except Exception as e:
        logging.error(f"Error fetching order_qty: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/')
def serve_dashboard():
    return render_template('index.html')

@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('.', filename)

if __name__ == '__main__':
    app.run(debug=True,host='0.0.0.0', port=5010)

# if __name__ == '__main__':
#     debug_mode = '--debug' in sys.argv
#     host = '0.0.0.0'
#     port = 5010

#     hostname = socket.gethostname()
#     local_ip = socket.gethostbyname(hostname)
    
#     if debug_mode:
#         print(f"Starting development server at:")
#         print(f"  Local:     http://127.0.0.1:{port}")
#         print(f"  Network:   http://{local_ip}:{port}")
#         app.run(host=host, port=port, debug=True)
#     else:
#         print(f"Starting production server at:")
#         print(f"  Local:     http://127.0.0.1:{port}")
#         print(f"  Network:   http://{local_ip}:{port}")
#         waitress.serve(app, host=host, port=port, threads=4, url_scheme='http')
