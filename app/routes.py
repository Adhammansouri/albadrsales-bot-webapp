import os
from flask import Blueprint, render_template, jsonify
import pandas as pd

main = Blueprint('main', __name__)

@main.route('/')
def index():
    return render_template('index.html')

@main.route('/api/sales-data')
def get_sales_data():
    try:
        # Construct the absolute path to the Excel file
        excel_file_path = os.path.join(os.path.dirname(__file__), '..\..\customer_data.xlsx')
        df = pd.read_excel(excel_file_path)
        return jsonify({
            'status': 'success',
            'data': df.to_dict(orient='records')
        })
    except FileNotFoundError:
         return jsonify({
            'status': 'error',
            'message': 'ملف customer_data.xlsx غير موجود. يرجى التأكد من وجوده في المسار الصحيح.'
        }), 404
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500 