import os
import pandas as pd
from flask import render_template, request, redirect, url_for, flash, current_app, send_file
from werkzeug.utils import secure_filename
from models import Product, db
from services.excel_importer import import_excel_to_db_with_map
from services.order_service import get_active_order, add_product_to_order, update_item_quantity, remove_item
from services.export_service import export_order_to_excel

def init_routes(app):
    
    @app.route('/')
    def index():
        search = request.args.get('search', '')
        category = request.args.get('category', '')
        
        query = Product.query
        if search:
            query = query.filter(Product.name.ilike(f'%{search}%') | Product.sku.ilike(f'%{search}%'))
        if category:
            query = query.filter(Product.sheet_name == category)
            
        products = query.order_by(Product.id.asc()).all()
        
        categories_tuples = db.session.query(Product.sheet_name).distinct().all()
        categories = [c[0] for c in categories_tuples if c[0]]
        
        order = get_active_order()
        return render_template('index.html', products=products, search=search, categories=categories, current_category=category, order=order)

    @app.route('/upload', methods=['POST'])
    def upload_file():
        if 'file' not in request.files:
            flash('Dosya seçilmedi', 'error')
            return redirect(url_for('index'))
            
        file = request.files['file']
        if file.filename == '':
            flash('Dosya seçilmedi', 'error')
            return redirect(url_for('index'))
            
        if file and file.filename.endswith(('.xls', '.xlsx')):
            filename = secure_filename(file.filename)
            filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            return redirect(url_for('map_columns', filename=filename))
                
        return redirect(url_for('index'))

    @app.route('/map_columns', methods=['GET'])
    def map_columns():
        filename = request.args.get('filename')
        if not filename:
            return redirect(url_for('index'))
            
        filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
        
        try:
            xl = pd.ExcelFile(filepath)
            df = pd.read_excel(filepath, sheet_name=xl.sheet_names[0], header=None, nrows=10)
            
            # Show first 3 rows as sample string
            sample_rows = []
            for i in range(min(3, len(df))):
                row_vals = [f"Sütun {chr(65+j)}: {str(v)[:20]}" if not pd.isna(v) else f"Sütun {chr(65+j)}: (Boş)" for j, v in enumerate(df.iloc[i])]
                sample_rows.append(" | ".join(row_vals))
                
            cols_count = len(df.columns)
        except Exception as e:
            flash(f'Dosya okunurken hata: {e}', 'error')
            return redirect(url_for('index'))
            
        return render_template('map_columns.html', filename=filename, cols_count=cols_count, sample_rows=sample_rows, chr=chr)

    @app.route('/process_import', methods=['POST'])
    def process_import():
        filename = request.form.get('filename')
        mapping = {
            'sku': request.form.get('sku', type=int),
            'name': request.form.get('name', type=int),
            'desc': request.form.get('desc', type=int),
            'price': request.form.get('price', type=int),
            'color': request.form.get('color', type=int, default=-1),
            'img': request.form.get('img', type=int)
        }
        
        filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
        success, msg = import_excel_to_db_with_map(filepath, mapping)
        
        if success:
            flash(msg, 'success')
        else:
            flash(msg, 'error')
            
        return redirect(url_for('index'))
        

    @app.route('/search', methods=['GET'])
    def search_products():
        search = request.args.get('search', '')
        category = request.args.get('category', '')
        
        query = Product.query
        if search:
            query = query.filter(Product.name.ilike(f'%{search}%') | Product.sku.ilike(f'%{search}%'))
        if category:
            query = query.filter(Product.sheet_name == category)
            
        products = query.order_by(Product.id.asc()).all()
        return render_template('partials/product_list.html', products=products)

    @app.route('/cart/add', methods=['POST'])
    def add_to_cart():
        product_id = request.form.get('product_id', type=int)
        quantity = request.form.get('quantity', type=int, default=1)
        if product_id:
            add_product_to_order(product_id, quantity)
        order = get_active_order()
        return render_template('partials/order_list.html', order=order)
        
    @app.route('/cart/update/<int:item_id>', methods=['POST'])
    def update_cart(item_id):
        quantity = request.form.get('quantity', type=int)
        update_item_quantity(item_id, quantity)
        order = get_active_order()
        return render_template('partials/order_list.html', order=order)
        
    @app.route('/cart/remove/<int:item_id>', methods=['POST'])
    def remove_from_cart(item_id):
        remove_item(item_id)
        order = get_active_order()
        return render_template('partials/order_list.html', order=order)

    @app.route('/export', methods=['POST'])
    def export_excel():
        filepath = export_order_to_excel(current_app.config['UPLOAD_FOLDER'])
        if filepath:
            filename = os.path.basename(filepath)
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            flash('İndirilecek sipariş öğesi bulunamadı', 'error')
            return redirect(url_for('index'))
