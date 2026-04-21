import pandas as pd
import openpyxl
import os
import uuid
import math
from flask import current_app
from models import db, Product, Order, OrderItem
import logging

logging.basicConfig(level=logging.INFO)

def import_excel_to_db_with_map(filepath, mapping, currency='₺'):
    """
    Imports excel using the exact indices provided in the mapping dictionary.
    mapping format: {'sku': 0, 'img': 1, 'name': 2, 'desc': 3, 'price': 4, 'color': 5}
    Supports multiple images per row.
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        images_map = {} # (sheetname, row_idx) -> [list of img_paths]
        
        images_dir = os.path.join(current_app.root_path, 'static', 'images')
        os.makedirs(images_dir, exist_ok=True)
        
        img_col = mapping.get('img', -1)

        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            if hasattr(ws, '_images'):
                for image in ws._images:
                    if hasattr(image, 'anchor') and hasattr(image.anchor, '_from'):
                         # Filter by column if specified
                        if img_col != -1 and image.anchor._from.col != img_col:
                            continue
                            
                        row_idx = image.anchor._from.row + 1 
                        img_filename = f"{uuid.uuid4().hex[:10]}.png"
                        img_path = os.path.join(images_dir, img_filename)
                        
                        try:
                            if hasattr(image, '_data'):
                                img_data = image._data()
                            else:
                                img_data = image.ref.read() 
                                
                            if callable(img_data):
                                data_bytes = img_data()
                            else:
                                data_bytes = img_data

                            with open(img_path, 'wb') as f:
                                f.write(data_bytes)
                            
                            key = (sheetname, row_idx)
                            if key not in images_map:
                                images_map[key] = []
                            images_map[key].append(f"images/{img_filename}")
                            
                        except Exception as e:
                            logging.warning(f"Could not extract image at {sheetname} row {row_idx}: {e}")

        xl = pd.ExcelFile(filepath)
        sheet_names = xl.sheet_names
        products_to_add = []
        
        for sheet in sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet, header=None)
            if df.empty:
                continue
            
            for index in range(len(df)):
                row = df.iloc[index]
                
                sku_idx = mapping.get('sku', -1)
                name_idx = mapping.get('name', -1)
                desc_idx = mapping.get('desc', -1)
                price_idx = mapping.get('price', -1)
                color_idx = mapping.get('color', -1)
                
                sku = str(row[sku_idx]) if sku_idx != -1 and not pd.isna(row[sku_idx]) else ""
                name = str(row[name_idx]) if name_idx != -1 and not pd.isna(row[name_idx]) else ""
                desc = str(row[desc_idx]) if desc_idx != -1 and not pd.isna(row[desc_idx]) else ""
                color = str(row[color_idx]) if color_idx != -1 and not pd.isna(row[color_idx]) else ""
                
                price = 0.0
                if price_idx != -1 and not pd.isna(row[price_idx]):
                    try:
                        price_str = str(row[price_idx]).replace(',', '.')
                        # Strip non-numeric except dot
                        price_cleaned = ''.join(c for c in price_str if c.isdigit() or c == '.')
                        price = float(price_cleaned) if price_cleaned else 0.0
                    except:
                        pass
                
                # Header skip logic
                if name.lower() in ['isim', 'ad', 'ürün adı', 'name'] or (sku.lower() in ['kod', 'sku'] and price == 0.0):
                    continue
                    
                if not name and not sku:
                    continue
                    
                excel_row_num = index + 1
                
                # Collect images from this row (or adjacent if offset exists)
                row_images = images_map.get((sheet, excel_row_num), [])
                if not row_images: row_images = images_map.get((sheet, excel_row_num + 1), [])
                if not row_images: row_images = images_map.get((sheet, excel_row_num - 1), [])
                
                image_path_str = "|".join(row_images) if row_images else None
                
                prod = Product(
                    sku=sku if sku != "nan" else "",
                    name=name if name and name != "nan" else ("Ürün" if sku else "İsimsiz"),
                    description=desc if desc != "nan" else "",
                    color=color if color and color != "nan" else "",
                    unit_price=price,
                    image_path=image_path_str,
                    sheet_name=sheet,
                    source_row=excel_row_num 
                )
                products_to_add.append(prod)
                    
        Product.query.delete()
        OrderItem.query.delete()
        Order.query.delete()
        
        # Create new order with selected currency
        new_order = Order(currency=currency)
        db.session.add(new_order)
        
        db.session.add_all(products_to_add)
        db.session.commit()
        return True, f"{len(products_to_add)} ürün başarıyla aktarıldı."
        
    except Exception as e:
        logging.error(f"Import failed: {e}")
        return False, f"Okuma hatası: {str(e)}"
