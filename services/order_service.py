from models import db, Order, OrderItem, Product

def get_active_order():
    order = Order.query.first()
    if not order:
        order = Order()
        db.session.add(order)
        db.session.commit()
    return order

def add_product_to_order(product_id, quantity=1):
    order = get_active_order()
    product = Product.query.get(product_id)
    if not product:
        return None
        
    item = OrderItem.query.filter_by(order_id=order.id, product_id=product_id).first()
    if item:
        item.quantity += quantity
    else:
        item = OrderItem(
            order_id=order.id, 
            product_id=product.id, 
            quantity=quantity,
            unit_price=product.unit_price
        )
        db.session.add(item)
    
    db.session.commit()
    return item

def update_item_quantity(item_id, quantity):
    item = OrderItem.query.get(item_id)
    if item:
        if quantity <= 0:
            db.session.delete(item)
        else:
            item.quantity = quantity
        db.session.commit()

def remove_item(item_id):
    item = OrderItem.query.get(item_id)
    if item:
        db.session.delete(item)
        db.session.commit()
