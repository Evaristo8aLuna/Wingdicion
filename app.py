import os
import PIL.Image
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, timedelta
import pytz
import google.generativeai as genai
import json
from werkzeug.utils import secure_filename
import io
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

app = Flask(__name__)
app.config['SECRET_KEY'] = 'boneless_super_secret_key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///boneless.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# Configuración de carga de imágenes
UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# Configuración de Gemini
GEMINI_KEY = "AIzaSyDCxlzbXoedrxA-Izgf2r3xvGG-u99rSas"
app.config['GOOGLE_API_KEY'] = GEMINI_KEY

db = SQLAlchemy(app)

from sqlalchemy import event
from sqlalchemy.engine import Engine

@event.listens_for(Engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

# --- AI Setup ---
if GEMINI_KEY and GEMINI_KEY != "TU_API_KEY_AQUI":
    genai.configure(api_key=GEMINI_KEY)
    
    # Intentar detectar el mejor modelo disponible automáticamente
    try:
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        if 'models/gemini-1.5-flash' in available_models:
            model_name = 'gemini-1.5-flash'
        elif 'models/gemini-1.5-flash-latest' in available_models:
            model_name = 'gemini-1.5-flash-latest'
        elif 'models/gemini-pro' in available_models:
            model_name = 'gemini-pro'
        elif available_models:
            model_name = available_models[0].replace('models/', '')
        else:
            model_name = 'gemini-1.5-flash' # Fallback
        
        model = genai.GenerativeModel(model_name)
        print(f"IA activa con el modelo: {model_name}")
    except Exception as e:
        print(f"Error al listar modelos: {e}")
        model = genai.GenerativeModel('gemini-1.5-flash')
else:
    print("AI Warning: No valid API Key found.")
    model = None

CACHE_FILE = 'ai_cache.json'

def get_cached_ai():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                cache = json.load(f)
                cache_time = datetime.fromisoformat(cache['timestamp'])
                # Válido por 12 horas para ahorrar cuota
                if datetime.now() - cache_time < timedelta(hours=12):
                    return cache['data']
        except Exception:
            return None
    return None

def save_ai_cache(data):
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump({'timestamp': datetime.now().isoformat(), 'data': data}, f)
    except Exception:
        pass

# --- Timezone Setup ---
JUAREZ_TZ = pytz.timezone('America/Ciudad_Juarez')

def get_local_time():
    # Get time with timezone info and then convert to naive for SQLite compatibility
    return datetime.now(JUAREZ_TZ).replace(tzinfo=None)

def is_store_open():
    now = get_local_time()
    weekday = now.weekday() # 0: Monday, 3: Thursday, 4: Friday, 5: Saturday, 6: Sunday
    hour = now.hour
    
    # Thursday (3) and Friday (4): 5 PM (17) to 10 PM (22)
    if weekday in [3, 4]:
        return 17 <= hour < 22
    # Saturday (5) and Sunday (6): 4 PM (16) to 10 PM (22)
    elif weekday in [5, 6]:
        return 16 <= hour < 22
    return False

# --- Helper for Files ---
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- Models ---
class MenuItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    price = db.Column(db.Float, nullable=False)
    category = db.Column(db.String(50), nullable=False)
    image_url = db.Column(db.String(255), nullable=True)

class InventoryItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False) # e.g. Pechuga, Alitas, Salsa BBQ
    category = db.Column(db.String(50), nullable=False) # e.g. Materia Prima, Salsas, Empaques
    quantity = db.Column(db.Float, default=0.0)
    unit = db.Column(db.String(20), nullable=False) # e.g. kg, litros, piezas
    provider = db.Column(db.String(100), nullable=True)

class BOMItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    menu_item_id = db.Column(db.Integer, db.ForeignKey('menu_item.id'), nullable=False)
    inventory_item_id = db.Column(db.Integer, db.ForeignKey('inventory_item.id'), nullable=False)
    quantity = db.Column(db.Float, nullable=False)

    # Relationships
    menu_item = db.relationship('MenuItem', backref=db.backref('bom_items', lazy=True, cascade="all, delete-orphan"))
    inventory_item = db.relationship('InventoryItem', backref=db.backref('bom_usages', lazy=True))

RECIPES = {
    "Boneless/Alitas 300 GR (7-8 pz)": [
        {"name": "Pechuga de Pollo", "quantity": 0.3},
        {"name": "Cajas para Boneless", "quantity": 1}
    ],
    "Boneless/Alitas 480 GR (11-13 pz)": [
        {"name": "Pechuga de Pollo", "quantity": 0.48},
        {"name": "Cajas para Boneless", "quantity": 1}
    ],
    "Boneless/Alitas 1 KG (24-26 pz)": [
        {"name": "Pechuga de Pollo", "quantity": 1.0},
        {"name": "Cajas para Boneless", "quantity": 1}
    ],
    "Hamburguesa Clásica Sencilla": [
        {"name": "Carne para Hamburguesa Sirloin", "quantity": 1},
        {"name": "Pan para Hamburguesa", "quantity": 1},
        {"name": "Cajas para Hamburguesa", "quantity": 1}
    ],
    "Hamburguesa Clásica Doble": [
        {"name": "Carne para Hamburguesa Sirloin", "quantity": 2},
        {"name": "Pan para Hamburguesa", "quantity": 1},
        {"name": "Cajas para Hamburguesa", "quantity": 1}
    ],
    "Hamburguesa Texas Sencilla": [
        {"name": "Carne para Hamburguesa Arrachera", "quantity": 1},
        {"name": "Pan para Hamburguesa", "quantity": 1},
        {"name": "Cajas para Hamburguesa", "quantity": 1}
    ],
    "Hamburguesa Texas Doble": [
        {"name": "Carne para Hamburguesa Arrachera", "quantity": 2},
        {"name": "Pan para Hamburguesa", "quantity": 1},
        {"name": "Cajas para Hamburguesa", "quantity": 1}
    ],
    "Hamburguesa Hawaiana Sencilla": [
        {"name": "Carne para Hamburguesa Sirloin", "quantity": 1},
        {"name": "Pan para Hamburguesa", "quantity": 1},
        {"name": "Piña", "quantity": 1},
        {"name": "Jamón", "quantity": 1},
        {"name": "Cajas para Hamburguesa", "quantity": 1}
    ],
    "Hamburguesa Hawaiana Doble": [
        {"name": "Carne para Hamburguesa Sirloin", "quantity": 2},
        {"name": "Pan para Hamburguesa", "quantity": 1},
        {"name": "Piña", "quantity": 1},
        {"name": "Jamón", "quantity": 1},
        {"name": "Cajas para Hamburguesa", "quantity": 1}
    ],
    "Hamburguesa Pollo Crispy": [
        {"name": "Pechuga Empanizada (Pollo Crispy)", "quantity": 1},
        {"name": "Pan para Hamburguesa", "quantity": 1},
        {"name": "Cajas para Hamburguesa", "quantity": 1}
    ],
    "Extra Aderezo Ranch": [
        {"name": "Aderezo Ranch", "quantity": 0.05},
        {"name": "Vasitos 2oz (Aderezos)", "quantity": 1}
    ],
    "Extra Queso Amarillo/Salsa": [
        {"name": "Queso Amarillo", "quantity": 1}
    ],
    "Extra Papas Fritas": [
        {"name": "Papas a la francesa", "quantity": 0.15}
    ],
    "Coca-Cola 600 ML": [
        {"name": "Coca-Cola 600 ML", "quantity": 1}
    ]
}

class Purchase(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    item_id = db.Column(db.Integer, db.ForeignKey('inventory_item.id'), nullable=False)
    quantity_bought = db.Column(db.Float, nullable=False)
    cost = db.Column(db.Float, nullable=False)
    date = db.Column(db.DateTime, default=get_local_time)
    provider = db.Column(db.String(100), nullable=True)
    
    item = db.relationship('InventoryItem', backref=db.backref('purchases', lazy=True))

class Sale(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_name = db.Column(db.String(100), nullable=False) # e.g. Orden Boneless Mediana BBQ
    total_price = db.Column(db.Float, nullable=False)
    date = db.Column(db.DateTime, default=get_local_time)
    customer_name = db.Column(db.String(100), nullable=True)
    phone = db.Column(db.String(20), nullable=True)
    delivery_address = db.Column(db.Text, nullable=True)
    notes = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(20), default='completada') # completada, pendiente, cancelada


class Consumption(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    sale_id = db.Column(db.Integer, db.ForeignKey('sale.id', ondelete='CASCADE'), nullable=False)
    inventory_item_id = db.Column(db.Integer, db.ForeignKey('inventory_item.id'), nullable=False)
    quantity_consumed = db.Column(db.Float, nullable=False)
    date = db.Column(db.DateTime, default=get_local_time)
    
    # Relationships
    sale = db.relationship('Sale', backref=db.backref('consumptions', lazy=True, cascade="all, delete-orphan", passive_deletes=True))
    inventory_item = db.relationship('InventoryItem', backref=db.backref('consumptions', lazy=True))


class InfoRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    inventory_item_id = db.Column(db.Integer, db.ForeignKey('inventory_item.id', ondelete='CASCADE'), nullable=False)
    provider = db.Column(db.String(100), nullable=False)
    cost_per_unit = db.Column(db.Float, nullable=False)
    is_default = db.Column(db.Boolean, default=True)
    last_updated = db.Column(db.DateTime, default=get_local_time, onupdate=get_local_time)
    
    # Relationships
    inventory_item = db.relationship('InventoryItem', backref=db.backref('info_records', lazy=True, cascade="all, delete-orphan", passive_deletes=True))


# --- API Routes for Live Tracking ---
@app.route('/api/order_status/<int:id>')
def get_order_status(id):
    order = Sale.query.get_or_404(id)
    return {'status': order.status}

# --- Helper Functions ---
def deduct_inventory(sale):
    product_name = sale.product_name
    items_sold = [item.strip() for item in product_name.split('\n')]
    # Handle both newline and comma separated for compatibility
    cleaned_items = []
    for line in items_sold:
        cleaned_items.extend([i.strip() for i in line.split(',')])

    for item_name in cleaned_items:
        # Check if item name contains quantity like "2x "
        qty = 1
        name_to_check = item_name
        if 'x ' in item_name:
            try:
                parts = item_name.split('x ', 1)
                qty = int(parts[0])
                name_to_check = parts[1]
            except ValueError:
                pass
        
        # 1. Descontar receta base
        menu_item = MenuItem.query.filter(MenuItem.name.contains(name_to_check.split(' (')[0])).first()
        if menu_item and menu_item.bom_items:
            for bom_item in menu_item.bom_items:
                inv_item = bom_item.inventory_item
                if inv_item:
                    consumption_qty = bom_item.quantity * qty
                    inv_item.quantity -= consumption_qty
                    
                    # Registrar en el historial de consumos
                    consumption = Consumption(
                        sale_id=sale.id,
                        inventory_item_id=inv_item.id,
                        quantity_consumed=consumption_qty,
                        date=sale.date
                    )
                    db.session.add(consumption)
        
        # 2. Detección automática de salsas por palabras clave en la orden
        import unicodedata
        def strip_accents(text):
            return "".join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
        
        normalized_line = strip_accents(item_name.lower())
        
        # Buscar todas las salsas y aderezos en el inventario
        salsas = InventoryItem.query.filter(InventoryItem.category.in_(['Salsas', 'Salsas y Aderezos'])).all()
        for salsa in salsas:
            # Limpiar descriptores comunes para extraer la palabra clave (ej: "BBQ", "Bufalo")
            name_clean = salsa.name.replace('Salsa', '').replace('Polvo', '').replace('Aderezo', '').strip()
            normalized_salsa_keyword = strip_accents(name_clean.lower())
            
            if normalized_salsa_keyword and normalized_salsa_keyword in normalized_line:
                # Asignar porción por defecto según la unidad
                if salsa.unit == 'litros':
                    port_qty = 0.05 # 50 ml
                elif salsa.unit == 'kg':
                    port_qty = 0.01 # 10 g
                else:
                    port_qty = 1.0  # 1 pieza/unidad
                
                consumption_qty = port_qty * qty
                salsa.quantity -= consumption_qty
                
                # Registrar en el historial de consumos
                consumption = Consumption(
                    sale_id=sale.id,
                    inventory_item_id=salsa.id,
                    quantity_consumed=consumption_qty,
                    date=sale.date
                )
                db.session.add(consumption)
                
    db.session.commit()


# Ensure tables are created
with app.app_context():
    db.create_all()
    
    # Migración automática de base de datos para la columna 'provider'
    try:
        db.session.execute(db.text("ALTER TABLE inventory_item ADD COLUMN provider VARCHAR(100)"))
        db.session.commit()
        print("Migración: Columna 'provider' agregada exitosamente a la tabla 'inventory_item'.")
    except Exception:
        db.session.rollback()
        
    try:
        db.session.execute(db.text("ALTER TABLE purchase ADD COLUMN provider VARCHAR(100)"))
        db.session.commit()
        print("Migración: Columna 'provider' agregada exitosamente a la tabla 'purchase'.")
    except Exception:
        db.session.rollback()
    
    # Seed MenuItems if empty
    if not MenuItem.query.first():
        initial_menu = [
            MenuItem(name="Boneless/Alitas 300 GR (7-8 pz)", price=145, category="Boneless / Alitas"),
            MenuItem(name="Boneless/Alitas 480 GR (11-13 pz)", price=230, category="Boneless / Alitas"),
            MenuItem(name="Boneless/Alitas 1 KG (24-26 pz)", price=479, category="Boneless / Alitas"),
            MenuItem(name="Hamburguesa Clásica Sencilla", price=56, category="Hamburguesas Clásica"),
            MenuItem(name="Hamburguesa Clásica Doble", price=85, category="Hamburguesas Clásica"),
            MenuItem(name="Hamburguesa Texas Sencilla", price=95, category="Hamburguesas Estilo Texas"),
            MenuItem(name="Hamburguesa Texas Doble", price=125, category="Hamburguesas Estilo Texas"),
            MenuItem(name="Hamburguesa Hawaiana Sencilla", price=95, category="Hamburguesas Especiales"),
            MenuItem(name="Hamburguesa Hawaiana Doble", price=125, category="Hamburguesas Especiales"),
            MenuItem(name="Hamburguesa Pollo Crispy", price=84, category="Hamburguesas Especiales"),
            MenuItem(name="Extra Aderezo Ranch", price=15, category="Extras"),
            MenuItem(name="Extra Queso Amarillo/Salsa", price=20, category="Extras"),
            MenuItem(name="Extra Papas Fritas", price=35, category="Extras"),
            MenuItem(name="Coca-Cola 600 ML", price=25, category="Extras")
        ]
        db.session.bulk_save_objects(initial_menu)
        db.session.commit()

    # Seed initial inventory items if empty
    if not InventoryItem.query.first():
        initial_items = [
            InventoryItem(name='Pechuga de Pollo', category='Materia Prima', unit='kg'),
            InventoryItem(name='Alitas', category='Materia Prima', unit='kg'),
            InventoryItem(name='Salsa Búfalo', category='Salsas', unit='litros'),
            InventoryItem(name='Salsa BBQ', category='Salsas', unit='litros'),
            InventoryItem(name='Salsa Mango Habanero', category='Salsas', unit='litros'),
            InventoryItem(name='Piña', category='Materia Prima', unit='rebanadas'),
            InventoryItem(name='Jamón', category='Materia Prima', unit='rebanadas'),
            InventoryItem(name='Cajas para Boneless', category='Empaques', unit='piezas'),
            InventoryItem(name='Carne para Hamburguesa Sirloin', category='Materia Prima', unit='piezas'),
            InventoryItem(name='Pan para Hamburguesa', category='Materia Prima', unit='piezas'),
            InventoryItem(name='Cajas para Hamburguesa', category='Empaques', unit='piezas'),
            InventoryItem(name='Carne para Hamburguesa Arrachera', category='Materia Prima', unit='piezas'),
            InventoryItem(name='Pechuga Empanizada (Pollo Crispy)', category='Materia Prima', unit='piezas'),
            InventoryItem(name='Aderezo Ranch', category='Salsas', unit='litros'),
            InventoryItem(name='Vasitos 2oz (Aderezos)', category='Empaques', unit='piezas'),
            InventoryItem(name='Queso Amarillo', category='Salsas', unit='piezas'),
            InventoryItem(name='Papas a la francesa', category='Materia Prima', unit='kg'),
            InventoryItem(name='Coca-Cola 600 ML', category='Bebidas', unit='piezas')
        ]
        db.session.bulk_save_objects(initial_items)
        db.session.commit()

        # Seed BOM items based on RECIPES
        for recipe_name, ingredients in RECIPES.items():
            menu_item = MenuItem.query.filter_by(name=recipe_name).first()
            if menu_item:
                for ing in ingredients:
                    inv_item = InventoryItem.query.filter_by(name=ing['name']).first()
                    if inv_item:
                        bom_item = BOMItem(menu_item_id=menu_item.id, inventory_item_id=inv_item.id, quantity=ing['quantity'])
                        db.session.add(bom_item)
        db.session.commit()

# --- Routes ---
@app.route('/')
def index():
    recent_sales = Sale.query.order_by(Sale.date.desc()).limit(5).all()
    total_sales = db.session.query(db.func.sum(Sale.total_price)).scalar() or 0.0
    items = InventoryItem.query.all()
    
    # --- Datos para la gráfica semanal (últimas 8 semanas) ---
    today = get_local_time().date()
    # Ir 8 semanas atrás desde el lunes de la semana actual
    start_of_week = today - timedelta(days=today.weekday())
    start_date = start_of_week - timedelta(weeks=7)
    
    sales_data = Sale.query.filter(Sale.date >= datetime.combine(start_date, datetime.min.time())).all()
    
    # Inicializar las 8 semanas con 0
    weekly_data = {}
    for i in range(8):
        week_start = start_date + timedelta(weeks=i)
        label = f"Sem. {week_start.strftime('%d/%m')}"
        weekly_data[label] = 0.0
        
    for sale in sales_data:
        # Encontrar a qué semana pertenece la venta
        sale_date = sale.date.date()
        days_from_start = (sale_date - start_date).days
        week_index = days_from_start // 7
        if 0 <= week_index < 8:
            week_start = start_date + timedelta(weeks=week_index)
            label = f"Sem. {week_start.strftime('%d/%m')}"
            weekly_data[label] += sale.total_price
            
    chart_labels = list(weekly_data.keys())
    chart_values = list(weekly_data.values())

    # --- Datos para la gráfica mensual (últimos 6 meses) ---
    monthly_data = {}
    # Obtener fecha de hace 6 meses
    six_months_ago = today - timedelta(days=180)
    monthly_sales_query = Sale.query.filter(Sale.date >= datetime.combine(six_months_ago, datetime.min.time())).all()
    
    # Inicializar meses
    for i in range(6):
        month_date = today.replace(day=1) - timedelta(days=i*30)
        month_date = month_date.replace(day=1)
        label = month_date.strftime('%b %Y')
        monthly_data[label] = 0.0
    
    for sale in monthly_sales_query:
        label = sale.date.strftime('%b %Y')
        if label in monthly_data:
            monthly_data[label] += sale.total_price
            
    # Invertir para que sea cronológico
    monthly_chart_labels = list(reversed(list(monthly_data.keys())))
    monthly_chart_values = list(reversed(list(monthly_data.values())))
    
    return render_template('index.html', 
                           recent_sales=recent_sales, 
                           total_sales=total_sales, 
                           items=items,
                           chart_labels=chart_labels,
                           chart_values=chart_values,
                           monthly_chart_labels=monthly_chart_labels,
                           monthly_chart_values=monthly_chart_values)

@app.route('/menu', methods=['GET', 'POST'])
def menu():
    if request.method == 'POST':
        name = request.form['name']
        price = float(request.form['price'])
        category = request.form['category']

        image_url = None
        if 'image' in request.files:
            file = request.files['image']
            if file and allowed_file(file.filename):
                filename = secure_filename(f"{get_local_time().timestamp()}_{file.filename}")
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                image_url = f"uploads/{filename}"

        new_item = MenuItem(name=name, price=price, category=category, image_url=image_url)
        db.session.add(new_item)
        db.session.commit()
        flash('Producto agregado al menú.', 'success')
        return redirect(url_for('menu'))

    menu_items = MenuItem.query.order_by(MenuItem.category, MenuItem.name).all()
    return render_template('menu.html', menu_items=menu_items)

@app.route('/edit_menu_item/<int:id>', methods=['GET', 'POST'])
def edit_menu_item(id):
    item = MenuItem.query.get_or_404(id)
    if request.method == 'POST':
        item.name = request.form['name']
        item.price = float(request.form['price'])
        item.category = request.form['category']

        if 'image' in request.files:
            file = request.files['image']
            if file and allowed_file(file.filename):
                # Opcional: eliminar imagen vieja si existe
                filename = secure_filename(f"{get_local_time().timestamp()}_{file.filename}")
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                item.image_url = f"uploads/{filename}"

        db.session.commit()
        flash('Producto del menú actualizado.', 'success')
        return redirect(url_for('menu'))
    return render_template('edit_menu_item.html', item=item)
@app.route('/delete_menu_item/<int:id>', methods=['POST'])
def delete_menu_item(id):
    item = MenuItem.query.get_or_404(id)
    db.session.delete(item)
    db.session.commit()
    flash('Producto eliminado del menú.', 'danger')
    return redirect(url_for('menu'))

@app.route('/inventory', methods=['GET', 'POST'])
def inventory():
    if request.method == 'POST':
        name = request.form['name']
        category = request.form['category']
        unit = request.form['unit']
        provider = request.form.get('provider')
        new_item = InventoryItem(name=name, category=category, unit=unit, provider=provider)
        db.session.add(new_item)
        db.session.commit()
        flash('Nuevo artículo agregado al inventario.', 'success')
        return redirect(url_for('inventory'))
    
    items = InventoryItem.query.all()
    return render_template('inventory.html', items=items)

@app.route('/purchases_history')
def purchases_history():
    purchases_list = Purchase.query.order_by(Purchase.date.desc()).all()
    return render_template('purchases_history.html', purchases=purchases_list)

@app.route('/purchases', methods=['GET', 'POST'])
def purchases():
    if request.method == 'POST':
        item_id = request.form['item_id']
        quantity = float(request.form['quantity'])
        cost = float(request.form['cost'])
        provider = request.form.get('provider')
        
        # Add purchase record
        purchase = Purchase(item_id=item_id, quantity_bought=quantity, cost=cost, provider=provider)
        db.session.add(purchase)
        
        # Update inventory
        item = InventoryItem.query.get(item_id)
        item.quantity += quantity
        
        db.session.commit()
        flash('Compra registrada con éxito y el inventario ha sido actualizado.', 'success')
        return redirect(url_for('purchases'))
        
    items = InventoryItem.query.all()
    purchases_list = Purchase.query.order_by(Purchase.date.desc()).all()
    return render_template('purchases.html', items=items, purchases=purchases_list)

@app.route('/sales', methods=['GET', 'POST'])
def sales():
    if request.method == 'POST':
        product_name = request.form['product_name']
        total_price = float(request.form['total_price'])
        
        # Handle custom date
        sale_date = get_local_time()
        sale_date_str = request.form.get('sale_date')
        if sale_date_str:
            try:
                # datetime-local format: YYYY-MM-DDTHH:MM
                sale_date = datetime.strptime(sale_date_str, '%Y-%m-%dT%H:%M')
            except ValueError:
                pass
        
        # Add sale record
        sale = Sale(
            product_name=product_name, 
            total_price=total_price, 
            status='completada',
            date=sale_date
        )
        db.session.add(sale)
        db.session.flush()
        
        # Deduct ingredients from inventory
        deduct_inventory(sale)
        
        db.session.commit()
        flash('Venta registrada con éxito y el inventario ha sido actualizado.', 'success')
        return redirect(url_for('sales'))
        
    # Fetch menu items from DB and group by category
    all_items = MenuItem.query.order_by(MenuItem.category, MenuItem.name).all()
    menu_data = {}
    for item in all_items:
        if item.category not in menu_data:
            menu_data[item.category] = []
        menu_data[item.category].append({
            'name': item.name, 
            'price': item.price,
            'image_url': item.image_url
        })
    
    # Format for the template (list of dicts with 'category' and 'items')
    formatted_menu = []
    for category, items in menu_data.items():
        formatted_menu.append({'category': category, 'items': items})
        
    return render_template('sales.html', menu=formatted_menu)

@app.route('/menu_digital')
def public_menu():
    # Fetch menu items from DB and group by category
    all_items = MenuItem.query.order_by(MenuItem.category, MenuItem.name).all()
    menu_data = {}
    for item in all_items:
        if item.category not in menu_data:
            menu_data[item.category] = []
        menu_data[item.category].append({
            'name': item.name, 
            'price': item.price,
            'image_url': item.image_url
        })
    
    # Format for the template (list of dicts with 'category' and 'items')
    formatted_menu = []
    for category, items in menu_data.items():
        formatted_menu.append({'category': category, 'items': items})
        
    return render_template('public_menu.html', menu=formatted_menu)

@app.route('/qr_menu')
def qr_menu():
    import urllib.parse
    # URL de tu menú en PythonAnywhere
    base_url = "https://wingdicion.pythonanywhere.com/menu_digital"
    # Usaremos una API gratuita para generar el QR sin instalar librerías extras
    qr_api_url = f"https://api.qrserver.com/v1/create-qr-code/?size=300x300&data={urllib.parse.quote(base_url)}"
    
    return render_template('qr_display.html', qr_url=qr_api_url, site_url=base_url)

@app.route('/order', methods=['GET', 'POST'])
def public_order():
    if request.method == 'POST':
        product_name = request.form['product_name']
        total_price = float(request.form['total_price'])
        customer_name = request.form['customer_name']
        delivery_address = request.form.get('delivery_address')
        phone = request.form.get('phone')
        notes = request.form.get('notes')
        
        # Add pending order
        new_order = Sale(
            product_name=product_name, 
            total_price=total_price, 
            customer_name=customer_name,
            delivery_address=delivery_address,
            phone=phone,
            notes=notes,
            status='pendiente'
        )
        db.session.add(new_order)
        db.session.commit()
        return render_template('order_success.html', order=new_order)
        
    all_items = MenuItem.query.order_by(MenuItem.category, MenuItem.name).all()
    menu_data = {}
    for item in all_items:
        if item.category not in menu_data:
            menu_data[item.category] = []
        menu_data[item.category].append({
            'name': item.name, 
            'price': item.price,
            'image_url': item.image_url
        })
    
    formatted_menu = []
    for category, items in menu_data.items():
        formatted_menu.append({'category': category, 'items': items})
        
    return render_template('public_order.html', menu=formatted_menu)

@app.route('/admin/pending')
def pending_orders():
    orders = Sale.query.filter(Sale.status.in_(['pendiente', 'recibida'])).order_by(Sale.date.asc()).all()
    return render_template('pending_orders.html', orders=orders)

@app.route('/admin/confirm_order/<int:id>', methods=['POST'])
def confirm_order(id):
    order = Sale.query.get_or_404(id)
    if order.status == 'pendiente':
        order.status = 'recibida'
        db.session.commit()
        flash(f'Orden #{id} marcada como recibida.', 'info')
    return redirect(url_for('pending_orders'))

@app.route('/generate_receipt/<int:order_id>')
def generate_receipt(order_id):
    order = Sale.query.get_or_404(order_id)
    
    doc = Document()
    
    # Folio number
    folio = f"{order.id:06d}"
    
    # Título del Recibo
    titulo = doc.add_heading(f'RECIBO DE VENTA - FOLIO {folio}', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Información de la Empresa
    empresa_info = doc.add_paragraph()
    empresa_info.add_run('WINGDICIÓN\n').bold = True
    empresa_info.add_run('Ciudad Juárez, Chihuahua, México')
    empresa_info.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading('Datos del Cliente', level=1)
    cliente = doc.add_paragraph(
        f"Nombre: {order.customer_name or 'N/A'}\n"
        f"Teléfono: {order.phone or 'N/A'}\n"
        f"Dirección: {order.delivery_address or 'N/A'}"
    )

    doc.add_heading('Detalle de la Operación', level=1)
    
    # Crear tabla de detalles
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Concepto'
    hdr_cells[1].text = 'Total'

    row_cells = table.add_row().cells
    row_cells[0].text = order.product_name
    row_cells[1].text = f'${order.total_price:,.2f} MXN'

    doc.add_paragraph("\n")
    
    # Resumen Financiero
    resumen = doc.add_paragraph()
    resumen.add_run('Total Pagado: ').bold = True
    resumen.add_run(f'${order.total_price:,.2f} MXN')

    # Firma
    doc.add_paragraph("\n\n___________________________\nFirma de Emisor\nEvaristo Ochoa Luna")

    # Guardar en buffer para enviar
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return send_file(buffer, as_attachment=True, download_name=f'Recibo_Wingdicion_{order.id}.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/admin/complete_order/<int:id>', methods=['POST'])
def complete_order(id):
    order = Sale.query.get_or_404(id)
    if order.status in ['pendiente', 'recibida']:
        order.status = 'completada'
        # Deduct inventory ONLY when confirmed
        deduct_inventory(order)
        db.session.commit()
        flash(f'Orden #{id} completada y descontada del inventario.', 'success')
    return redirect(url_for('pending_orders'))


@app.route('/sales_history')
def sales_history():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    query = Sale.query
    
    if start_date:
        query = query.filter(Sale.date >= datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        # Add one day to end_date to include all sales on that day
        end_dt = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)
        query = query.filter(Sale.date < end_dt)
        
    sales_list = query.order_by(Sale.date.desc()).all()
    
    # Calculate grand total
    grand_total = sum(sale.total_price for sale in sales_list)
    
    # Group sales by day
    sales_by_day = {}
    for sale in sales_list:
        day = sale.date.strftime('%Y-%m-%d')
        if day not in sales_by_day:
            sales_by_day[day] = {'total': 0.0, 'sales': []}
        sales_by_day[day]['total'] += sale.total_price
        sales_by_day[day]['sales'].append(sale)
    
    # Sort days descending
    sorted_days = sorted(sales_by_day.keys(), reverse=True)
    
    return render_template('sales_history.html', 
                           sales_by_day=sales_by_day, 
                           sorted_days=sorted_days,
                           start_date=start_date,
                           end_date=end_date,
                           grand_total=grand_total)

@app.route('/export_sales_history')
def export_sales_history():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    query = Sale.query
    
    if start_date:
        query = query.filter(Sale.date >= datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        end_dt = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)
        query = query.filter(Sale.date < end_dt)
        
    sales_list = query.order_by(Sale.date.desc()).all()
    
    data = []
    for s in sales_list:
        data.append({
            'Fecha': s.date.strftime('%Y-%m-%d %H:%M'),
            'Producto': s.product_name,
            'Total': s.total_price
        })
    
    df = pd.DataFrame(data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Ventas')
    output.seek(0)
    
    return send_file(output, as_attachment=True, download_name='ventas_wingdicion.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/edit_sale/<int:id>', methods=['GET', 'POST'])
def edit_sale(id):
    sale = Sale.query.get_or_404(id)
    if request.method == 'POST':
        sale.product_name = request.form['product_name']
        sale.total_price = float(request.form['total_price'])
        # Update date if provided
        new_date_str = request.form.get('date')
        if new_date_str:
            try:
                # Expecting format from datetime-local input: YYYY-MM-DDTHH:MM
                sale.date = datetime.strptime(new_date_str, '%Y-%m-%dT%H:%M')
            except ValueError:
                flash('Formato de fecha no válido.', 'danger')
        
        db.session.commit()
        flash('Venta actualizada correctamente.', 'success')
        return redirect(url_for('sales_history'))
    
    # Fetch menu items for the adding items feature
    all_items = MenuItem.query.order_by(MenuItem.category, MenuItem.name).all()
    menu_data = {}
    for item in all_items:
        if item.category not in menu_data:
            menu_data[item.category] = []
        menu_data[item.category].append({'name': item.name, 'price': item.price})
    
    formatted_menu = []
    for category, items in menu_data.items():
        formatted_menu.append({'category': category, 'items': items})
        
    return render_template('edit_sale.html', sale=sale, menu=formatted_menu)

@app.route('/delete_sale/<int:id>', methods=['POST'])
def delete_sale(id):
    sale = Sale.query.get_or_404(id)
    db.session.delete(sale)
    db.session.commit()
    flash('Venta eliminada correctamente.', 'danger')
    return redirect(url_for('sales_history'))

@app.route('/ai_insights', methods=['GET', 'POST'])
def ai_insights():
    if not model:
        flash('La API Key de Google AI no está configurada. Por favor configúrala en app.py', 'warning')
        return render_template('ai_insights.html', analysis=None, inventory_tips=None, ad_post=None, topic="")

    topic = ""
    if request.method == 'POST':
        topic = request.form.get('topic', '').strip()
    else:
        topic = request.args.get('topic', '').strip()

    # 0. Intentar cargar desde el cache (a menos que se pida actualizar o haya un tema personalizado)
    refresh = request.args.get('refresh') == 'true'
    if not refresh and not topic:
        cached_data = get_cached_ai()
        if cached_data:
            return render_template('ai_insights.html', 
                                   analysis=cached_data['analysis'], 
                                   inventory_tips=cached_data['inventory_tips'], 
                                   ad_post=cached_data['ad_post'],
                                   topic="")

    try:
        # 1. Preparar datos de ventas (últimos 30 días)
        month_ago = datetime.now() - timedelta(days=30)
        recent_sales = Sale.query.filter(Sale.date >= month_ago).all()
        
        sales_summary = ""
        for s in recent_sales:
            sales_summary += f"{s.date.strftime('%Y-%m-%d')}: {s.product_name} - ${s.total_price}\n"

        # 2. Preparar datos de inventario
        inventory = InventoryItem.query.all()
        inv_summary = ""
        for i in inventory:
            inv_summary += f"{i.name}: {i.quantity} {i.unit}\n"

        # --- Prompt Combinado para ahorrar Cuota (Evitar Error 429) ---
        topic_clause = ""
        ad_prompt_clause = "Crea un post de Facebook llamativo, con emojis y hashtags, promocionando Alitas y Boneless con nuestras salsas."
        if topic:
            topic_clause = f"\nEnfócate especialmente en el siguiente tema/evento para la publicidad de hoy: {topic}.\n"
            ad_prompt_clause = f"Crea un post de Facebook llamativo, con emojis y hashtags, promocionando Alitas y Boneless con nuestras salsas, alineado fuertemente con el siguiente tema o evento especial de hoy: {topic}."

        combined_prompt = f"""
        Como experto en negocios de comida rápida y marketing, analiza los siguientes datos de mi negocio 'Wingdición' en Ciudad Juárez:

        DATOS DE VENTAS (Últimos 30 días):
        {sales_summary}

        DATOS DE INVENTARIO ACTUAL:
        {inv_summary}
        {topic_clause}
        POR FAVOR, GENERA TRES SECCIONES SEPARADAS POR EL MARCADOR '---':
        1. ANÁLISIS: Dime cuál es el día más flojo, el producto más vendido y una promoción para el día flojo.
        2. INVENTARIO: Qué productos comprar para el fin de semana basándote en las ventas y stock actual.
        3. PUBLICIDAD: {ad_prompt_clause}

        Responde exactamente en este formato:
        [ANALISIS]
        (tu análisis aquí)
        ---
        [INVENTARIO]
        (tus sugerencias aquí)
        ---
        [PUBLICIDAD]
        (tu post aquí)
        """
        
        full_response = model.generate_content(combined_prompt).text
        
        # Separar la respuesta por el marcador '---'
        parts = full_response.split('---')
        
        analysis = parts[0].replace('[ANALISIS]', '').strip() if len(parts) > 0 else "No disponible"
        inventory_tips = parts[1].replace('[INVENTARIO]', '').strip() if len(parts) > 1 else "No disponible"
        ad_post = parts[2].replace('[PUBLICIDAD]', '').strip() if len(parts) > 2 else "No disponible"

        # Guardar en cache solo si NO hay un tema personalizado
        if not topic:
            ai_results = {
                'analysis': analysis,
                'inventory_tips': inventory_tips,
                'ad_post': ad_post
            }
            save_ai_cache(ai_results)

        return render_template('ai_insights.html', 
                               analysis=analysis, 
                               inventory_tips=inventory_tips, 
                               ad_post=ad_post,
                               topic=topic)
    except Exception as e:
        print(f"AI Error: {e}")
        flash(f'Error al conectar con la IA: {str(e)}', 'danger')
        return render_template('ai_insights.html', analysis=None, inventory_tips=None, ad_post=None, topic=topic)

@app.route('/stats')
def stats():
    # Obtener fechas de filtro
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')
    
    # Valores por defecto: últimos 30 días
    now_local = get_local_time()
    default_start = (now_local - timedelta(days=30)).date()
    default_end = now_local.date()
    
    try:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date() if start_date_str else default_start
    except ValueError:
        start_date = default_start
        
    try:
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date() if end_date_str else default_end
    except ValueError:
        end_date = default_end
        
    # Convertir a datetime para la consulta (desde el inicio del start_date hasta el final del end_date)
    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())
    
    # Consultar ventas completadas en el rango
    sales = Sale.query.filter(
        Sale.date >= start_dt,
        Sale.date <= end_dt,
        Sale.status == 'completada'
    ).all()
    
    # Cargar todos los MenuItems en memoria para mapear nombres y categorías
    menu_items = MenuItem.query.all()
    # Mapeo de nombre de item a MenuItem object
    menu_map = {item.name: item for item in menu_items}
    
    # Para búsquedas más flexibles (por si hay pequeñas variaciones)
    def find_menu_item(name):
        if name in menu_map:
            return menu_map[name]
        clean_name = name.split(' (')[0].strip()
        for menu_name, item in menu_map.items():
            if clean_name in menu_name:
                return item
        return None
    
    # Inicializar acumuladores para estadísticas
    total_revenue = 0.0
    total_orders = len(sales)
    
    # 1. Ventas por Categoría: {categoria: ingresos} y {categoria: cantidad}
    sales_by_category_revenue = {}
    sales_by_category_qty = {}
    
    # 2. Top Productos: {product_name: {qty: X, revenue: Y}}
    product_stats = {}
    
    # 3. Ventas por Día de la Semana: {día_nombre: ingresos}
    days_of_week = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    sales_by_day = {day: 0.0 for day in days_of_week}
    
    # 4. Ventas por Hora: {hora: ingresos}
    sales_by_hour = {hour: 0.0 for hour in range(24)}
    orders_by_hour = {hour: 0 for hour in range(24)}
    
    # Parsear cada venta
    for sale in sales:
        total_revenue += sale.total_price
        
        # Obtener el día de la semana
        weekday = sale.date.weekday() # 0: Lunes, 6: Domingo
        day_name = days_of_week[weekday]
        sales_by_day[day_name] += sale.total_price
        
        # Obtener la hora
        hour = sale.date.hour
        sales_by_hour[hour] += sale.total_price
        orders_by_hour[hour] += 1
        
        # Parsear productos de la venta
        lines = [line.strip() for line in sale.product_name.replace(',', '\n').split('\n') if line.strip()]
        for line in lines:
            qty = 1
            name_to_check = line
            if 'x ' in line:
                try:
                    parts = line.split('x ', 1)
                    qty = int(parts[0])
                    name_to_check = parts[1].strip()
                except ValueError:
                    pass
            
            menu_item = find_menu_item(name_to_check)
            if menu_item:
                item_name = menu_item.name
                category = menu_item.category
                item_price = menu_item.price
                item_revenue = item_price * qty
                
                # Acumular por categoría
                sales_by_category_revenue[category] = sales_by_category_revenue.get(category, 0.0) + item_revenue
                sales_by_category_qty[category] = sales_by_category_qty.get(category, 0) + qty
                
                # Acumular por producto
                if item_name not in product_stats:
                    product_stats[item_name] = {'qty': 0, 'revenue': 0.0}
                product_stats[item_name]['qty'] += qty
                product_stats[item_name]['revenue'] += item_revenue

    # Ticket Promedio
    avg_ticket = total_revenue / total_orders if total_orders > 0 else 0.0
    
    # Encontrar Producto Estrella (por cantidad)
    star_product_name = "N/A"
    star_product_qty = 0
    if product_stats:
        best_product = max(product_stats.items(), key=lambda x: x[1]['qty'])
        star_product_name = best_product[0]
        star_product_qty = best_product[1]['qty']
        
    # Formatear datos de Top Productos (ordenar por cantidad descendente)
    sorted_products = sorted(product_stats.items(), key=lambda x: x[1]['qty'], reverse=True)[:10]
    top_products_labels = [item[0] for item in sorted_products]
    top_products_qty = [item[1]['qty'] for item in sorted_products]
    
    # Preparar datos para los gráficos en formato lista
    category_labels = list(sales_by_category_revenue.keys())
    category_values = list(sales_by_category_revenue.values())
    
    day_labels = days_of_week
    day_values = [sales_by_day[d] for d in days_of_week]
    
    # Filtrar horas activas para mostrar en el gráfico (de 12:00 a 23:00)
    active_hours = list(range(12, 23))
    hour_labels = [f"{h:02d}:00" for h in active_hours]
    hour_values_revenue = [sales_by_hour[h] for h in active_hours]
    
    return render_template(
        'stats.html',
        start_date=start_date.strftime('%Y-%m-%d'),
        end_date=end_date.strftime('%Y-%m-%d'),
        total_revenue=total_revenue,
        total_orders=total_orders,
        avg_ticket=avg_ticket,
        star_product_name=star_product_name,
        star_product_qty=star_product_qty,
        category_labels=category_labels,
        category_values=category_values,
        day_labels=day_labels,
        day_values=day_values,
        hour_labels=hour_labels,
        hour_values_revenue=hour_values_revenue,
        top_products_labels=top_products_labels,
        top_products_qty=top_products_qty
    )

def get_unit_cost(item_id):
    # 1. Intentar buscar en el Listado Maestro (InfoRecord) el precio activo para costeo
    default_pir = InfoRecord.query.filter_by(inventory_item_id=item_id, is_default=True).first()
    if default_pir:
        return default_pir.cost_per_unit
        
    # 2. Fallback: Obtener promedio ponderado de las compras de este artículo
    purchases = Purchase.query.filter_by(item_id=item_id).all()
    if not purchases:
        return 0.0
    total_qty = sum(p.quantity_bought for p in purchases)
    total_cost = sum(p.cost for p in purchases)
    return total_cost / total_qty if total_qty > 0 else 0.0

@app.route('/recipes')
def recipes():
    all_items = MenuItem.query.order_by(MenuItem.category, MenuItem.name).all()
    
    # Calcular costos y márgenes para cada MenuItem
    recipes_data = []
    for item in all_items:
        recipe_cost = 0.0
        missing_cost_history = False
        
        for bom in item.bom_items:
            unit_cost = get_unit_cost(bom.inventory_item_id)
            if unit_cost == 0.0:
                missing_cost_history = True
            recipe_cost += bom.quantity * unit_cost
            
        profit = item.price - recipe_cost
        margin_pct = (profit / item.price * 100) if item.price > 0 else 0.0
        
        recipes_data.append({
            'item': item,
            'cost': recipe_cost,
            'profit': profit,
            'margin_pct': margin_pct,
            'missing_cost_history': missing_cost_history
        })
        
    return render_template('recipes.html', recipes=recipes_data)

@app.route('/recipes/edit/<int:id>', methods=['GET', 'POST'])
def edit_recipe(id):
    menu_item = MenuItem.query.get_or_404(id)
    
    if request.method == 'POST':
        action = request.form.get('action')
        
        if action == 'add':
            inv_item_id = int(request.form.get('inventory_item_id'))
            qty = float(request.form.get('quantity'))
            
            # Verificar si ya existe en la receta
            existing_bom = BOMItem.query.filter_by(menu_item_id=menu_item.id, inventory_item_id=inv_item_id).first()
            if existing_bom:
                existing_bom.quantity += qty
            else:
                new_bom = BOMItem(menu_item_id=menu_item.id, inventory_item_id=inv_item_id, quantity=qty)
                db.session.add(new_bom)
            
            db.session.commit()
            flash('Ingrediente agregado a la receta.', 'success')
            
        elif action == 'update':
            bom_id = int(request.form.get('bom_id'))
            qty = float(request.form.get('quantity'))
            
            bom_item = BOMItem.query.get_or_404(bom_id)
            if qty <= 0:
                db.session.delete(bom_item)
                flash('Ingrediente eliminado de la receta.', 'warning')
            else:
                bom_item.quantity = qty
                flash('Cantidad actualizada correctamente.', 'success')
            
            db.session.commit()
            
        elif action == 'delete':
            bom_id = int(request.form.get('bom_id'))
            bom_item = BOMItem.query.get_or_404(bom_id)
            db.session.delete(bom_item)
            db.session.commit()
            flash('Ingrediente eliminado de la receta.', 'danger')
            
        return redirect(url_for('edit_recipe', id=menu_item.id))
        
    # Obtener ingredientes actuales en la receta
    current_bom = menu_item.bom_items
    current_inv_ids = [bom.inventory_item_id for bom in current_bom]
    
    # Obtener ingredientes disponibles para agregar (los que no están en la receta)
    available_ingredients = InventoryItem.query.filter(~InventoryItem.id.in_(current_inv_ids)).all()
    
    # Calcular desglose detallado con costos unitarios
    bom_details = []
    total_recipe_cost = 0.0
    for bom in current_bom:
        unit_cost = get_unit_cost(bom.inventory_item_id)
        total_item_cost = bom.quantity * unit_cost
        total_recipe_cost += total_item_cost
        
        bom_details.append({
            'bom': bom,
            'unit_cost': unit_cost,
            'total_cost': total_item_cost
        })
        
    return render_template(
        'edit_recipe.html',
        menu_item=menu_item,
        bom_details=bom_details,
        total_recipe_cost=total_recipe_cost,
        available_ingredients=available_ingredients
    )

@app.route('/delete_purchase/<int:id>', methods=['POST'])
def delete_purchase(id):
    purchase = Purchase.query.get_or_404(id)
    
    # Revertir cantidad del inventario
    item = InventoryItem.query.get(purchase.item_id)
    if item:
        item.quantity -= purchase.quantity_bought
        
    db.session.delete(purchase)
    db.session.commit()
    flash('Compra eliminada y el inventario ha sido ajustado (restado) correctamente.', 'warning')
    
    referrer = request.referrer or url_for('purchases')
    return redirect(referrer)

@app.route('/edit_purchase/<int:id>', methods=['GET', 'POST'])
def edit_purchase(id):
    purchase = Purchase.query.get_or_404(id)
    
    if request.method == 'POST':
        new_item_id = int(request.form['item_id'])
        new_qty = float(request.form['quantity'])
        new_cost = float(request.form['cost'])
        
        # Procesar fecha personalizada si se proporciona
        new_date_str = request.form.get('date')
        new_date = purchase.date
        if new_date_str:
            try:
                new_date = datetime.strptime(new_date_str, '%Y-%m-%dT%H:%M')
            except ValueError:
                pass
        
        old_item_id = purchase.item_id
        old_qty = purchase.quantity_bought
        
        old_item = InventoryItem.query.get(old_item_id)
        new_item = InventoryItem.query.get(new_item_id)
        
        # Ajustar inventario
        if old_item_id == new_item_id:
            if old_item:
                diff = new_qty - old_qty
                old_item.quantity += diff
        else:
            if old_item:
                old_item.quantity -= old_qty
            if new_item:
                new_item.quantity += new_qty
                
        # Actualizar registro
        purchase.item_id = new_item_id
        purchase.quantity_bought = new_qty
        purchase.cost = new_cost
        purchase.date = new_date
        purchase.provider = request.form.get('provider')
        
        db.session.commit()
        flash('Compra editada con éxito y el inventario ha sido actualizado.', 'success')
        return redirect(url_for('purchases'))
        
    items = InventoryItem.query.all()
    formatted_date = purchase.date.strftime('%Y-%m-%dT%H:%M')
    return render_template('edit_purchase.html', purchase=purchase, items=items, formatted_date=formatted_date)

@app.route('/edit_inventory_item/<int:id>', methods=['GET', 'POST'])
def edit_inventory_item(id):
    item = InventoryItem.query.get_or_404(id)
    
    if request.method == 'POST':
        item.name = request.form['name']
        item.category = request.form['category']
        item.unit = request.form['unit']
        item.quantity = float(request.form['quantity'])
        item.provider = request.form.get('provider')
        
        db.session.commit()
        flash('Artículo de inventario actualizado correctamente.', 'success')
        return redirect(url_for('inventory'))
        
    return render_template('edit_inventory_item.html', item=item)

@app.route('/delete_inventory_item/<int:id>', methods=['POST'])
def delete_inventory_item(id):
    item = InventoryItem.query.get_or_404(id)
    
    # Validar integridad: no estar en recetas (BOMItem)
    if item.bom_usages:
        flash(f'No se puede eliminar "{item.name}" porque se utiliza en recetas (BOM). Elimina el artículo de las recetas correspondientes primero.', 'danger')
        return redirect(url_for('inventory'))
        
    # Validar integridad: no tener compras registradas
    if item.purchases:
        flash(f'No se puede eliminar "{item.name}" porque tiene compras registradas en el historial. Elimina las compras primero.', 'danger')
        return redirect(url_for('inventory'))
        
    db.session.delete(item)
    db.session.commit()
    flash('Artículo de inventario eliminado correctamente.', 'warning')
    return redirect(url_for('inventory'))

@app.route('/consumptions')
def consumptions():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    query = Consumption.query
    
    if start_date:
        query = query.filter(Consumption.date >= datetime.strptime(start_date, '%Y-%m-%d'))
    if end_date:
        # Include the whole end date day by adding 23:59:59
        end_dt = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1) - timedelta(seconds=1)
        query = query.filter(Consumption.date <= end_dt)
        
    consumptions_list = query.order_by(Consumption.date.desc()).all()
    return render_template('consumptions.html', consumptions=consumptions_list, start_date=start_date, end_date=end_date)

@app.route('/info_records', methods=['GET', 'POST'])
def info_records():
    if request.method == 'POST':
        item_id = int(request.form['inventory_item_id'])
        provider = request.form['provider'].strip()
        cost_per_unit = float(request.form['cost_per_unit'])
        is_default = 'is_default' in request.form
        
        # Si se marca como default, quitar default de otros del mismo item
        if is_default:
            InfoRecord.query.filter_by(inventory_item_id=item_id).update({InfoRecord.is_default: False})
            
        # Verificar si ya existe este proveedor para este item
        existing = InfoRecord.query.filter_by(inventory_item_id=item_id, provider=provider).first()
        if existing:
            existing.cost_per_unit = cost_per_unit
            existing.is_default = is_default
            db.session.commit()
            flash(f'Registro Maestro actualizado para {existing.inventory_item.name} de {provider}.', 'success')
        else:
            new_record = InfoRecord(
                inventory_item_id=item_id,
                provider=provider,
                cost_per_unit=cost_per_unit,
                is_default=is_default
            )
            db.session.add(new_record)
            db.session.commit()
            flash('Nuevo Registro Maestro (PIR) creado con éxito.', 'success')
            
        return redirect(url_for('info_records'))
        
    items = InventoryItem.query.all()
    records = InfoRecord.query.order_by(InfoRecord.inventory_item_id).all()
    return render_template('info_records.html', items=items, records=records)

@app.route('/info_records/edit/<int:id>', methods=['GET', 'POST'])
def edit_info_record(id):
    record = InfoRecord.query.get_or_404(id)
    if request.method == 'POST':
        cost_per_unit = float(request.form['cost_per_unit'])
        provider = request.form['provider'].strip()
        is_default = 'is_default' in request.form
        
        if is_default and not record.is_default:
            InfoRecord.query.filter_by(inventory_item_id=record.inventory_item_id).update({InfoRecord.is_default: False})
            
        record.cost_per_unit = cost_per_unit
        record.provider = provider
        record.is_default = is_default
        db.session.commit()
        flash('Registro Maestro actualizado.', 'success')
        return redirect(url_for('info_records'))
        
    return render_template('edit_info_record.html', record=record)

@app.route('/info_records/delete/<int:id>', methods=['POST'])
def delete_info_record(id):
    record = InfoRecord.query.get_or_404(id)
    was_default = record.is_default
    item_id = record.inventory_item_id
    
    db.session.delete(record)
    db.session.commit()
    
    # Si borramos el default, e intentamos que otro pase a ser default
    if was_default:
        next_record = InfoRecord.query.filter_by(inventory_item_id=item_id).first()
        if next_record:
            next_record.is_default = True
            db.session.commit()
            
    flash('Registro Maestro eliminado.', 'warning')
    return redirect(url_for('info_records'))

@app.route('/info_records/make_default/<int:id>', methods=['POST'])
def make_default_info_record(id):
    record = InfoRecord.query.get_or_404(id)
    item_id = record.inventory_item_id
    
    # Poner todos los del mismo item en False
    InfoRecord.query.filter_by(inventory_item_id=item_id).update({InfoRecord.is_default: False})
    
    record.is_default = True
    db.session.commit()
    flash(f'"{record.inventory_item.name}" costeará por defecto con el precio de {record.provider}.', 'success')
    return redirect(url_for('info_records'))

@app.route('/debug_models')
def debug_models():
    try:
        models_list = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                models_list.append(m.name)
        return {"available_models": models_list}
    except Exception as e:
        return {"error": str(e)}

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
