from app import db
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime


class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='staff')  # admin / staff
    department = db.Column(db.String(100), nullable=True)
    email = db.Column(db.String(120), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Assets assigned to this user
    assigned_assets = db.relationship('Stock', backref='assigned_user', lazy=True,
                                      foreign_keys='Stock.assigned_to')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password, method='pbkdf2:sha256')

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def is_admin(self):
        return self.role == 'admin'

    def __repr__(self):
        return f'<User {self.username}>'


class Campus(db.Model):
    __tablename__ = 'campuses'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), unique=True, nullable=False)
    code = db.Column(db.String(20), unique=True, nullable=False)
    address = db.Column(db.String(300), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    stocks = db.relationship('Stock', backref='campus', lazy=True, cascade='all, delete-orphan')

    def __repr__(self):
        return f'<Campus {self.name}>'


class Stock(db.Model):
    __tablename__ = 'stocks'

    id = db.Column(db.Integer, primary_key=True)
    item_name = db.Column(db.String(200), nullable=False)
    category = db.Column(db.String(100), nullable=True)
    quantity = db.Column(db.Integer, nullable=False, default=0)
    unit = db.Column(db.String(50), nullable=True)  # pcs, kg, litre, etc.
    unit_price = db.Column(db.Float, nullable=True, default=0.0)
    total_value = db.Column(db.Float, nullable=True, default=0.0)
    condition = db.Column(db.String(50), nullable=True, default='Good')  # Good / Damaged / Needs Repair
    low_stock_threshold = db.Column(db.Integer, nullable=True, default=10)
    remarks = db.Column(db.String(500), nullable=True)
    campus_id = db.Column(db.Integer, db.ForeignKey('campuses.id'), nullable=False)
    added_by = db.Column(db.String(80), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # --- Microsoft Lists IT Asset Management fields ---
    asset_tag = db.Column(db.String(100), nullable=True, unique=True)
    serial_number = db.Column(db.String(200), nullable=True)
    manufacturer = db.Column(db.String(150), nullable=True)
    model = db.Column(db.String(150), nullable=True)
    purchase_date = db.Column(db.Date, nullable=True)
    warranty_expiry = db.Column(db.Date, nullable=True)
    department = db.Column(db.String(100), nullable=True)
    status = db.Column(db.String(50), nullable=True, default='Active')
    # Active / In Storage / Retired / Under Repair / Lost-Stolen / Disposed

    # Staff assignment
    assigned_to = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)

    @property
    def is_low_stock(self):
        threshold = self.low_stock_threshold if self.low_stock_threshold is not None else 10
        return (self.quantity or 0) <= threshold

    @property
    def is_warranty_expired(self):
        if self.warranty_expiry:
            return self.warranty_expiry < datetime.utcnow().date()
        return False

    def __repr__(self):
        return f'<Stock {self.item_name} @ {self.campus.name if self.campus else "N/A"}>'


class StockHistory(db.Model):
    """Audit trail for all stock changes."""
    __tablename__ = 'stock_history'

    id = db.Column(db.Integer, primary_key=True)
    stock_id = db.Column(db.Integer, nullable=True)
    item_name = db.Column(db.String(200), nullable=False)
    campus_name = db.Column(db.String(120), nullable=True)
    action = db.Column(db.String(50), nullable=False)  # created, updated, deleted, transferred_out, transferred_in, assigned, unassigned
    field_changed = db.Column(db.String(100), nullable=True)
    old_value = db.Column(db.String(500), nullable=True)
    new_value = db.Column(db.String(500), nullable=True)
    changed_by = db.Column(db.String(80), nullable=False)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<StockHistory {self.action} {self.item_name} by {self.changed_by}>'


class StockTransfer(db.Model):
    """Track stock transfers between campuses."""
    __tablename__ = 'stock_transfers'

    id = db.Column(db.Integer, primary_key=True)
    stock_id = db.Column(db.Integer, db.ForeignKey('stocks.id'), nullable=True)
    item_name = db.Column(db.String(200), nullable=False)
    quantity_transferred = db.Column(db.Integer, nullable=False)
    from_campus_id = db.Column(db.Integer, db.ForeignKey('campuses.id'), nullable=False)
    to_campus_id = db.Column(db.Integer, db.ForeignKey('campuses.id'), nullable=False)
    transferred_by = db.Column(db.String(80), nullable=False)
    remarks = db.Column(db.String(500), nullable=True)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)

    from_campus = db.relationship('Campus', foreign_keys=[from_campus_id])
    to_campus = db.relationship('Campus', foreign_keys=[to_campus_id])

    def __repr__(self):
        return f'<Transfer {self.item_name} x{self.quantity_transferred}>'
