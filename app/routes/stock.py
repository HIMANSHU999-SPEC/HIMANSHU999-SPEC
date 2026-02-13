from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required, current_user
from app import db
from app.models import Stock, Campus
from app.forms import StockForm, CampusForm

stock_bp = Blueprint('stock', __name__)


@stock_bp.route('/dashboard')
@login_required
def dashboard():
    campuses = Campus.query.order_by(Campus.name).all()
    campus_stats = []
    total_items = 0
    total_value = 0

    for campus in campuses:
        stocks = Stock.query.filter_by(campus_id=campus.id).all()
        item_count = len(stocks)
        value = sum((s.quantity or 0) * (s.unit_price or 0) for s in stocks)
        total_items += item_count
        total_value += value
        campus_stats.append({
            'campus': campus,
            'item_count': item_count,
            'total_value': value,
        })

    return render_template('dashboard.html',
                           campus_stats=campus_stats,
                           total_items=total_items,
                           total_value=total_value)


@stock_bp.route('/campus/add', methods=['GET', 'POST'])
@login_required
def add_campus():
    if not current_user.is_admin():
        flash('Only admins can add campuses.', 'danger')
        return redirect(url_for('stock.dashboard'))

    form = CampusForm()
    if form.validate_on_submit():
        if Campus.query.filter_by(code=form.code.data).first():
            flash('Campus code already exists.', 'danger')
        else:
            campus = Campus(
                name=form.name.data,
                code=form.code.data.upper(),
                address=form.address.data,
            )
            db.session.add(campus)
            db.session.commit()
            flash(f'Campus "{campus.name}" added successfully!', 'success')
            return redirect(url_for('stock.dashboard'))

    return render_template('campus_form.html', form=form, title='Add Campus')


@stock_bp.route('/campus/<int:campus_id>/edit', methods=['GET', 'POST'])
@login_required
def edit_campus(campus_id):
    if not current_user.is_admin():
        flash('Only admins can edit campuses.', 'danger')
        return redirect(url_for('stock.dashboard'))

    campus = db.session.get(Campus, campus_id)
    if not campus:
        flash('Campus not found.', 'danger')
        return redirect(url_for('stock.dashboard'))

    form = CampusForm(obj=campus)
    if form.validate_on_submit():
        campus.name = form.name.data
        campus.code = form.code.data.upper()
        campus.address = form.address.data
        db.session.commit()
        flash(f'Campus "{campus.name}" updated.', 'success')
        return redirect(url_for('stock.dashboard'))

    return render_template('campus_form.html', form=form, title='Edit Campus')


@stock_bp.route('/campus/<int:campus_id>/delete', methods=['POST'])
@login_required
def delete_campus(campus_id):
    if not current_user.is_admin():
        flash('Only admins can delete campuses.', 'danger')
        return redirect(url_for('stock.dashboard'))

    campus = db.session.get(Campus, campus_id)
    if campus:
        db.session.delete(campus)
        db.session.commit()
        flash(f'Campus "{campus.name}" and all its stock deleted.', 'success')
    return redirect(url_for('stock.dashboard'))


@stock_bp.route('/campus/<int:campus_id>/stocks')
@login_required
def campus_stocks(campus_id):
    campus = db.session.get(Campus, campus_id)
    if not campus:
        flash('Campus not found.', 'danger')
        return redirect(url_for('stock.dashboard'))

    search = request.args.get('search', '').strip()
    category_filter = request.args.get('category', '').strip()

    query = Stock.query.filter_by(campus_id=campus_id)
    if search:
        query = query.filter(Stock.item_name.ilike(f'%{search}%'))
    if category_filter:
        query = query.filter(Stock.category == category_filter)

    stocks = query.order_by(Stock.category, Stock.item_name).all()

    categories = db.session.query(Stock.category).filter_by(campus_id=campus_id)\
        .distinct().order_by(Stock.category).all()
    categories = [c[0] for c in categories if c[0]]

    return render_template('campus_stocks.html', campus=campus, stocks=stocks,
                           categories=categories, search=search, category_filter=category_filter)


@stock_bp.route('/stock/add', methods=['GET', 'POST'])
@login_required
def add_stock():
    form = StockForm()
    form.campus_id.choices = [(c.id, f"{c.name} ({c.code})") for c in Campus.query.order_by(Campus.name).all()]

    if form.validate_on_submit():
        quantity = form.quantity.data or 0
        unit_price = form.unit_price.data or 0.0
        stock = Stock(
            item_name=form.item_name.data,
            category=form.category.data,
            quantity=quantity,
            unit=form.unit.data,
            unit_price=unit_price,
            total_value=quantity * unit_price,
            condition=form.condition.data,
            campus_id=form.campus_id.data,
            remarks=form.remarks.data,
            added_by=current_user.username,
        )
        db.session.add(stock)
        db.session.commit()
        flash(f'Stock item "{stock.item_name}" added.', 'success')
        return redirect(url_for('stock.campus_stocks', campus_id=stock.campus_id))

    return render_template('stock_form.html', form=form, title='Add Stock Item')


@stock_bp.route('/stock/<int:stock_id>/edit', methods=['GET', 'POST'])
@login_required
def edit_stock(stock_id):
    stock = db.session.get(Stock, stock_id)
    if not stock:
        flash('Stock item not found.', 'danger')
        return redirect(url_for('stock.dashboard'))

    form = StockForm(obj=stock)
    form.campus_id.choices = [(c.id, f"{c.name} ({c.code})") for c in Campus.query.order_by(Campus.name).all()]

    if form.validate_on_submit():
        stock.item_name = form.item_name.data
        stock.category = form.category.data
        stock.quantity = form.quantity.data or 0
        stock.unit = form.unit.data
        stock.unit_price = form.unit_price.data or 0.0
        stock.total_value = stock.quantity * stock.unit_price
        stock.condition = form.condition.data
        stock.campus_id = form.campus_id.data
        stock.remarks = form.remarks.data
        db.session.commit()
        flash(f'Stock item "{stock.item_name}" updated.', 'success')
        return redirect(url_for('stock.campus_stocks', campus_id=stock.campus_id))

    return render_template('stock_form.html', form=form, title='Edit Stock Item')


@stock_bp.route('/stock/<int:stock_id>/delete', methods=['POST'])
@login_required
def delete_stock(stock_id):
    stock = db.session.get(Stock, stock_id)
    if stock:
        campus_id = stock.campus_id
        db.session.delete(stock)
        db.session.commit()
        flash('Stock item deleted.', 'success')
        return redirect(url_for('stock.campus_stocks', campus_id=campus_id))
    return redirect(url_for('stock.dashboard'))
