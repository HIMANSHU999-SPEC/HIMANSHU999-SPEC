from io import BytesIO
from datetime import datetime as dt
from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify, send_file
from flask_login import login_required, current_user
from sqlalchemy import func
from app import db
from app.models import Stock, Campus, StockHistory, StockTransfer
from app.forms import StockForm, CampusForm, StockTransferForm

stock_bp = Blueprint('stock', __name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def log_stock_action(stock, action, changed_by, field_changed=None, old_value=None, new_value=None):
    entry = StockHistory(
        stock_id=stock.id if stock else None,
        item_name=stock.item_name if stock else 'N/A',
        campus_name=stock.campus.name if stock and stock.campus else None,
        action=action,
        field_changed=field_changed,
        old_value=str(old_value) if old_value is not None else None,
        new_value=str(new_value) if new_value is not None else None,
        changed_by=changed_by,
    )
    db.session.add(entry)


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

@stock_bp.route('/dashboard')
@login_required
def dashboard():
    campuses = Campus.query.order_by(Campus.name).all()
    campus_stats = []
    total_items = 0
    total_value = 0
    low_stock_count = 0
    category_data = {}
    campus_labels = []
    campus_values = []

    for campus in campuses:
        stocks = Stock.query.filter_by(campus_id=campus.id).all()
        item_count = len(stocks)
        value = sum((s.quantity or 0) * (s.unit_price or 0) for s in stocks)
        campus_low = sum(1 for s in stocks if s.is_low_stock)
        total_items += item_count
        total_value += value
        low_stock_count += campus_low
        campus_labels.append(campus.name)
        campus_values.append(round(value, 2))

        for s in stocks:
            cat = s.category or 'Uncategorized'
            category_data[cat] = category_data.get(cat, 0) + (s.quantity or 0)

        campus_stats.append({
            'campus': campus,
            'item_count': item_count,
            'total_value': value,
            'low_stock': campus_low,
        })

    low_stock_items = Stock.query.filter(
        Stock.quantity <= Stock.low_stock_threshold
    ).order_by(Stock.quantity.asc()).limit(20).all()

    recent_activity = StockHistory.query.order_by(
        StockHistory.timestamp.desc()
    ).limit(10).all()

    return render_template('dashboard.html',
                           campus_stats=campus_stats,
                           total_items=total_items,
                           total_value=total_value,
                           low_stock_count=low_stock_count,
                           low_stock_items=low_stock_items,
                           recent_activity=recent_activity,
                           category_labels=list(category_data.keys()),
                           category_values=list(category_data.values()),
                           campus_labels=campus_labels,
                           campus_values=campus_values)


# ---------------------------------------------------------------------------
# Charts API
# ---------------------------------------------------------------------------

@stock_bp.route('/api/charts')
@login_required
def charts_api():
    cat_rows = db.session.query(
        Stock.category, func.sum(Stock.quantity)
    ).group_by(Stock.category).all()
    category_labels = [r[0] or 'Uncategorized' for r in cat_rows]
    category_values = [int(r[1] or 0) for r in cat_rows]

    campuses = Campus.query.order_by(Campus.name).all()
    campus_labels = []
    campus_values = []
    for c in campuses:
        stocks = Stock.query.filter_by(campus_id=c.id).all()
        val = sum((s.quantity or 0) * (s.unit_price or 0) for s in stocks)
        campus_labels.append(c.name)
        campus_values.append(round(val, 2))

    cond_rows = db.session.query(
        Stock.condition, func.count(Stock.id)
    ).group_by(Stock.condition).all()
    condition_labels = [r[0] or 'Unknown' for r in cond_rows]
    condition_values = [int(r[1] or 0) for r in cond_rows]

    return jsonify({
        'category': {'labels': category_labels, 'values': category_values},
        'campus': {'labels': campus_labels, 'values': campus_values},
        'condition': {'labels': condition_labels, 'values': condition_values},
    })


# ---------------------------------------------------------------------------
# Campus CRUD
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Campus Stocks View
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Stock CRUD (with audit trail)
# ---------------------------------------------------------------------------

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
            low_stock_threshold=form.low_stock_threshold.data if form.low_stock_threshold.data is not None else 10,
            campus_id=form.campus_id.data,
            remarks=form.remarks.data,
            added_by=current_user.username,
        )
        db.session.add(stock)
        db.session.flush()
        log_stock_action(stock, 'created', current_user.username)
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
        changes = []
        if stock.item_name != form.item_name.data:
            changes.append(('item_name', stock.item_name, form.item_name.data))
        if stock.category != form.category.data:
            changes.append(('category', stock.category, form.category.data))
        if stock.quantity != (form.quantity.data or 0):
            changes.append(('quantity', stock.quantity, form.quantity.data or 0))
        if stock.unit_price != (form.unit_price.data or 0.0):
            changes.append(('unit_price', stock.unit_price, form.unit_price.data or 0.0))
        if stock.condition != form.condition.data:
            changes.append(('condition', stock.condition, form.condition.data))
        if stock.campus_id != form.campus_id.data:
            old_campus = db.session.get(Campus, stock.campus_id)
            new_campus = db.session.get(Campus, form.campus_id.data)
            changes.append(('campus', old_campus.name if old_campus else str(stock.campus_id),
                            new_campus.name if new_campus else str(form.campus_id.data)))

        stock.item_name = form.item_name.data
        stock.category = form.category.data
        stock.quantity = form.quantity.data or 0
        stock.unit = form.unit.data
        stock.unit_price = form.unit_price.data or 0.0
        stock.total_value = stock.quantity * stock.unit_price
        stock.condition = form.condition.data
        stock.low_stock_threshold = form.low_stock_threshold.data if form.low_stock_threshold.data is not None else 10
        stock.campus_id = form.campus_id.data
        stock.remarks = form.remarks.data

        if changes:
            for field, old, new in changes:
                log_stock_action(stock, 'updated', current_user.username, field, old, new)
        else:
            log_stock_action(stock, 'updated', current_user.username)

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
        log_stock_action(stock, 'deleted', current_user.username)
        db.session.delete(stock)
        db.session.commit()
        flash('Stock item deleted.', 'success')
        return redirect(url_for('stock.campus_stocks', campus_id=campus_id))
    return redirect(url_for('stock.dashboard'))


# ---------------------------------------------------------------------------
# Global Quick Search
# ---------------------------------------------------------------------------

@stock_bp.route('/search')
@login_required
def global_search():
    q = request.args.get('q', '').strip()
    results = []
    if q:
        results = Stock.query.filter(
            Stock.item_name.ilike(f'%{q}%') |
            Stock.category.ilike(f'%{q}%') |
            Stock.remarks.ilike(f'%{q}%')
        ).order_by(Stock.item_name).all()
    return render_template('search_results.html', query=q, results=results)


# ---------------------------------------------------------------------------
# Stock Transfer Between Campuses
# ---------------------------------------------------------------------------

@stock_bp.route('/transfer/<int:campus_id>', methods=['GET', 'POST'])
@login_required
def transfer_stock(campus_id):
    campus = db.session.get(Campus, campus_id)
    if not campus:
        flash('Campus not found.', 'danger')
        return redirect(url_for('stock.dashboard'))

    form = StockTransferForm()
    stocks = Stock.query.filter_by(campus_id=campus_id).order_by(Stock.item_name).all()
    form.stock_id.choices = [(s.id, f"{s.item_name} (Qty: {s.quantity})") for s in stocks]
    other_campuses = Campus.query.filter(Campus.id != campus_id).order_by(Campus.name).all()
    form.to_campus_id.choices = [(c.id, f"{c.name} ({c.code})") for c in other_campuses]

    if form.validate_on_submit():
        stock = db.session.get(Stock, form.stock_id.data)
        to_campus = db.session.get(Campus, form.to_campus_id.data)
        qty = form.quantity.data

        if not stock or stock.campus_id != campus_id:
            flash('Invalid stock item.', 'danger')
        elif not to_campus:
            flash('Invalid destination campus.', 'danger')
        elif qty > stock.quantity:
            flash(f'Cannot transfer {qty}. Only {stock.quantity} available.', 'danger')
        else:
            stock.quantity -= qty
            stock.total_value = stock.quantity * (stock.unit_price or 0)

            dest_stock = Stock.query.filter_by(
                item_name=stock.item_name, campus_id=to_campus.id, category=stock.category
            ).first()
            if dest_stock:
                dest_stock.quantity += qty
                dest_stock.total_value = dest_stock.quantity * (dest_stock.unit_price or 0)
            else:
                dest_stock = Stock(
                    item_name=stock.item_name, category=stock.category, quantity=qty,
                    unit=stock.unit, unit_price=stock.unit_price,
                    total_value=qty * (stock.unit_price or 0), condition=stock.condition,
                    low_stock_threshold=stock.low_stock_threshold,
                    campus_id=to_campus.id, remarks=f'Transferred from {campus.name}',
                    added_by=current_user.username,
                )
                db.session.add(dest_stock)

            transfer = StockTransfer(
                stock_id=stock.id, item_name=stock.item_name,
                quantity_transferred=qty, from_campus_id=campus_id,
                to_campus_id=to_campus.id, transferred_by=current_user.username,
                remarks=form.remarks.data,
            )
            db.session.add(transfer)
            log_stock_action(stock, 'transferred_out', current_user.username,
                             'quantity', stock.quantity + qty, stock.quantity)
            db.session.flush()
            if dest_stock.id:
                log_stock_action(dest_stock, 'transferred_in', current_user.username,
                                 'quantity', dest_stock.quantity - qty, dest_stock.quantity)
            db.session.commit()
            flash(f'Transferred {qty} x "{stock.item_name}" to {to_campus.name}.', 'success')
            return redirect(url_for('stock.campus_stocks', campus_id=campus_id))

    return render_template('transfer_stock.html', form=form, campus=campus)


# ---------------------------------------------------------------------------
# Activity Log
# ---------------------------------------------------------------------------

@stock_bp.route('/activity')
@login_required
def activity_log():
    page = request.args.get('page', 1, type=int)
    action_filter = request.args.get('action', '').strip()
    query = StockHistory.query
    if action_filter:
        query = query.filter(StockHistory.action == action_filter)
    logs = query.order_by(StockHistory.timestamp.desc()).paginate(
        page=page, per_page=50, error_out=False)
    return render_template('activity_log.html', logs=logs, action_filter=action_filter)


# ---------------------------------------------------------------------------
# PDF Report Export
# ---------------------------------------------------------------------------

@stock_bp.route('/report/pdf/<int:campus_id>')
@login_required
def download_pdf(campus_id):
    campus = db.session.get(Campus, campus_id)
    if not campus:
        flash('Campus not found.', 'danger')
        return redirect(url_for('stock.dashboard'))
    stocks = Stock.query.filter_by(campus_id=campus_id).order_by(Stock.category, Stock.item_name).all()
    html = _build_pdf_html(campus, stocks)
    buf = BytesIO(html.encode('utf-8'))
    buf.seek(0)
    return send_file(buf, mimetype='text/html', as_attachment=True,
                     download_name=f'stock_report_{campus.code}.html')


@stock_bp.route('/report/pdf/all')
@login_required
def download_pdf_all():
    campuses = Campus.query.order_by(Campus.name).all()
    parts = []
    for campus in campuses:
        stocks = Stock.query.filter_by(campus_id=campus.id).order_by(Stock.category, Stock.item_name).all()
        parts.append(_build_pdf_section(campus, stocks))
    html = _wrap_pdf_html('All Campuses Stock Report', '\n'.join(parts))
    buf = BytesIO(html.encode('utf-8'))
    buf.seek(0)
    return send_file(buf, mimetype='text/html', as_attachment=True,
                     download_name='stock_report_all_campuses.html')


def _build_pdf_html(campus, stocks):
    return _wrap_pdf_html(f'Stock Report - {campus.name} ({campus.code})',
                          _build_pdf_section(campus, stocks))


def _build_pdf_section(campus, stocks):
    now = dt.utcnow().strftime('%Y-%m-%d %H:%M UTC')
    grand_total = 0
    rows = ''
    for i, s in enumerate(stocks, 1):
        tv = (s.quantity or 0) * (s.unit_price or 0)
        grand_total += tv
        low_flag = ' style="background:#ffe0e0;"' if s.is_low_stock else ''
        rows += (f'<tr{low_flag}><td>{i}</td><td>{s.item_name}</td>'
                 f'<td>{s.category or "-"}</td><td>{s.quantity}</td><td>{s.unit or "-"}</td>'
                 f'<td>{s.unit_price or 0:.2f}</td><td>{tv:.2f}</td>'
                 f'<td>{s.condition}</td><td>{s.remarks or "-"}</td></tr>')
    return (f'<h2>{campus.name} ({campus.code})</h2>'
            f'<p>Address: {campus.address or "N/A"} | Items: {len(stocks)} | Generated: {now}</p>'
            f'<table><thead><tr><th>#</th><th>Item</th><th>Category</th><th>Qty</th><th>Unit</th>'
            f'<th>Price</th><th>Total</th><th>Condition</th><th>Remarks</th></tr></thead>'
            f'<tbody>{rows}</tbody>'
            f'<tfoot><tr><td colspan="6" style="text-align:right;font-weight:bold;">Grand Total:</td>'
            f'<td style="font-weight:bold;">{grand_total:.2f}</td><td colspan="2"></td></tr></tfoot></table>')


def _wrap_pdf_html(title, body):
    return f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>{title}</title>
<style>
@media print {{ @page {{ margin: 1cm; }} }}
body {{ font-family: Arial, sans-serif; margin: 20px; color: #333; }}
h1 {{ color: #1a5276; border-bottom: 2px solid #1a5276; padding-bottom: 8px; }}
h2 {{ color: #2e86c1; margin-top: 30px; }}
table {{ width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px; }}
th, td {{ border: 1px solid #ccc; padding: 6px 10px; text-align: left; }}
th {{ background: #2e86c1; color: #fff; }}
tr:nth-child(even) {{ background: #f2f2f2; }}
.print-btn {{ background: #2e86c1; color: #fff; border: none; padding: 10px 25px;
    cursor: pointer; font-size: 14px; border-radius: 5px; margin-bottom: 15px; }}
.print-btn:hover {{ background: #1a5276; }}
@media print {{ .print-btn {{ display: none; }} }}
</style></head><body>
<button class="print-btn" onclick="window.print()">Print / Save as PDF</button>
<h1>{title}</h1>{body}</body></html>'''
