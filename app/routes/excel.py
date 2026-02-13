import os
from io import BytesIO
from flask import Blueprint, render_template, redirect, url_for, flash, request, send_file, current_app
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from app import db
from app.models import Stock, Campus
from app.forms import UploadExcelForm

excel_bp = Blueprint('excel', __name__)

EXPECTED_COLUMNS = ['item_name', 'category', 'quantity', 'unit', 'unit_price', 'condition', 'remarks']


@excel_bp.route('/upload', methods=['GET', 'POST'])
@login_required
def upload_excel():
    form = UploadExcelForm()
    form.campus_id.choices = [(c.id, f"{c.name} ({c.code})") for c in Campus.query.order_by(Campus.name).all()]

    if form.validate_on_submit():
        file = form.file.data
        campus_id = form.campus_id.data
        campus = db.session.get(Campus, campus_id)
        if not campus:
            flash('Selected campus not found.', 'danger')
            return redirect(url_for('excel.upload_excel'))

        filename = secure_filename(file.filename)
        filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        try:
            df = pd.read_excel(filepath, engine='openpyxl')
            df.columns = [col.strip().lower().replace(' ', '_') for col in df.columns]

            imported = 0
            errors = []
            for idx, row in df.iterrows():
                row_num = idx + 2  # Excel row (1-indexed header + data)
                try:
                    item_name = str(row.get('item_name', '')).strip()
                    if not item_name or item_name == 'nan':
                        errors.append(f"Row {row_num}: Missing item_name, skipped.")
                        continue

                    quantity = int(row.get('quantity', 0)) if pd.notna(row.get('quantity')) else 0
                    unit_price = float(row.get('unit_price', 0)) if pd.notna(row.get('unit_price')) else 0.0

                    stock = Stock(
                        item_name=item_name,
                        category=str(row.get('category', '')).strip() if pd.notna(row.get('category')) else '',
                        quantity=quantity,
                        unit=str(row.get('unit', 'pcs')).strip() if pd.notna(row.get('unit')) else 'pcs',
                        unit_price=unit_price,
                        total_value=quantity * unit_price,
                        condition=str(row.get('condition', 'Good')).strip() if pd.notna(row.get('condition')) else 'Good',
                        remarks=str(row.get('remarks', '')).strip() if pd.notna(row.get('remarks')) else '',
                        campus_id=campus_id,
                        added_by=current_user.username,
                    )
                    db.session.add(stock)
                    imported += 1

                except Exception as e:
                    errors.append(f"Row {row_num}: {str(e)}")

            db.session.commit()
            flash(f'Successfully imported {imported} items to {campus.name}.', 'success')
            if errors:
                flash(f'{len(errors)} rows had issues: ' + '; '.join(errors[:5]), 'warning')

        except Exception as e:
            flash(f'Error reading Excel file: {str(e)}', 'danger')
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)

        return redirect(url_for('stock.dashboard'))

    return render_template('upload_excel.html', form=form)


@excel_bp.route('/download/template')
@login_required
def download_template():
    """Download a blank Excel template for stock data entry."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Stock Template'

    headers = ['item_name', 'category', 'quantity', 'unit', 'unit_price', 'condition', 'remarks']
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

    # Sample row
    sample = ['Laptop', 'Electronics', 10, 'pcs', 45000, 'Good', 'Dell Latitude']
    for col_idx, val in enumerate(sample, 1):
        cell = ws.cell(row=2, column=col_idx, value=val)
        cell.border = thin_border

    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col) + 4
        ws.column_dimensions[col[0].column_letter].width = max_length

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='stock_upload_template.xlsx'
    )


@excel_bp.route('/download/campus/<int:campus_id>')
@login_required
def download_campus_stock(campus_id):
    """Download stock data for a specific campus as Excel."""
    campus = db.session.get(Campus, campus_id)
    if not campus:
        flash('Campus not found.', 'danger')
        return redirect(url_for('stock.dashboard'))

    stocks = Stock.query.filter_by(campus_id=campus_id).order_by(Stock.category, Stock.item_name).all()

    wb = Workbook()
    ws = wb.active
    ws.title = f'{campus.code} Stock'

    # Title row
    ws.merge_cells('A1:H1')
    title_cell = ws.cell(row=1, column=1, value=f'Stock Report - {campus.name} ({campus.code})')
    title_cell.font = Font(bold=True, size=14, color='1F4E79')
    title_cell.alignment = Alignment(horizontal='center')

    # Header row
    headers = ['S.No', 'Item Name', 'Category', 'Quantity', 'Unit', 'Unit Price', 'Total Value', 'Condition', 'Remarks']
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

    # Data rows
    grand_total = 0
    for row_idx, stock in enumerate(stocks, 4):
        sno = row_idx - 3
        total_val = (stock.quantity or 0) * (stock.unit_price or 0)
        grand_total += total_val

        data = [sno, stock.item_name, stock.category, stock.quantity,
                stock.unit, stock.unit_price, total_val, stock.condition, stock.remarks]
        for col_idx, val in enumerate(data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            if col_idx in (6, 7):  # price columns
                cell.number_format = '#,##0.00'

    # Grand total row
    total_row = len(stocks) + 4
    ws.cell(row=total_row, column=6, value='Grand Total:').font = Font(bold=True)
    total_cell = ws.cell(row=total_row, column=7, value=grand_total)
    total_cell.font = Font(bold=True, size=12)
    total_cell.number_format = '#,##0.00'

    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col) + 4
        ws.column_dimensions[col[0].column_letter].width = min(max_length, 40)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'stock_{campus.code}.xlsx'
    )


@excel_bp.route('/download/all')
@login_required
def download_all_stock():
    """Download stock data for ALL campuses, each campus on a separate sheet."""
    campuses = Campus.query.order_by(Campus.name).all()
    if not campuses:
        flash('No campuses found.', 'warning')
        return redirect(url_for('stock.dashboard'))

    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    headers = ['S.No', 'Item Name', 'Category', 'Quantity', 'Unit', 'Unit Price', 'Total Value', 'Condition', 'Remarks']

    for campus in campuses:
        ws = wb.create_sheet(title=campus.code[:31])  # sheet name max 31 chars
        stocks = Stock.query.filter_by(campus_id=campus.id).order_by(Stock.category, Stock.item_name).all()

        # Title
        ws.merge_cells('A1:I1')
        title_cell = ws.cell(row=1, column=1, value=f'Stock Report - {campus.name} ({campus.code})')
        title_cell.font = Font(bold=True, size=14, color='1F4E79')
        title_cell.alignment = Alignment(horizontal='center')

        # Headers
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border

        # Data
        grand_total = 0
        for row_idx, stock in enumerate(stocks, 4):
            total_val = (stock.quantity or 0) * (stock.unit_price or 0)
            grand_total += total_val
            data = [row_idx - 3, stock.item_name, stock.category, stock.quantity,
                    stock.unit, stock.unit_price, total_val, stock.condition, stock.remarks]
            for col_idx, val in enumerate(data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                if col_idx in (6, 7):
                    cell.number_format = '#,##0.00'

        total_row = len(stocks) + 4
        ws.cell(row=total_row, column=6, value='Grand Total:').font = Font(bold=True)
        total_cell = ws.cell(row=total_row, column=7, value=grand_total)
        total_cell.font = Font(bold=True, size=12)
        total_cell.number_format = '#,##0.00'

        for col in ws.columns:
            max_length = max(len(str(cell.value or '')) for cell in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_length, 40)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='stock_all_campuses.xlsx'
    )
