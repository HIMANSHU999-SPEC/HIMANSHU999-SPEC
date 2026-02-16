# Stock Check System

A Flask-based web application for managing stock/inventory across multiple campuses with Excel import/export support.

## Features

- **Login System** - Role-based authentication (Admin / Staff)
- **Campus Management** - Add, edit, delete campuses (admin only)
- **Stock Management** - Full CRUD for stock items with search and category filtering
- **Excel Import** - Upload stock data from `.xlsx` files using a provided template
- **Campus-wise Excel Download** - Download formatted stock reports per campus
- **Download All** - Export all campuses into a single Excel file with separate sheets
- **Dashboard** - Overview of all campuses with item counts and total values

## Project Structure

```
├── run.py                  # Application entry point
├── requirements.txt        # Python dependencies
├── .gitignore
├── app/
│   ├── __init__.py         # App factory, DB init, blueprints
│   ├── models.py           # User, Campus, Stock models
│   ├── forms.py            # WTForms for all inputs
│   ├── routes/
│   │   ├── auth.py         # Login, register, logout
│   │   ├── stock.py        # Dashboard, campus CRUD, stock CRUD
│   │   └── excel.py        # Upload, download template, download per campus/all
│   ├── templates/
│   │   ├── base.html       # Layout with navbar
│   │   ├── login.html
│   │   ├── register.html
│   │   ├── dashboard.html
│   │   ├── campus_form.html
│   │   ├── campus_stocks.html
│   │   ├── stock_form.html
│   │   └── upload_excel.html
│   └── static/css/
│       └── style.css
├── instance/               # SQLite database (auto-created)
└── uploads/                # Temp upload folder (auto-cleaned)
```

## Setup & Run

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python run.py
```

The app will be available at `http://localhost:5000`.

## Default Login

| Username | Password  | Role  |
|----------|-----------|-------|
| admin    | admin123  | Admin |

Change the default password after first login.

## Excel Template

The upload expects these columns in the Excel file:

| Column      | Required | Description                            |
|-------------|----------|----------------------------------------|
| item_name   | Yes      | Name of the stock item                 |
| quantity    | Yes      | Number of items                        |
| category    | No       | Category (e.g. Electronics, Furniture) |
| unit        | No       | Unit of measurement (pcs, kg, litre)   |
| unit_price  | No       | Price per unit                         |
| condition   | No       | Good / Damaged / Needs Repair          |
| remarks     | No       | Additional notes                       |

Download the template from the app: **Excel > Download Template**.
