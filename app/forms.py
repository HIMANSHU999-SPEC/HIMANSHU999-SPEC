from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed
from wtforms import (
    StringField, PasswordField, SubmitField, IntegerField,
    FloatField, SelectField, TextAreaField
)
from wtforms.validators import DataRequired, Length, EqualTo, Optional, NumberRange


class LoginForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired(), Length(min=3, max=80)])
    password = PasswordField('Password', validators=[DataRequired()])
    submit = SubmitField('Login')


class RegisterForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired(), Length(min=3, max=80)])
    password = PasswordField('Password', validators=[DataRequired(), Length(min=6)])
    confirm_password = PasswordField('Confirm Password', validators=[
        DataRequired(), EqualTo('password', message='Passwords must match')
    ])
    role = SelectField('Role', choices=[('staff', 'Staff'), ('admin', 'Admin')],
                       validators=[DataRequired()])
    submit = SubmitField('Register')


class CampusForm(FlaskForm):
    name = StringField('Campus Name', validators=[DataRequired(), Length(max=120)])
    code = StringField('Campus Code', validators=[DataRequired(), Length(max=20)])
    address = StringField('Address', validators=[Optional(), Length(max=300)])
    submit = SubmitField('Save Campus')


class StockForm(FlaskForm):
    item_name = StringField('Item Name', validators=[DataRequired(), Length(max=200)])
    category = StringField('Category', validators=[Optional(), Length(max=100)])
    quantity = IntegerField('Quantity', validators=[DataRequired(), NumberRange(min=0)])
    unit = StringField('Unit (pcs/kg/litre)', validators=[Optional(), Length(max=50)])
    unit_price = FloatField('Unit Price', validators=[Optional(), NumberRange(min=0)])
    condition = SelectField('Condition', choices=[
        ('Good', 'Good'), ('Damaged', 'Damaged'), ('Needs Repair', 'Needs Repair')
    ], validators=[DataRequired()])
    low_stock_threshold = IntegerField('Low Stock Threshold', validators=[Optional(), NumberRange(min=0)], default=10)
    campus_id = SelectField('Campus', coerce=int, validators=[DataRequired()])
    remarks = TextAreaField('Remarks', validators=[Optional(), Length(max=500)])
    submit = SubmitField('Save Stock')


class StockTransferForm(FlaskForm):
    stock_id = SelectField('Stock Item', coerce=int, validators=[DataRequired()])
    to_campus_id = SelectField('Transfer To Campus', coerce=int, validators=[DataRequired()])
    quantity = IntegerField('Quantity to Transfer', validators=[DataRequired(), NumberRange(min=1)])
    remarks = TextAreaField('Remarks', validators=[Optional(), Length(max=500)])
    submit = SubmitField('Transfer Stock')


class UploadExcelForm(FlaskForm):
    campus_id = SelectField('Target Campus', coerce=int, validators=[DataRequired()])
    file = FileField('Excel File', validators=[
        DataRequired(),
        FileAllowed(['xlsx', 'xls'], 'Only Excel files (.xlsx, .xls) are allowed!')
    ])
    submit = SubmitField('Upload & Import')
