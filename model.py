# model.py
from wtforms import FileField, FloatField, SubmitField, validators
from flask_wtf import Form
class PriceForm(Form):
  survey_responses = FileField('Upload Qualtrics .CSV data export', [validators.DataRequired()])
  produce = FloatField('Produce Share',  [validators.DataRequired()])
  meat = FloatField('Meat Share',  [validators.DataRequired()])
  egg = FloatField('Egg Share',  [validators.DataRequired()])
  bread = FloatField('Bread Share',  [validators.DataRequired()])
  dairy = FloatField('Dairy Share',  [validators.DataRequired()])
  cheese = FloatField('Cheese Share',  [validators.DataRequired()])
  preserves = FloatField('Preserves Share',  [validators.DataRequired()])
  coffee = FloatField('Coffee Share',  [validators.DataRequired()])
  tofu = FloatField('Tofu Share',  [validators.DataRequired()])
  seitan = FloatField('Seitan Share',  [validators.DataRequired()])
  mushroom = FloatField('Mushroom Share',  [validators.DataRequired()])

  submit = SubmitField('Submit')