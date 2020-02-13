import io

import pandas as pd
from flask import Flask, render_template, request, send_file
from flask_bootstrap import Bootstrap

from generate_sheet import generate_totals_sheet
from model import PriceForm

app = Flask(__name__)
app.config.from_mapping(
    SECRET_KEY=b'\xd6\x04\xbdj\xfe\xed$c\x1e@\xad\x0f\x13,@G')

Bootstrap(app)

@app.route('/', methods=['GET', 'POST'])
def compute_totals():
    form = PriceForm(request.form)

    if request.method == 'POST':

        f_string= request.files['survey_responses'].read()
        resp_df = pd.read_csv(io.BytesIO(f_string), encoding='utf8')

        resp_df = resp_df.drop(resp_df.columns[list(range(12))], axis=1).iloc[1:]

        shares = [form.produce.data, form.meat.data, form.egg.data, form.bread.data, 
                  form.dairy.data, form.cheese.data, form.preserves.data, form.coffee.data, 
                  form.tofu.data, form.seitan.data, form.mushroom.data]

        generate_totals_sheet(shares, resp_df)
        return send_file("coop_totals.xlsx", as_attachment=True)

    return render_template('form.html', form=form)

if __name__ == '__main__':
    app.run()
