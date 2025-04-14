from flask import Flask, request, send_file, render_template
from docxtpl import DocxTemplate
from fpdf import FPDF
from io import BytesIO
import datetime
import random
from num2words import num2words
from werkzeug.utils import secure_filename
import os
import json

app = Flask(__name__)

# ---------- CONFIG ----------
ID_FILE = "quote_ids.json"

def generate_id():
    """Read last ID from JSON, increment, write back, return Qxxxxx"""
    last_id = 0
    if os.path.exists(ID_FILE):
        with open(ID_FILE, "r") as f:
            try:
                data = json.load(f)
                last_id = data.get("last_id", 0)
            except json.JSONDecodeError:
                last_id = 0

    new_id = last_id + 1

    with open(ID_FILE, "w") as f:
        json.dump({"last_id": new_id}, f)

    return f"Q{new_id:05d}"


def month_year():
    today = datetime.date.today()
    return today.strftime("%b").upper(), str(today.year)

def number_to_words(num):
    try:
        return num2words(num, to='cardinal', lang='en_IN').upper() + " RUPEES ONLY"
    except:
        return ""

@app.route('/')
def form():
    return render_template("form.html")

@app.route('/enquiry')
def enquiry():
    return render_template("enquiry.html")

@app.route('/generate-doc', methods=['POST'])
def generate_doc():
    form_data = request.get_json()

    form_data['id'] = generate_id()
    form_data['m'], form_data['y'] = month_year()
    form_data['date'] = datetime.date.today().strftime("%d/%m/%Y")

    form_data['city'] = form_data.get('client_city', '')
    form_data['qty'] = form_data.get('quantity', '')
    form_data['amt'] = form_data.get('amount', '')
    form_data['ip'] = form_data.get('installation_period', '')

    try:
        form_data['weight'] = int(form_data.get('capacity', 0)) * 68
    except:
        form_data['weight'] = ''

    try:
        rate = int(form_data.get('rate', 0))
        qty = int(form_data.get('quantity', 0))
        amt = rate * qty
        form_data['amt'] = amt
        form_data['amount'] = amt
        form_data['amount_in_words'] = number_to_words(amt)
    except:
        form_data['amount'] = ''
        form_data['amount_in_words'] = ''

    doc = DocxTemplate("template.docx")
    doc.render(form_data)

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    filename = secure_filename(f"{form_data.get('client_name', 'Client')}_{form_data.get('site_name', 'Site')}.docx")

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


# ========== Enhanced PDF Proposal Generator ==========
class ProposalPDF(FPDF):
    def header(self):
        self.set_fill_color(50, 50, 150)
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", "B", 16)
        self.cell(0, 12, "Lift Project Proposal", 0, 1, "C", fill=True)
        self.ln(5)

    def add_section(self, title):
        self.set_fill_color(230, 230, 250)
        self.set_text_color(0, 0, 0)
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, f"{title}", 0, 1, 'L', fill=True)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(2)

    def add_field(self, label, value):
        self.set_font("Arial", "", 11)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 8, f"{label}: {value}")
        self.ln(1)

@app.route('/generate-pdf', methods=['POST'])
def generate_pdf():
    data = request.get_json()
    pdf = ProposalPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.add_section("Marketing person")
    pdf.add_field("Marketing/person" , data.get("enquiry_person",""))

    pdf.add_section("Client / Site Info")
    pdf.add_field("Builder/Construction Name", data.get("builder_name", ""))
    pdf.add_field("Site Name", data.get("site_name", ""))
    pdf.add_field("Location", data.get("location", ""))
    pdf.add_field("City", data.get("city", ""))

    pdf.add_section("Decision Maker")
    pdf.add_field("Name", data.get("dm_name", ""))
    pdf.add_field("Contact", data.get("dm_contact", ""))
    pdf.add_field("Email", data.get("dm_email", ""))

    pdf.add_section("Project Details")
    pdf.add_field("Number of Floors", data.get("floors", ""))
    num_lifts = int(data.get("num_lifts", 0))
    pdf.add_field("Number of Lifts", num_lifts)

    for i in range(1, num_lifts + 1):
        pdf.add_section(f"Lift #{i}")
        pdf.add_field("Shaft Size", data.get(f"shaft_size_{i}", ""))
        pdf.add_field("Lift Type", data.get(f"lift_type_{i}", ""))

    output = BytesIO()
    output.write(pdf.output(dest='S').encode('latin1'))
    output.seek(0)

    filename = secure_filename(f"{data.get('builder_name', 'Proposal')}_{data.get('site_name', '')}.pdf")

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/pdf'
    )

if __name__ == '__main__':
    app.run(debug=True)
