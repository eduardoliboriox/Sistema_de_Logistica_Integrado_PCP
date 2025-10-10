from flask import Blueprint, render_template, request, redirect, url_for, flash, current_app
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
import pandas as pd
import os
from models import db, Cliente, Item, PCPUpload, ItemHistory

pcp_bp = Blueprint('pcp', __name__, url_prefix='/pcp')
ALLOWED_EXT = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.',1)[1].lower() in ALLOWED_EXT

@pcp_bp.route('/upload', methods=['GET', 'POST'])
@login_required
def upload_excel():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not allowed_file(file.filename):
            flash('Envie um arquivo Excel válido (.xls ou .xlsx)', 'danger')
            return redirect(url_for('pcp.upload_excel'))

        filename = secure_filename(file.filename)
        path = os.path.join('./uploads', filename)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        file.save(path)

        df = pd.read_excel(path)
        columns = list(df.columns)
        preview = df.head(10).to_dict(orient='records')

        upload = PCPUpload(filename=filename, uploaded_by=current_user.id)
        db.session.add(upload)
        db.session.commit()

        return render_template('upload_preview.html',
                               columns=columns,
                               preview_rows=preview,
                               upload_id=upload.id)
    return render_template('upload_form.html')

@pcp_bp.route('/confirm_import', methods=['POST'])
@login_required
def confirm_import():
    upload_id = request.form.get('upload_id')
    cliente_col = request.form.get('map_cliente')
    modelo_col = request.form.get('map_modelo')
    qtd_col = request.form.get('map_quantidade')
    pronto_col = request.form.get('map_pronto')

    upload = PCPUpload.query.get(upload_id)
    path = os.path.join('./uploads', upload.filename)
    df = pd.read_excel(path)

    created = 0
    for _, row in df.iterrows():
        cliente_nome = row.get(cliente_col, 'Cliente não informado')
        modelo = str(row.get(modelo_col, 'N/A'))
        quantidade = int(row.get(qtd_col, 0))
        pronto_flag = bool(row.get(pronto_col)) if pronto_col else False

        cliente = Cliente.query.filter_by(nome=cliente_nome).first()
        if not cliente:
            cliente = Cliente(nome=cliente_nome)
            db.session.add(cliente)
            db.session.commit()

        item = Item(cliente_id=cliente.id, modelo=modelo, quantidade=quantidade,
                    origem_upload_id=upload.id, status='Pronto' if pronto_flag else 'Recebido',
                    criado_por=current_user.id)
        db.session.add(item)
        db.session.commit()

        hist = ItemHistory(item_id=item.id, from_status=None,
                           to_status=item.status, by_user_id=current_user.id)
        db.session.add(hist)
        db.session.commit()
        created += 1

    flash(f'{created} itens importados com sucesso!', 'success')
    return redirect(url_for('index'))
