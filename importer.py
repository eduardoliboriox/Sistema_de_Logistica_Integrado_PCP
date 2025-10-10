# importer.py
import pandas as pd
from datetime import datetime, date
import os
from flask import current_app
from models import db, Cliente, ItemStatus

PLANILHA_CAMINHO = r"Q:\EDUARDO LIBORIO\Programação (f)\Venttos Logistica - Arquivos\pcp-venttos-manaus.xlsm"
PLANILHA_ABA = "Plan-VenttosLogistica"

def safe_parse_int_quantity(val):
    if pd.isna(val):
        return 0
    try:
        return int(val)
    except:
        try:
            return int(float(str(val).replace(',', '.')))
        except:
            return 0

def parse_bool_pronto(val):
    if pd.isna(val):
        return False
    return str(val).strip().lower() in ['sim', 's', 'yes', 'true', '1', 'ok', 'pronto']

def importar_planilha(path=PLANILHA_CAMINHO, sheet_name=PLANILHA_ABA):
    resumo = {'created': 0, 'updated': 0, 'errors': []}
    if not os.path.exists(path):
        resumo['errors'].append(f"Arquivo não encontrado: {path}")
        current_app.logger.warning(resumo['errors'][-1])
        return resumo

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        resumo['errors'].append(f"Erro ao ler Excel: {e}")
        current_app.logger.exception("Erro lendo Excel")
        return resumo

    if df.empty:
        resumo['errors'].append("Planilha vazia.")
        return resumo

    # mapeamento de colunas
    col_map = {c.strip().lower(): c for c in df.columns}
    def find_col(candidates):
        for c in candidates:
            k = c.strip().lower()
            if k in col_map:
                return col_map[k]
        return None

    col_data = find_col(['Data', 'data'])
    col_cliente = find_col(['Cliente', 'cliente', 'o cliente', 'cliente nome'])
    col_modelo = find_col(['Modelo', 'modelo'])
    col_quant = find_col(['Quantidade', 'quantidade', 'qtd', 'qtd.'])
    col_pronto = find_col(['Pronto', 'pronto', 'ok'])

    if not (col_data and col_cliente and col_modelo and col_quant):
        resumo['errors'].append("Colunas obrigatórias não encontradas (Data, Cliente, Modelo, Quantidade).")
        current_app.logger.error(resumo['errors'][-1])
        return resumo

    for idx, row in df.iterrows():
        try:
            raw_data = row[col_data]
            data_item = pd.to_datetime(raw_data, errors='coerce').date() if not pd.isna(raw_data) else date.today()
            cliente_nome = str(row[col_cliente]).strip()
            modelo = str(row[col_modelo]).strip()
            quantidade = safe_parse_int_quantity(row[col_quant])
            pronto_flag = parse_bool_pronto(row[col_pronto]) if col_pronto else False

            if not cliente_nome or not modelo:
                resumo['errors'].append(f"Linha {idx+2}: cliente ou modelo vazio. Pulando.")
                continue

            cliente = Cliente.query.filter_by(nome=cliente_nome).first()
            if not cliente:
                cliente = Cliente(nome=cliente_nome)
                db.session.add(cliente)
                db.session.flush()

            status_text = 'Pronto' if pronto_flag else 'Recebido'
            now = datetime.now()

            item_status = ItemStatus.query.filter_by(modelo=modelo, data=data_item).first()
            if item_status:
                item_status.quantidade = quantidade
                item_status.status = status_text
                item_status.usuario_ultimo_update = 'importacao_automatica'
                item_status.hora_ultimo_update = now
                item_status.cliente = cliente.nome
                resumo['updated'] += 1
            else:
                item_status = ItemStatus(
                    cliente=cliente.nome,
                    modelo=modelo,
                    quantidade=quantidade,
                    status=status_text,
                    usuario_ultimo_update='importacao_automatica',
                    hora_ultimo_update=now,
                    data=data_item
                )
                db.session.add(item_status)
                resumo['created'] += 1
        except Exception as e:
            resumo['errors'].append(f"Linha {idx+2}: erro - {e}")
            current_app.logger.exception(f"Erro importando linha {idx+2}")

    try:
        db.session.commit()
    except Exception as e:
        resumo['errors'].append(f"Erro no commit do DB: {e}")
        current_app.logger.exception("Erro no commit final")
    return resumo
