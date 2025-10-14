# importer_xlwings.py
import xlwings as xw
import pandas as pd
from datetime import datetime, date
import os
from flask import current_app
from models import db, Cliente, ItemStatus

# Caminho da planilha
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
PLANILHA_CAMINHO = os.path.join(BASE_DIR, 'pcp-venttos-manaus.xlsm')
SHEET_NAME = "Plan-VenttosLogistica"

# -----------------------------
# Funções auxiliares
# -----------------------------
def clean_quantity_value(x):
    """Converte strings/valores de quantidade para float ou int."""
    if pd.isna(x):
        return 0
    if isinstance(x, str):
        s = x.strip().replace(".", "").replace(",", ".")
        try:
            return int(float(s))
        except Exception:
            return 0
    try:
        return int(x)
    except Exception:
        return 0

def try_parse_date_column(s):
    """Converte a coluna para datetime.date corrigindo inversão dia/mês."""
    if pd.isna(s):
        return date.today()
    
    if isinstance(s, datetime):
        return s.date()
    
    if isinstance(s, (int, float)):  # número serial do Excel
        try:
            return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(s) - 2).date()
        except Exception:
            return date.today()
    
    # Se for string
    try:
        # Força formato dd/mm/yyyy
        return datetime.strptime(str(s).strip(), "%d/%m/%Y").date()
    except Exception:
        try:
            # Tenta formato yyyy-mm-dd
            return datetime.strptime(str(s).strip(), "%Y-%m-%d").date()
        except Exception:
            return date.today()

def clean_header_name(h):
    """Limpa nomes de coluna."""
    if not h:
        return None
    return str(h).strip().replace("\u00a0", " ")

# -----------------------------
# Função principal
# -----------------------------
def importar_planilha_xlwings(path=PLANILHA_CAMINHO, sheet_name=SHEET_NAME):
    resumo = {'created': 0, 'updated': 0, 'errors': []}

    if not os.path.exists(path):
        msg = f"Arquivo não encontrado: {path}"
        resumo['errors'].append(msg)
        current_app.logger.warning(msg)
        return resumo

    app = xw.App(visible=False)
    wb = None

    try:
        wb = app.books.open(path)
        sht = wb.sheets[sheet_name]

        # Detectar última linha pela coluna B (Data)
        last_row = sht.range("B" + str(sht.cells.last_cell.row)).end("up").row

        # Cabeçalho B6:H6
        header_range = sht.range("B6:H6").value
        if not header_range:
            raise ValueError("Cabeçalho não encontrado em B6:H6")

        # Dados B7:H<last_row>
        data_range = sht.range(f"B7:H{last_row}").value
        if not data_range:
            raise ValueError(f"Nenhum dado encontrado entre B7:H{last_row}")

        # Normalizar linhas
        max_cols = len(header_range)
        normalized = []
        for row in data_range:
            if row is None:
                row = [None] * max_cols
            else:
                row = list(row) + [None] * (max_cols - len(row))
            normalized.append(row)

        # DataFrame
        header_clean = [clean_header_name(h) if clean_header_name(h) else f"col_{i+1}" 
                        for i, h in enumerate(header_range)]
        df = pd.DataFrame(normalized, columns=header_clean)

        # Limpar strings e propagar valores vazios
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        df = df.ffill(axis=0)

        # Detectar colunas
        col_lower = [c.lower() for c in df.columns]
        col_data = next((c for i,c in enumerate(df.columns) if 'data' in col_lower[i]), None)
        col_cliente = next((c for i,c in enumerate(df.columns) if 'cliente' in col_lower[i]), None)
        col_modelo = next((c for i,c in enumerate(df.columns) if 'modelo' in col_lower[i]), None)
        col_quant = next((c for i,c in enumerate(df.columns) if 'quant' in col_lower[i]), None)
        col_pronto = next((c for i,c in enumerate(df.columns) if 'pronto' in col_lower[i] or 'ok' in col_lower[i]), None)

        if not all([col_data, col_cliente, col_modelo, col_quant]):
            msg = "Colunas obrigatórias não encontradas (Data, Cliente, Modelo, Quantidade)"
            resumo['errors'].append(msg)
            current_app.logger.error(msg)
            return resumo

        # Iterar linhas
        for idx, row in df.iterrows():
            try:
                data_item = try_parse_date_column(row[col_data])
                cliente_nome = str(row[col_cliente]).strip()
                modelo = str(row[col_modelo]).strip()
                quantidade = clean_quantity_value(row[col_quant])
                pronto_flag = False
                if col_pronto:
                    val = row[col_pronto]
                    pronto_flag = str(val).strip().lower() in ['sim', 's', 'yes', 'true', '1', 'ok', 'pronto']

                if not cliente_nome or not modelo:
                    resumo['errors'].append(f"Linha {idx+2}: cliente ou modelo vazio. Pulando.")
                    continue

                cliente = Cliente.query.filter_by(nome=cliente_nome).first()
                if not cliente:
                    cliente = Cliente(nome=cliente_nome)
                    db.session.add(cliente)
                    db.session.flush()

                item_status = ItemStatus.query.filter_by(modelo=modelo, data=data_item).first()
                now = datetime.now()
                status_text = 'Pronto' if pronto_flag else 'Recebido'

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
                msg = f"Linha {idx+2}: Erro ao processar - {str(e)}"
                resumo['errors'].append(msg)
                current_app.logger.error(msg)

        db.session.commit()

    except Exception as e:
        resumo['errors'].append(f"Erro geral: {str(e)}")
        current_app.logger.exception("Erro ao importar planilha")
    finally:
        if wb:
            wb.close()
        app.quit()

    return resumo

# -----------------------------
# Exemplo de uso
# -----------------------------
if __name__ == "__main__":
    resultado = importar_planilha_xlwings()
    print("Resumo da importação:", resultado)
