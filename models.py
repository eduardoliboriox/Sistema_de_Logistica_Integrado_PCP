from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

# ==========================
# Usuários
# ==========================
class User(UserMixin, db.Model):
    __tablename__ = 'user'
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(120), nullable=False)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True)
    password_hash = db.Column(db.String(200))
    role = db.Column(db.String(50))  # pcp, faturamento, logistica, admin

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

# ==========================
# Status dos itens (dashboard)
# ==========================
class ItemStatus(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cliente = db.Column(db.String(200), nullable=False)
    modelo = db.Column(db.String(100), nullable=False)
    quantidade = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(50), nullable=False, default='Apontamento PCP')
    usuario_ultimo_update = db.Column(db.String(100))
    hora_ultimo_update = db.Column(db.DateTime)
    data = db.Column(db.Date, nullable=False, default=datetime.today)

    def to_dict(self):
        return {
            'cliente': self.cliente,
            'modelo': self.modelo,
            'quantidade': self.quantidade,
            'status': self.status,
            'usuario': self.usuario_ultimo_update,
            'hora': self.hora_ultimo_update.strftime('%d/%m/%Y %H:%M') if self.hora_ultimo_update else '',
            'data': self.data.strftime('%Y-%m-%d')
        }

# ==========================
# Clientes
# ==========================
class Cliente(db.Model):
    __tablename__ = 'cliente'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

# ==========================
# Uploads de PCP
# ==========================
class PCPUpload(db.Model):
    __tablename__ = 'pcp_upload'
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200))
    uploaded_by = db.Column(db.Integer, db.ForeignKey('user.id'))
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

# ==========================
# Itens
# ==========================
class Item(db.Model):
    __tablename__ = 'item'
    id = db.Column(db.Integer, primary_key=True)
    cliente_id = db.Column(db.Integer, db.ForeignKey('cliente.id'))
    modelo = db.Column(db.String(120))
    quantidade = db.Column(db.Integer)
    status = db.Column(db.String(50))
    origem_upload_id = db.Column(db.Integer, db.ForeignKey('pcp_upload.id'))
    criado_por = db.Column(db.Integer, db.ForeignKey('user.id'))
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

# ==========================
# Histórico de alterações dos itens
# ==========================
class ItemHistory(db.Model):
    __tablename__ = 'item_history'
    id = db.Column(db.Integer, primary_key=True)
    item_id = db.Column(db.Integer, db.ForeignKey('item.id'))
    from_status = db.Column(db.String(50))
    to_status = db.Column(db.String(50))
    by_user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    comment = db.Column(db.String(250))
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
