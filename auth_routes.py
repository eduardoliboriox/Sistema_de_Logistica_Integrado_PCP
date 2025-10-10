from flask import Blueprint, render_template, redirect, url_for, flash, request, current_app
from flask_login import login_user, logout_user, login_required
from models import db, User
import unicodedata, re

auth_bp = Blueprint('auth', __name__, url_prefix='/auth')

def make_username(full_name, setor):
    name = unicodedata.normalize('NFKD', full_name).encode('ascii','ignore').decode('ascii')
    parts = [p for p in re.split(r'\s+', name.strip()) if p]
    first = parts[0].lower()
    last = parts[-1].lower() if len(parts) > 1 else ''
    return f"{first}.{last}.{setor}".lower()

def enviar_email_cadastro(destino, username):
    mail = current_app.extensions.get('mail')
    if not mail:
        print("Flask-Mail não encontrado")
        return
    try:
        mail.send()
    except Exception as e:
        print(f"Erro enviando email: {e}")

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username'].lower()
        password = request.form['password']
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)
            flash('Login realizado com sucesso!', 'success')
            return redirect('/')  # mantém na mesma aba
        else:
            flash('Usuário ou senha incorretos.', 'danger')
    return render_template('login.html')

@auth_bp.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Sessão encerrada.', 'info')
    return redirect(url_for('auth.login'))

@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        nome = request.form['nome']
        email = request.form['email']
        setor = request.form['setor']
        senha = request.form['senha']

        username = make_username(nome, setor)

        if User.query.filter_by(username=username).first():
            flash('Usuário já existente.', 'danger')
            return redirect(url_for('auth.register'))

        user = User(full_name=nome, email=email, role=setor, username=username)
        user.set_password(senha)
        db.session.add(user)
        db.session.commit()

        enviar_email_cadastro(email, username)
        return render_template('register_success.html', username=username)
    return render_template('register.html')
