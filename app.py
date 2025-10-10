# app.py
from flask import Flask, render_template, request, jsonify
from flask_login import LoginManager, login_required, current_user
from flask_mail import Mail
from config import Config
from models import db, User, ItemStatus
from auth_routes import auth_bp
from pcp import pcp_bp
from datetime import datetime, date
from importer import importar_planilha

try:
    from apscheduler.schedulers.background import BackgroundScheduler
    APSCHEDULER_AVAILABLE = True
except:
    APSCHEDULER_AVAILABLE = False

mail = Mail()

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    db.init_app(app)
    mail.init_app(app)

    login_manager = LoginManager(app)
    login_manager.login_view = 'auth.login'
    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    app.register_blueprint(auth_bp)
    app.register_blueprint(pcp_bp)

    @app.route('/')
    @login_required
    def index():
        selected_date_str = request.args.get('data')
        selected_date = datetime.strptime(selected_date_str, '%Y-%m-%d').date() if selected_date_str else date.today()
        itens = ItemStatus.query.filter_by(data=selected_date).all()
        grouped = {}
        for item in itens:
            grouped.setdefault(item.cliente, []).append(item.to_dict())
        return render_template('dashboard.html', grouped=grouped, selected_date=selected_date, user=current_user)

    @app.route('/update_status', methods=['POST'])
    @login_required
    def update_status():
        data = request.get_json()
        modelo = data.get('modelo')
        novo_status = data.get('status')
        senha = data.get('senha')
        selected_date_str = data.get('data')
        selected_date = datetime.strptime(selected_date_str, '%Y-%m-%d').date() if selected_date_str else date.today()

        if not current_user.check_password(senha):
            return jsonify({'success': False, 'msg': 'Senha incorreta'})

        item = ItemStatus.query.filter_by(modelo=modelo, data=selected_date).first()
        if not item:
            return jsonify({'success': False, 'msg': 'Item não encontrado'})

        item.status = novo_status
        item.usuario_ultimo_update = current_user.username
        item.hora_ultimo_update = datetime.now()
        db.session.commit()

        return jsonify({
            'success': True,
            'hora': item.hora_ultimo_update.strftime('%d/%m/%Y %H:%M'),
            'usuario': item.usuario_ultimo_update
        })

    @app.route('/pcp/import_now', methods=['POST'])
    @login_required
    def import_now_route():
        if current_user.role not in ['pcp', 'admin']:
            return jsonify({'ok': False, 'msg': 'Acesso negado'}), 403
        resumo = importar_planilha()
        return jsonify({'ok': True, 'resumo': resumo})

    # Scheduler automático
    def start_scheduler():
        if not APSCHEDULER_AVAILABLE:
            app.logger.info("APScheduler não disponível.")
            return
        scheduler = BackgroundScheduler()
        def job_wrapper():
            with app.app_context():
                resumo = importar_planilha()
                app.logger.info(f"Importação automática executada: {resumo}")
        scheduler.add_job(job_wrapper, 'interval', minutes=5, id='import_pcp_every_5m', replace_existing=True)
        scheduler.start()
        app.logger.info("Scheduler iniciado: import_pcp_every_5m")

    with app.app_context():
        try:
            app.logger.info("Importação inicial da planilha...")
            resumo = importar_planilha()
            app.logger.info(f"Resumo inicial: {resumo}")
        except:
            app.logger.exception("Erro na importação inicial")
        start_scheduler()

    return app

if __name__ == '__main__':
    app = create_app()
    with app.app_context():
        db.create_all()
    app.run(host="0.0.0.0", port=5000, debug=True)
