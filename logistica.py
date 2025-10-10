from flask import Blueprint, jsonify
from flask_login import login_required, current_user
from models import db, Item, ItemHistory

logistica_bp = Blueprint('logistica', __name__, url_prefix='/logistica')

@logistica_bp.route('/atualizar_status/<int:item_id>/<novo_status>', methods=['POST'])
@login_required
def atualizar_status(item_id, novo_status):
    if current_user.role not in ['logistica', 'admin']:
        return jsonify({'erro': 'Acesso negado'}), 403

    item = Item.query.get_or_404(item_id)
    antigo = item.status
    item.status = novo_status
    db.session.commit()

    hist = ItemHistory(item_id=item.id, from_status=antigo,
                       to_status=novo_status, by_user_id=current_user.id)
    db.session.add(hist)
    db.session.commit()

    return jsonify({'ok': True, 'novo_status': item.status})
