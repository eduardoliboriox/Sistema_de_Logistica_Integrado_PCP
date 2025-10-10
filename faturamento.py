from flask import Blueprint, jsonify
from flask_login import login_required, current_user
from models import db, Item, ItemHistory

faturamento_bp = Blueprint('faturamento', __name__, url_prefix='/faturamento')

@faturamento_bp.route('/marcar_faturado/<int:item_id>', methods=['POST'])
@login_required
def marcar_faturado(item_id):
    if current_user.role not in ['faturamento', 'admin']:
        return jsonify({'erro': 'Acesso negado'}), 403

    item = Item.query.get_or_404(item_id)
    antigo = item.status
    item.status = 'Faturado'
    db.session.commit()

    hist = ItemHistory(item_id=item.id, from_status=antigo, to_status='Faturado',
                       by_user_id=current_user.id, comment='Marcado como Faturado')
    db.session.add(hist)
    db.session.commit()

    return jsonify({'ok': True, 'novo_status': item.status})
