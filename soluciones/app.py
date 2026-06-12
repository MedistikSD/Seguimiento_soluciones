from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory, abort
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import datetime, date
from functools import wraps
from models import db, User, Solicitud, Comentario, Bitacora, Documento
import os, uuid

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'soluciones-logisticas-secret-2026')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///database.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB
UPLOAD_BASE = os.path.join(os.path.dirname(__file__), 'uploads')
ALLOWED_EXTENSIONS = {'pdf','xlsx','xls','pptx','docx','png','jpg','jpeg','zip'}

db.init_app(app)

with app.app_context():
    db.create_all()

login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Inicia sesión para continuar.'
login_manager.login_message_category = 'warning'

# Roles:
#   lider_comercial  → gestiona usuarios comerciales, crea solicitudes, ve métricas
#   lider_soluciones → gestiona usuarios ingenieros, ve todas las solicitudes, cierra
#   comercial        → antes "hunter"
#   ingeniero        → antes "soluciones" / "analista"

ROLES_LABEL = {
    'lider_comercial':  'Líder Comercial',
    'lider_soluciones': 'Líder de Soluciones',
    'aux_comercial':    'Aux Comercial',
    'comercial':        'Comercial',
    'ingeniero':        'Ingeniero',
}

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


# ── Decoradores de rol ──────────────────────────────────────────────────────
def rol_requerido(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if current_user.rol not in roles:
                flash('No tienes permiso para realizar esta acción.', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return decorated
    return decorator


# ── Helpers ─────────────────────────────────────────────────────────────────
def generar_folio():
    anio = datetime.utcnow().year
    ultima = (Solicitud.query
              .filter(Solicitud.folio.like(f'SOL-{anio}-%'))
              .order_by(Solicitud.id.desc())
              .first())
    num = int(ultima.folio.split('-')[-1]) + 1 if ultima else 1
    return f'SOL-{anio}-{num:04d}'


def registrar_bitacora(solicitud_id, accion, usuario_id=None):
    b = Bitacora(
        solicitud_id=solicitud_id,
        usuario_id=usuario_id or (current_user.id if current_user.is_authenticated else None),
        accion=accion
    )
    db.session.add(b)


def es_lider():
    return current_user.rol in ('lider_comercial', 'lider_soluciones', 'aux_comercial')

def es_comercial():
    return current_user.rol in ('comercial', 'lider_comercial', 'aux_comercial')

def es_soluciones():
    return current_user.rol in ('ingeniero', 'lider_soluciones')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_upload_path(folio, tipo):
    path = os.path.join(UPLOAD_BASE, folio, tipo)
    os.makedirs(path, exist_ok=True)
    return path


# ── Auth ─────────────────────────────────────────────────────────────────────
@app.route('/', methods=['GET', 'POST'])
@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        user = User.query.filter_by(username=username, activo=True).first()
        if user and check_password_hash(user.password_hash, password):
            login_user(user, remember=True)
            return redirect(url_for('dashboard'))
        flash('Usuario o contraseña incorrectos.', 'danger')
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# ── Dashboard ────────────────────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    q = Solicitud.query

    if current_user.rol == 'lider_comercial':
        # Ve solo las solicitudes de sus comerciales
        ids = [u.id for u in User.query.filter_by(rol='comercial').all()]
        ids.append(current_user.id)
        q = q.filter(Solicitud.hunter_id.in_(ids))

    todas = q.all()
    abiertas  = [s for s in todas if s.estatus not in ('Cerrada', 'Propuesta Enviada')]
    cerradas  = [s for s in todas if s.estatus == 'Cerrada']
    vencidas  = [s for s in todas if s.dias_sin_movimiento() > 15 and s.estatus != 'Cerrada']
    monto_total = sum(s.monto_oportunidad or 0 for s in todas)

    tiempos = [s.dias_desde_captura() for s in cerradas]
    promedio_atencion = round(sum(tiempos) / len(tiempos), 1) if tiempos else 0

    por_comercial = {}
    por_ingeniero = {}
    if es_lider():
        for u in User.query.filter_by(rol='comercial').all():
            cnt = Solicitud.query.filter_by(hunter_id=u.id).filter(Solicitud.estatus != 'Cerrada').count()
            if cnt: por_comercial[u.nombre] = cnt
        for u in User.query.filter_by(rol='ingeniero').all():
            cnt = Solicitud.query.filter_by(responsable_id=u.id).filter(Solicitud.estatus != 'Cerrada').count()
            if cnt: por_ingeniero[u.nombre] = cnt

    estatus_list = ['Capturada','Asignada','En Análisis','Pendiente Información Cliente',
                    'Información Completa','Propuesta Enviada','Cerrada']
    por_estatus = {e: len([s for s in todas if s.estatus == e]) for e in estatus_list if len([s for s in todas if s.estatus == e])}

    ultimas = sorted(todas, key=lambda s: s.ultima_actualizacion or s.fecha_captura, reverse=True)[:5]

    return render_template('dashboard.html',
                           total=len(todas), abiertas=len(abiertas),
                           cerradas=len(cerradas), vencidas=len(vencidas),
                           monto_total=monto_total, promedio_atencion=promedio_atencion,
                           por_comercial=por_comercial, por_ingeniero=por_ingeniero,
                           por_estatus=por_estatus, ultimas=ultimas,
                           ROLES_LABEL=ROLES_LABEL)


# ── Solicitudes ──────────────────────────────────────────────────────────────
@app.route('/solicitudes')
@login_required
def solicitudes():
    q = Solicitud.query

    if current_user.rol == 'lider_comercial':
        ids = [u.id for u in User.query.filter_by(rol='comercial').all()]
        ids.append(current_user.id)
        q = q.filter(Solicitud.hunter_id.in_(ids))

    folio        = request.args.get('folio', '').strip()
    cliente      = request.args.get('cliente', '').strip()
    estatus      = request.args.get('estatus', '').strip()
    comercial_f  = request.args.get('comercial', '').strip()
    ingeniero_f  = request.args.get('ingeniero', '').strip()

    if folio:       q = q.filter(Solicitud.folio.ilike(f'%{folio}%'))
    if cliente:     q = q.filter(Solicitud.cliente.ilike(f'%{cliente}%'))
    if estatus:     q = q.filter(Solicitud.estatus == estatus)
    if comercial_f:
        q = q.join(User, Solicitud.hunter_id == User.id).filter(User.nombre.ilike(f'%{comercial_f}%'))
    if ingeniero_f:
        ra = db.aliased(User)
        q = q.join(ra, Solicitud.responsable_id == ra.id).filter(ra.nombre.ilike(f'%{ingeniero_f}%'))

    lista      = q.order_by(Solicitud.fecha_captura.desc()).all()
    comerciales = User.query.filter_by(rol='comercial', activo=True).all()
    ingenieros  = User.query.filter_by(rol='ingeniero', activo=True).all()
    estatus_list = ['Capturada','Asignada','En Análisis','Pendiente Información Cliente',
                    'Información Completa','Propuesta Enviada','Cerrada']
    return render_template('solicitudes.html', solicitudes=lista,
                           comerciales=comerciales, ingenieros=ingenieros,
                           estatus_list=estatus_list)


@app.route('/solicitudes/nueva', methods=['GET', 'POST'])
@login_required
@rol_requerido('comercial', 'lider_comercial', 'aux_comercial')
def nueva_solicitud():
    ingenieros  = User.query.filter_by(rol='ingeniero', activo=True).all()
    comerciales = User.query.filter(User.rol.in_(['comercial','aux_comercial']), User.activo==True).all()

    if request.method == 'POST':
        f = request.form
        try:
            fecha_sol = datetime.strptime(f['fecha_solicitud'], '%Y-%m-%d').date()
        except (ValueError, KeyError):
            flash('Fecha de solicitud inválida.', 'danger')
            return render_template('nueva_solicitud.html', ingenieros=ingenieros, comerciales=comerciales)

        responsable_id = int(f['responsable_id']) if f.get('responsable_id') else None
        sol = Solicitud(
            folio=generar_folio(),
            hunter_id=int(f.get('hunter_id', current_user.id)),
            responsable_id=responsable_id,
            fecha_solicitud=fecha_sol,
            cliente=f['cliente'].strip(),
            tema=f['tema'].strip(),
            comentarios_comerciales=f.get('comentarios_comerciales', '').strip(),
            monto_oportunidad=float(f['monto_oportunidad']) if f.get('monto_oportunidad') else None,
            prioridad=f.get('prioridad', 'Media'),
            estatus='Asignada' if responsable_id else 'Capturada',
        )
        db.session.add(sol)
        db.session.flush()
        registrar_bitacora(sol.id, f'{current_user.nombre} creó la solicitud.')
        db.session.commit()
        flash(f'Solicitud {sol.folio} creada exitosamente.', 'success')
        return redirect(url_for('detalle_solicitud', folio=sol.folio))

    return render_template('nueva_solicitud.html', ingenieros=ingenieros, comerciales=comerciales)


@app.route('/solicitudes/<folio>')
@login_required
def detalle_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()

    if current_user.rol == 'lider_comercial':
        ids = [u.id for u in User.query.filter_by(rol='comercial').all()]
        ids.append(current_user.id)
        if sol.hunter_id not in ids:
            flash('No tienes acceso a esta solicitud.', 'danger')
            return redirect(url_for('solicitudes'))

    ingenieros   = User.query.filter_by(rol='ingeniero', activo=True).all()
    estatus_list = ['Capturada','Asignada','En Análisis','Pendiente Información Cliente',
                    'Información Completa','Propuesta Enviada','Cerrada']
    return render_template('detalle_solicitud.html', sol=sol,
                           ingenieros=ingenieros, estatus_list=estatus_list)


@app.route('/solicitudes/<folio>/actualizar', methods=['POST'])
@login_required
@rol_requerido('ingeniero', 'lider_soluciones')
def actualizar_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    if current_user.rol == 'ingeniero' and sol.responsable_id != current_user.id:
        flash('No tienes permiso para actualizar esta solicitud.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio))

    f = request.form
    cambios = []
    checkboxes = {
        'historial_surtido':    'Historial de Surtido',
        'inventario':           'Inventario',
        'maestro_productos':    'Maestro de Productos',
        'historial_recepcion':  'Historial de Recepción',
        'cuestionario_logistico': 'Cuestionario Logístico',
    }
    for campo, label in checkboxes.items():
        nuevo = campo in f
        if getattr(sol, campo) != nuevo:
            setattr(sol, campo, nuevo)
            cambios.append(f'{current_user.nombre} {"marcó" if nuevo else "desmarcó"} {label}.')

    if f.get('fecha_compromiso'):
        try:
            nf = datetime.strptime(f['fecha_compromiso'], '%Y-%m-%d').date()
            if sol.fecha_compromiso != nf:
                sol.fecha_compromiso = nf
                cambios.append(f'{current_user.nombre} actualizó Fecha Compromiso a {nf.strftime("%d/%m/%Y")}.')
        except ValueError:
            pass

    if f.get('estatus'):
        nuevo_est = f['estatus']
        if sol.estatus != nuevo_est and nuevo_est != 'Cerrada':
            sol.estatus = nuevo_est
            cambios.append(f'{current_user.nombre} cambió estatus a "{nuevo_est}".')

    sol.ultima_actualizacion = datetime.utcnow()
    sol.actualizar_estatus_automatico()
    for c in cambios:
        registrar_bitacora(sol.id, c)
    db.session.commit()
    flash('Solicitud actualizada.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/envio', methods=['POST'])
@login_required
@rol_requerido('ingeniero', 'lider_soluciones')
def registrar_envio(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    if not request.form.get('comentarios_envio', '').strip():
        flash('Debes agregar un comentario del envío.', 'warning')
        return redirect(url_for('detalle_solicitud', folio=folio))
    sol.fecha_envio_cliente  = datetime.utcnow()
    sol.usuario_envio_id     = current_user.id
    sol.comentarios_envio    = request.form['comentarios_envio'].strip()
    sol.estatus              = 'Propuesta Enviada'
    sol.ultima_actualizacion = datetime.utcnow()
    registrar_bitacora(sol.id, f'{current_user.nombre} registró envío de propuesta al cliente.')
    db.session.commit()
    flash('Envío registrado exitosamente.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/cerrar', methods=['POST'])
@login_required
@rol_requerido('lider_soluciones')
def cerrar_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    sol.estatus              = 'Cerrada'
    sol.fecha_cierre         = datetime.utcnow()
    sol.usuario_cierre_id    = current_user.id
    sol.ultima_actualizacion = datetime.utcnow()
    registrar_bitacora(sol.id, f'{current_user.nombre} cerró la solicitud.')
    db.session.commit()
    flash('Solicitud cerrada.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/reasignar', methods=['POST'])
@login_required
@rol_requerido('lider_soluciones')
def reasignar_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    nuevo_id = request.form.get('responsable_id')
    if nuevo_id:
        nuevo_resp = db.session.get(User, int(nuevo_id))
        if nuevo_resp:
            sol.responsable_id = nuevo_resp.id
            if sol.estatus == 'Capturada':
                sol.estatus = 'Asignada'
            sol.ultima_actualizacion = datetime.utcnow()
            registrar_bitacora(sol.id, f'{current_user.nombre} reasignó la solicitud a {nuevo_resp.nombre}.')
            db.session.commit()
            flash(f'Solicitud reasignada a {nuevo_resp.nombre}.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/eliminar', methods=['POST'])
@login_required
@rol_requerido('lider_comercial')
def eliminar_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    folio_guardado = sol.folio
    cliente = sol.cliente
    # Eliminar registros relacionados primero
    Comentario.query.filter_by(solicitud_id=sol.id).delete()
    Bitacora.query.filter_by(solicitud_id=sol.id).delete()
    db.session.delete(sol)
    db.session.commit()
    flash(f'Solicitud {folio_guardado} ({cliente}) eliminada permanentemente.', 'success')
    return redirect(url_for('solicitudes'))


@app.route('/solicitudes/<folio>/comentario', methods=['POST'])
@login_required
def agregar_comentario(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    # Comercial e ingeniero pueden ver y comentar en cualquier solicitud
    texto = request.form.get('texto', '').strip()
    if not texto:
        flash('El comentario no puede estar vacío.', 'warning')
        return redirect(url_for('detalle_solicitud', folio=folio))
    db.session.add(Comentario(solicitud_id=sol.id, usuario_id=current_user.id, texto=texto))
    sol.ultima_actualizacion = datetime.utcnow()
    registrar_bitacora(sol.id, f'{current_user.nombre} agregó un comentario.')
    db.session.commit()
    flash('Comentario agregado.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


# ══════════════════════════════════════════════════════════════════════════════
# ── PANEL DE ADMINISTRACIÓN DE USUARIOS ──────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

@app.route('/admin/usuarios')
@login_required
@rol_requerido('lider_comercial', 'lider_soluciones')
def admin_usuarios():
    # Cada líder solo ve los roles que administra
    if current_user.rol in ('lider_comercial', 'aux_comercial'):
        roles_visibles = ['comercial', 'aux_comercial', 'lider_comercial']
    else:
        roles_visibles = ['ingeniero', 'lider_soluciones']

    usuarios = User.query.filter(User.rol.in_(roles_visibles)).order_by(User.nombre).all()
    return render_template('admin_usuarios.html',
                           usuarios=usuarios,
                           roles_visibles=roles_visibles,
                           ROLES_LABEL=ROLES_LABEL)


@app.route('/admin/usuarios/nuevo', methods=['GET', 'POST'])
@login_required
@rol_requerido('lider_comercial', 'lider_soluciones')
def admin_nuevo_usuario():
    if current_user.rol in ('lider_comercial', 'aux_comercial'):
        roles_permitidos = ['comercial', 'aux_comercial']
    else:
        roles_permitidos = ['ingeniero']

    if request.method == 'POST':
        f = request.form
        username = f.get('username', '').strip().lower()
        nombre   = f.get('nombre', '').strip()
        rol      = f.get('rol', '')
        password = f.get('password', '')
        confirm  = f.get('confirm', '')

        if not all([username, nombre, rol, password]):
            flash('Todos los campos son obligatorios.', 'danger')
        elif rol not in roles_permitidos:
            flash('Rol no permitido.', 'danger')
        elif password != confirm:
            flash('Las contraseñas no coinciden.', 'danger')
        elif len(password) < 6:
            flash('La contraseña debe tener al menos 6 caracteres.', 'danger')
        elif User.query.filter_by(username=username).first():
            flash(f'El usuario "{username}" ya existe.', 'danger')
        else:
            nuevo = User(username=username, nombre=nombre, rol=rol,
                         password_hash=generate_password_hash(password))
            db.session.add(nuevo)
            db.session.commit()
            flash(f'Usuario {nombre} creado correctamente.', 'success')
            return redirect(url_for('admin_usuarios'))

    return render_template('admin_form_usuario.html',
                           usuario=None,
                           roles_permitidos=roles_permitidos,
                           ROLES_LABEL=ROLES_LABEL,
                           accion='nuevo')


@app.route('/admin/usuarios/<int:uid>/editar', methods=['GET', 'POST'])
@login_required
@rol_requerido('lider_comercial', 'lider_soluciones')
def admin_editar_usuario(uid):
    usuario = db.session.get(User, uid)
    if not usuario:
        flash('Usuario no encontrado.', 'danger')
        return redirect(url_for('admin_usuarios'))

    # Verificar que el líder solo edite los suyos
    if current_user.rol in ('lider_comercial', 'aux_comercial'):
        roles_permitidos = ['comercial', 'aux_comercial']
    else:
        roles_permitidos = ['ingeniero']

    if usuario.rol not in roles_permitidos and usuario.id != current_user.id:
        flash('No puedes editar este usuario.', 'danger')
        return redirect(url_for('admin_usuarios'))

    if request.method == 'POST':
        f        = request.form
        nombre   = f.get('nombre', '').strip()
        username = f.get('username', '').strip().lower()
        password = f.get('password', '').strip()
        confirm  = f.get('confirm', '').strip()
        activo   = 'activo' in f

        existe = User.query.filter_by(username=username).first()

        if not nombre:
            flash('El nombre es obligatorio.', 'danger')
        elif not username:
            flash('El usuario es obligatorio.', 'danger')
        elif existe and existe.id != usuario.id:
            flash(f'El usuario "{username}" ya está en uso por otra persona.', 'danger')
        elif password and password != confirm:
            flash('Las contraseñas no coinciden.', 'danger')
        elif password and len(password) < 6:
            flash('La contraseña debe tener al menos 6 caracteres.', 'danger')
        else:
            usuario.nombre   = nombre
            usuario.username = username
            usuario.activo   = activo
            if password:
                usuario.password_hash = generate_password_hash(password)
            db.session.commit()
            flash(f'Usuario {nombre} actualizado correctamente.', 'success')
            return redirect(url_for('admin_usuarios'))

    return render_template('admin_form_usuario.html',
                           usuario=usuario,
                           roles_permitidos=roles_permitidos,
                           ROLES_LABEL=ROLES_LABEL,
                           accion='editar')


@app.route('/admin/usuarios/<int:uid>/toggle', methods=['POST'])
@login_required
@rol_requerido('lider_comercial', 'lider_soluciones')
def admin_toggle_usuario(uid):
    usuario = db.session.get(User, uid)
    if not usuario or usuario.id == current_user.id:
        flash('Operación no permitida.', 'danger')
        return redirect(url_for('admin_usuarios'))

    if current_user.rol in ('lider_comercial','aux_comercial') and usuario.rol not in ['comercial','aux_comercial']:
        flash('No puedes modificar este usuario.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if current_user.rol == 'lider_soluciones' and usuario.rol not in ['ingeniero']:
        flash('No puedes modificar este usuario.', 'danger')
        return redirect(url_for('admin_usuarios'))

    usuario.activo = not usuario.activo
    db.session.commit()
    estado = 'activado' if usuario.activo else 'desactivado'
    flash(f'Usuario {usuario.nombre} {estado}.', 'success')
    return redirect(url_for('admin_usuarios'))



# ══════════════════════════════════════════════════════════════════════════════
# ── GESTIÓN DOCUMENTAL ───────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def puede_subir_doc(tipo):
    """¿Puede el usuario actual subir/reemplazar/eliminar documentos de este tipo?"""
    if current_user.rol in ('lider_comercial', 'lider_soluciones', 'aux_comercial'):
        return True
    if tipo == 'comercial' and current_user.rol == 'comercial':
        return True
    if tipo == 'soluciones' and current_user.rol == 'ingeniero':
        return True
    return False

def puede_descargar_doc(tipo):
    """Todos los roles pueden descargar."""
    return True


@app.route('/solicitudes/<folio>/documentos/subir', methods=['POST'])
@login_required
def subir_documento(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    tipo = request.form.get('tipo_documento', '')

    if tipo not in ('comercial', 'soluciones'):
        flash('Tipo de documento inválido.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    if not puede_subir_doc(tipo):
        flash('No tienes permiso para subir documentos en esta sección.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    archivo = request.files.get('archivo')
    if not archivo or archivo.filename == '':
        flash('No seleccionaste ningún archivo.', 'warning')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    if not allowed_file(archivo.filename):
        flash('Formato no permitido. Usa: PDF, XLSX, XLS, PPTX, DOCX, PNG, JPG, ZIP.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    nombre_original = secure_filename(archivo.filename)
    ext = nombre_original.rsplit('.', 1)[1].lower()
    nombre_guardado = f"{uuid.uuid4().hex}.{ext}"

    ruta = get_upload_path(folio, tipo)
    archivo.save(os.path.join(ruta, nombre_guardado))

    doc = Documento(
        solicitud_id=sol.id,
        nombre_original=nombre_original,
        nombre_guardado=nombre_guardado,
        tipo_documento=tipo,
        usuario_id=current_user.id,
        version=1,
        activo=True,
    )
    db.session.add(doc)
    db.session.flush()
    registrar_bitacora(sol.id, f'{current_user.nombre} cargó archivo "{nombre_original}" ({tipo}).')
    db.session.commit()
    flash(f'Archivo "{nombre_original}" subido correctamente.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')


@app.route('/solicitudes/<folio>/documentos/<int:doc_id>/reemplazar', methods=['POST'])
@login_required
def reemplazar_documento(folio, doc_id):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    doc_anterior = db.session.get(Documento, doc_id)
    if not doc_anterior or not doc_anterior.activo:
        flash('Documento no encontrado.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    tipo = doc_anterior.tipo_documento
    if not puede_subir_doc(tipo):
        flash('No tienes permiso para reemplazar este documento.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    archivo = request.files.get('archivo')
    if not archivo or archivo.filename == '':
        flash('No seleccionaste ningún archivo.', 'warning')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    if not allowed_file(archivo.filename):
        flash('Formato no permitido.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    # Marcar versión anterior como inactiva (histórico)
    doc_anterior.activo = False

    nombre_original = secure_filename(archivo.filename)
    ext = nombre_original.rsplit('.', 1)[1].lower()
    nombre_guardado = f"{uuid.uuid4().hex}.{ext}"

    ruta = get_upload_path(folio, tipo)
    archivo.save(os.path.join(ruta, nombre_guardado))

    nueva_version = Documento(
        solicitud_id=sol.id,
        nombre_original=nombre_original,
        nombre_guardado=nombre_guardado,
        tipo_documento=tipo,
        usuario_id=current_user.id,
        version=doc_anterior.version + 1,
        activo=True,
        documento_padre_id=doc_anterior.id,
    )
    db.session.add(nueva_version)
    registrar_bitacora(sol.id,
        f'{current_user.nombre} reemplazó "{doc_anterior.nombre_original}" '
        f'→ "{nombre_original}" (v{nueva_version.version}, {tipo}).')
    db.session.commit()
    flash(f'Archivo reemplazado. Nueva versión {nueva_version.version} guardada.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')


@app.route('/solicitudes/<folio>/documentos/<int:doc_id>/eliminar', methods=['POST'])
@login_required
def eliminar_documento(folio, doc_id):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    doc = db.session.get(Documento, doc_id)
    if not doc or not doc.activo:
        flash('Documento no encontrado.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    if not puede_subir_doc(doc.tipo_documento):
        flash('No tienes permiso para eliminar este documento.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    # Eliminar archivo físico
    ruta = get_upload_path(folio, doc.tipo_documento)
    ruta_archivo = os.path.join(ruta, doc.nombre_guardado)
    if os.path.exists(ruta_archivo):
        os.remove(ruta_archivo)

    nombre = doc.nombre_original
    tipo   = doc.tipo_documento
    db.session.delete(doc)
    registrar_bitacora(sol.id, f'{current_user.nombre} eliminó archivo "{nombre}" ({tipo}).')
    db.session.commit()
    flash(f'Archivo "{nombre}" eliminado.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')


@app.route('/solicitudes/<folio>/documentos/<int:doc_id>/descargar')
@login_required
def descargar_documento(folio, doc_id):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    doc = db.session.get(Documento, doc_id)
    if not doc:
        abort(404)

    ruta = get_upload_path(folio, doc.tipo_documento)
    if not os.path.exists(os.path.join(ruta, doc.nombre_guardado)):
        flash('Archivo no encontrado en el servidor.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')

    registrar_bitacora(sol.id,
        f'{current_user.nombre} descargó "{doc.nombre_original}" ({doc.tipo_documento}).')
    db.session.commit()
    return send_from_directory(ruta, doc.nombre_guardado,
                               as_attachment=True,
                               download_name=doc.nombre_original)

# ── Inicialización de DB ──────────────────────────────────────────────────────
def init_db():
    db.create_all()

    if User.query.count() == 0:
        users = [
            User(username='Francisco_Cueva',     password_hash=generate_password_hash('Lcomercial123'),  nombre='Francisco Cueva',     rol='lider_comercial'),
            User(username='Andrés_Toledo',        password_hash=generate_password_hash('Lsoluciones123'), nombre='Andrés Toledo',        rol='lider_soluciones'),
            User(username='Gerardo_Velazquez',    password_hash=generate_password_hash('IngeSD1231'),     nombre='Gerardo Velazquez',    rol='ingeniero'),
            User(username='Elizabeth_Bastida',    password_hash=generate_password_hash('IngeSD1232'),     nombre='Elizabeth Bastida',    rol='ingeniero'),
            User(username='Diego_Arzate',         password_hash=generate_password_hash('IgeSDT123'),      nombre='Diego Arzate',         rol='ingeniero'),
            User(username='Jorge_Camarena',       password_hash=generate_password_hash('IngeSD1233'),     nombre='Jorge Camarena',       rol='ingeniero'),
            User(username='Teresa_Ruiz',          password_hash=generate_password_hash('Hunter1231'),     nombre='Teresa Ruiz',          rol='comercial'),
            User(username='Alejandra_Sanchez',    password_hash=generate_password_hash('Hunter1232'),     nombre='Alejandra Sánchez',    rol='comercial'),
            User(username='Diana_Pelcastre',      password_hash=generate_password_hash('Hunter1233'),     nombre='Diana Pelcastre',      rol='comercial'),
            User(username='Ida_Acosta',           password_hash=generate_password_hash('Hunter1234'),     nombre='Ida Acosta',           rol='comercial'),
            User(username='Malena_Baltazar',      password_hash=generate_password_hash('Hunter1235'),     nombre='Malena Baltazar',      rol='comercial'),
            User(username='José_Ortega',          password_hash=generate_password_hash('Hunter1236'),     nombre='José Ortega',          rol='comercial'),
        ]
        for u in users:
            db.session.add(u)
        db.session.flush()

        demos = [
            dict(hunter_id=7, responsable_id=3, fecha_solicitud=date(2026, 5, 10),
                 cliente='FEMSA Logística', tema='Análisis de rutas de distribución',
                 comentarios_comerciales='Cliente interesado en optimizar última milla.',
                 monto_oportunidad=850000, prioridad='Alta', estatus='En Análisis',
                 historial_surtido=True, inventario=True),
            dict(hunter_id=8, responsable_id=4, fecha_solicitud=date(2026, 5, 20),
                 cliente='Grupo Bimbo', tema='Propuesta de almacenamiento refrigerado',
                 comentarios_comerciales='Necesitan solución inmediata para Q3.',
                 monto_oportunidad=1200000, prioridad='Alta', estatus='Información Completa',
                 historial_surtido=True, inventario=True, maestro_productos=True,
                 historial_recepcion=True, cuestionario_logistico=True),
            dict(hunter_id=9, responsable_id=5, fecha_solicitud=date(2026, 6, 1),
                 cliente='Liverpool', tema='Gestión de devoluciones e-commerce',
                 monto_oportunidad=500000, prioridad='Media', estatus='Capturada'),
            dict(hunter_id=10, responsable_id=None, fecha_solicitud=date(2026, 6, 5),
                 cliente='Soriana', tema='Estudio de factibilidad CD regional',
                 monto_oportunidad=3000000, prioridad='Alta', estatus='Capturada'),
            dict(hunter_id=11, responsable_id=6, fecha_solicitud=date(2026, 4, 15),
                 cliente='Amazon MX', tema='Fulfillment centers integración',
                 monto_oportunidad=4500000, prioridad='Alta', estatus='Propuesta Enviada',
                 historial_surtido=True, inventario=True, maestro_productos=True,
                 historial_recepcion=True, cuestionario_logistico=True,
                 fecha_envio_cliente=datetime(2026, 5, 28), usuario_envio_id=6,
                 comentarios_envio='Se envió propuesta económica completa al cliente vía correo.'),
        ]
        for d in demos:
            s = Solicitud(folio=generar_folio(), **d)
            db.session.add(s)
            db.session.flush()
            hunter = db.session.get(User, s.hunter_id)
            registrar_bitacora(s.id, f'{hunter.nombre} creó la solicitud.', s.hunter_id)

        db.session.commit()
        print('✅ Base de datos inicializada.')
        print()
        print('  Usuario               Password         Rol')
        print('  Francisco_Cueva       Lcomercial123    Líder Comercial')
        print('  Andrés_Toledo         Lsoluciones123   Líder de Soluciones')
        print('  Gerardo_Velazquez     IngeSD1231       Ingeniero')
        print('  Elizabeth_Bastida     IngeSD1232       Ingeniero')
        print('  Diego_Arzate          IgeSDT123        Ingeniero')
        print('  Jorge_Camarena        IngeSD1233       Ingeniero')
        print('  Teresa_Ruiz           Hunter1231       Comercial')
        print('  Alejandra_Sanchez     Hunter1232       Comercial')
        print('  Diana_Pelcastre       Hunter1233       Comercial')
        print('  Ida_Acosta            Hunter1234       Comercial')
        print('  Malena_Baltazar       Hunter1235       Comercial')
        print('  José_Ortega           Hunter1236       Comercial')


if __name__ == '__main__':
    with app.app_context():
        init_db()
    port  = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
