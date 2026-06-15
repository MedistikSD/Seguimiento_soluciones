from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory, abort
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import datetime, date
from functools import wraps
from models import db, User, Solicitud, Comentario, Bitacora, Documento, TemasSolicitud, cdmx_now
import os, uuid, csv, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from flask import Response

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'soluciones-logisticas-secret-2026')
# Render usa 'postgres://' pero SQLAlchemy requiere 'postgresql://'
_db_url = os.environ.get('DATABASE_URL', 'sqlite:///database.db')
if _db_url.startswith('postgres://'):
    _db_url = _db_url.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = _db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
UPLOAD_BASE = os.path.join(os.path.dirname(__file__), 'uploads')
ALLOWED_EXTENSIONS = {'pdf','xlsx','xls','pptx','docx','png','jpg','jpeg','zip'}

db.init_app(app)

with app.app_context():
    db.create_all()

login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Inicia sesión para continuar.'
login_manager.login_message_category = 'warning'

ROLES_LABEL = {
    'administrador':    'Administrador',
    'lider_comercial':  'Líder Comercial',
    'lider_soluciones': 'Líder de Soluciones',
    'aux_comercial':    'Aux Comercial',
    'comercial':        'Comercial',
    'ingeniero':        'Ingeniero',
}

TEMAS_DEFAULT = ['Reingeniería','Transporte','Almacenaje','VAS','Transporte y almacenaje']

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


# ── Decoradores ──────────────────────────────────────────────────────────────
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


# ── Helpers ──────────────────────────────────────────────────────────────────
def es_admin():
    return current_user.rol == 'administrador'

def es_lider():
    return current_user.rol in ('administrador','lider_comercial','lider_soluciones','aux_comercial')

def es_comercial():
    return current_user.rol in ('administrador','comercial','lider_comercial','aux_comercial')

def es_soluciones():
    return current_user.rol in ('administrador','ingeniero','lider_soluciones')

def puede_asignar_ingeniero():
    return current_user.rol in ('administrador','lider_soluciones')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_upload_path(folio, tipo):
    path = os.path.join(UPLOAD_BASE, folio, tipo)
    os.makedirs(path, exist_ok=True)
    return path

def generar_folio():
    anio = cdmx_now().year
    ultima = (Solicitud.query
              .filter(Solicitud.folio.like(f'SOL-{anio}-%'))
              .order_by(Solicitud.id.desc()).first())
    num = int(ultima.folio.split('-')[-1]) + 1 if ultima else 1
    return f'SOL-{anio}-{num:04d}'

def registrar_bitacora(solicitud_id, accion, usuario_id=None):
    db.session.add(Bitacora(
        solicitud_id=solicitud_id,
        usuario_id=usuario_id or (current_user.id if current_user.is_authenticated else None),
        accion=accion
    ))

def get_temas():
    return TemasSolicitud.query.filter_by(activo=True).order_by(TemasSolicitud.orden, TemasSolicitud.nombre).all()


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


# ── Recuperar contraseña ─────────────────────────────────────────────────────
@app.route('/recuperar', methods=['GET', 'POST'])
def recuperar_password():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        nueva    = request.form.get('nueva_password', '')
        confirm  = request.form.get('confirm_password', '')
        user = User.query.filter_by(username=username, activo=True).first()
        if not user:
            flash('Usuario no encontrado.', 'danger')
        elif len(nueva) < 6:
            flash('La contraseña debe tener al menos 6 caracteres.', 'danger')
        elif nueva != confirm:
            flash('Las contraseñas no coinciden.', 'danger')
        else:
            user.password_hash = generate_password_hash(nueva)
            db.session.commit()
            flash('Contraseña actualizada. Ya puedes iniciar sesión.', 'success')
            return redirect(url_for('login'))
    return render_template('recuperar_password.html')


# ── Dashboard ────────────────────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    q = Solicitud.query

    # administrador y lider_soluciones ven TODAS sin restricción
    if current_user.rol in ('lider_comercial', 'aux_comercial'):
        ids = [u.id for u in User.query.filter(User.rol.in_(['comercial','aux_comercial'])).all()]
        ids.append(current_user.id)
        q = q.filter(Solicitud.hunter_id.in_(ids))
    # lider_soluciones e ingeniero tienen su propio scope en lógica de charts

    # Filtros dashboard
    f_estatus  = request.args.get('f_estatus', '').strip()
    f_comercial= request.args.get('f_comercial', '').strip()
    f_tema     = request.args.get('f_tema', '').strip()

    if f_estatus:   q = q.filter(Solicitud.estatus == f_estatus)
    if f_comercial: q = q.join(User, Solicitud.hunter_id == User.id).filter(User.id == int(f_comercial))
    if f_tema:      q = q.filter(Solicitud.tema == f_tema)

    todas = q.all()
    abiertas  = [s for s in todas if s.estatus not in ('Cerrada','Propuesta Enviada')]
    cerradas  = [s for s in todas if s.estatus == 'Cerrada']
    vencidas  = [s for s in todas if s.dias_sin_movimiento() > 15 and s.estatus != 'Cerrada']
    monto_total = sum(s.monto_oportunidad or 0 for s in todas)

    tiempos = [s.dias_desde_captura() for s in cerradas]
    promedio_atencion = round(sum(tiempos)/len(tiempos), 1) if tiempos else 0

    # Gráficas: calculadas desde `todas` para respetar todos los filtros activos
    por_comercial_monto = {}
    por_ingeniero_monto = {}
    if es_lider():
        from collections import defaultdict
        com_montos = defaultdict(float)
        ing_montos = defaultdict(float)
        for s in todas:
            if s.estatus != 'Cerrada':
                monto = s.monto_oportunidad or 0
                if s.hunter:
                    com_montos[s.hunter.nombre] += monto
                if s.responsable and s.responsable.rol == 'ingeniero':
                    ing_montos[s.responsable.nombre] += monto
        por_comercial_monto = {k: round(v) for k, v in
            sorted(com_montos.items(), key=lambda x: x[1], reverse=True) if v}
        por_ingeniero_monto = {k: round(v) for k, v in
            sorted(ing_montos.items(), key=lambda x: x[1], reverse=True) if v}

    estatus_list = ['Capturada','Asignada','En Análisis','Pendiente Información Cliente',
                    'Información Completa','Propuesta Enviada','Cerrada']
    por_estatus = {e: len([s for s in todas if s.estatus == e])
                   for e in estatus_list if any(s.estatus == e for s in todas)}

    ultimas = sorted(todas, key=lambda s: s.ultima_actualizacion or s.fecha_captura, reverse=True)[:5]
    comerciales_list = User.query.filter(User.rol.in_(['comercial','aux_comercial','lider_comercial'])).all()
    temas_list = get_temas()

    return render_template('dashboard.html',
        total=len(todas), abiertas=len(abiertas), cerradas=len(cerradas),
        vencidas=len(vencidas), monto_total=monto_total,
        promedio_atencion=promedio_atencion,
        por_comercial_monto=por_comercial_monto,
        por_ingeniero_monto=por_ingeniero_monto,
        por_estatus=por_estatus, ultimas=ultimas,
        ROLES_LABEL=ROLES_LABEL,
        estatus_list=estatus_list,
        comerciales_list=comerciales_list,
        temas_list=temas_list,
        f_estatus=f_estatus, f_comercial=f_comercial, f_tema=f_tema)


# ── Solicitudes ──────────────────────────────────────────────────────────────
@app.route('/solicitudes')
@login_required
def solicitudes():
    q = Solicitud.query
    # administrador y lider_soluciones ven TODAS sin restricción
    if current_user.rol in ('lider_comercial','aux_comercial'):
        ids = [u.id for u in User.query.filter(User.rol.in_(['comercial','aux_comercial'])).all()]
        ids.append(current_user.id)
        q = q.filter(Solicitud.hunter_id.in_(ids))
    elif current_user.rol == 'ingeniero':
        q = q.filter(Solicitud.responsable_id == current_user.id)

    folio       = request.args.get('folio','').strip()
    cliente     = request.args.get('cliente','').strip()
    estatus     = request.args.get('estatus','').strip()
    comercial_f = request.args.get('comercial','').strip()
    ingeniero_f = request.args.get('ingeniero','').strip()
    tema_f      = request.args.get('tema','').strip()

    if folio:       q = q.filter(Solicitud.folio.ilike(f'%{folio}%'))
    if cliente:     q = q.filter(Solicitud.cliente.ilike(f'%{cliente}%'))
    if estatus:     q = q.filter(Solicitud.estatus == estatus)
    if tema_f:      q = q.filter(Solicitud.tema == tema_f)
    if comercial_f:
        q = q.join(User, Solicitud.hunter_id == User.id).filter(User.nombre.ilike(f'%{comercial_f}%'))
    if ingeniero_f:
        ra = db.aliased(User)
        q = q.join(ra, Solicitud.responsable_id == ra.id).filter(ra.nombre.ilike(f'%{ingeniero_f}%'))

    lista      = q.order_by(Solicitud.fecha_captura.desc()).all()
    ingenieros = User.query.filter_by(rol='ingeniero', activo=True).all()
    estatus_list = ['Capturada','Asignada','En Análisis','Pendiente Información Cliente',
                    'Información Completa','Propuesta Enviada','Cerrada']
    return render_template('solicitudes.html', solicitudes=lista,
                           ingenieros=ingenieros, estatus_list=estatus_list,
                           temas_list=get_temas())


@app.route('/solicitudes/nueva', methods=['GET', 'POST'])
@login_required
@rol_requerido('comercial','lider_comercial','aux_comercial','administrador')
def nueva_solicitud():
    ingenieros  = User.query.filter_by(rol='ingeniero', activo=True).all()
    comerciales = User.query.filter(User.rol.in_(['comercial','aux_comercial','lider_comercial']), User.activo==True).all()
    temas       = get_temas()

    if request.method == 'POST':
        f = request.form
        try:
            fecha_sol = datetime.strptime(f['fecha_solicitud'], '%Y-%m-%d').date()
        except (ValueError, KeyError):
            flash('Fecha de solicitud inválida.', 'danger')
            return render_template('nueva_solicitud.html', ingenieros=ingenieros,
                                   comerciales=comerciales, temas=temas)

        # Fecha no puede ser anterior a hoy
        if fecha_sol < date.today():
            flash('La fecha de solicitud no puede ser anterior a hoy.', 'danger')
            return render_template('nueva_solicitud.html', ingenieros=ingenieros,
                                   comerciales=comerciales, temas=temas)

        # El ingeniero siempre se asigna después desde el detalle (solo lider_soluciones)
        responsable_id = None

        sol = Solicitud(
            folio=generar_folio(),
            hunter_id=int(f.get('hunter_id', current_user.id)),
            responsable_id=responsable_id,
            fecha_solicitud=fecha_sol,
            cliente=f['cliente'].strip(),
            tema=f['tema'].strip(),
            comentarios_comerciales=f.get('comentarios_comerciales','').strip(),
            monto_oportunidad=float(f['monto_oportunidad']) if f.get('monto_oportunidad') else None,
            prioridad=None,  # Asignada por líder después
            estatus='Asignada' if responsable_id else 'Capturada',
        )
        db.session.add(sol)
        db.session.flush()
        registrar_bitacora(sol.id, f'{current_user.nombre} creó la solicitud.')

        # Archivos adjuntos al crear (comercial)
        archivos = request.files.getlist('archivos_iniciales')
        for archivo in archivos:
            if archivo and archivo.filename and allowed_file(archivo.filename):
                nombre_original = secure_filename(archivo.filename)
                ext = nombre_original.rsplit('.', 1)[1].lower()
                nombre_guardado = f"{uuid.uuid4().hex}.{ext}"
                ruta = get_upload_path(sol.folio, 'comercial')
                archivo.save(os.path.join(ruta, nombre_guardado))
                doc = Documento(
                    solicitud_id=sol.id, nombre_original=nombre_original,
                    nombre_guardado=nombre_guardado, tipo_documento='comercial',
                    usuario_id=current_user.id, version=1, activo=True,
                )
                db.session.add(doc)
                registrar_bitacora(sol.id, f'{current_user.nombre} adjuntó "{nombre_original}" al crear la solicitud.')

        db.session.commit()
        flash(f'Solicitud {sol.folio} creada exitosamente.', 'success')
        return redirect(url_for('detalle_solicitud', folio=sol.folio))

    return render_template('nueva_solicitud.html', ingenieros=ingenieros,
                           comerciales=comerciales, temas=temas)


@app.route('/solicitudes/<folio>')
@login_required
def detalle_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    # administrador y lider_soluciones acceden a cualquier solicitud sin filtro
    if current_user.rol in ('lider_comercial', 'aux_comercial'):
        ids = [u.id for u in User.query.filter(User.rol.in_(['comercial','aux_comercial'])).all()]
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
@rol_requerido('ingeniero','lider_soluciones','administrador')
def actualizar_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    if current_user.rol == 'ingeniero' and sol.responsable_id != current_user.id and not es_admin():
        flash('No tienes permiso para actualizar esta solicitud.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio))

    f = request.form
    cambios = []
    checkboxes = {
        'historial_surtido':'Historial de Surtido','inventario':'Inventario',
        'maestro_productos':'Maestro de Productos','historial_recepcion':'Historial de Recepción',
        'cuestionario_logistico':'Cuestionario Logístico',
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

    sol.ultima_actualizacion = cdmx_now()
    sol.actualizar_estatus_automatico()
    for c in cambios:
        registrar_bitacora(sol.id, c)
    db.session.commit()
    flash('Solicitud actualizada.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/envio', methods=['POST'])
@login_required
@rol_requerido('ingeniero','lider_soluciones','administrador')
def registrar_envio(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    if not request.form.get('comentarios_envio','').strip():
        flash('Debes agregar un comentario del envío.', 'warning')
        return redirect(url_for('detalle_solicitud', folio=folio))
    sol.fecha_envio_cliente  = cdmx_now()
    sol.usuario_envio_id     = current_user.id
    sol.comentarios_envio    = request.form['comentarios_envio'].strip()
    sol.estatus              = 'Propuesta Enviada'
    sol.ultima_actualizacion = cdmx_now()
    registrar_bitacora(sol.id, f'{current_user.nombre} registró envío de propuesta al cliente.')
    db.session.commit()
    flash('Envío registrado exitosamente.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/cerrar', methods=['POST'])
@login_required
@rol_requerido('lider_soluciones','administrador')
def cerrar_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    sol.estatus = 'Cerrada'; sol.fecha_cierre = cdmx_now()
    sol.usuario_cierre_id = current_user.id; sol.ultima_actualizacion = cdmx_now()
    registrar_bitacora(sol.id, f'{current_user.nombre} cerró la solicitud.')
    db.session.commit()
    flash('Solicitud cerrada.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/reasignar', methods=['POST'])
@login_required
@rol_requerido('lider_soluciones','administrador')
def reasignar_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    nuevo_id = request.form.get('responsable_id')
    if nuevo_id:
        nuevo_resp = db.session.get(User, int(nuevo_id))
        if nuevo_resp:
            sol.responsable_id = nuevo_resp.id
            if sol.estatus == 'Capturada': sol.estatus = 'Asignada'
            sol.ultima_actualizacion = cdmx_now()
            registrar_bitacora(sol.id, f'{current_user.nombre} reasignó a {nuevo_resp.nombre}.')
            db.session.commit()
            flash(f'Reasignada a {nuevo_resp.nombre}.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/comentario', methods=['POST'])
@login_required
def agregar_comentario(folio):
    sol  = Solicitud.query.filter_by(folio=folio).first_or_404()
    texto = request.form.get('texto','').strip()
    if not texto:
        flash('El comentario no puede estar vacío.', 'warning')
        return redirect(url_for('detalle_solicitud', folio=folio))
    db.session.add(Comentario(solicitud_id=sol.id, usuario_id=current_user.id, texto=texto))
    sol.ultima_actualizacion = cdmx_now()
    registrar_bitacora(sol.id, f'{current_user.nombre} agregó un comentario.')
    db.session.commit()
    flash('Comentario agregado.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))


@app.route('/solicitudes/<folio>/eliminar', methods=['POST'])
@login_required
@rol_requerido('lider_comercial','administrador')
def eliminar_solicitud(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    folio_g = sol.folio; cliente = sol.cliente
    Comentario.query.filter_by(solicitud_id=sol.id).delete()
    Bitacora.query.filter_by(solicitud_id=sol.id).delete()
    Documento.query.filter_by(solicitud_id=sol.id).delete()
    db.session.delete(sol)
    db.session.commit()
    flash(f'Solicitud {folio_g} ({cliente}) eliminada.', 'success')
    return redirect(url_for('solicitudes'))


# ── Admin: Temas ─────────────────────────────────────────────────────────────
@app.route('/admin/temas')
@login_required
@rol_requerido('administrador')
def admin_temas():
    temas = TemasSolicitud.query.order_by(TemasSolicitud.orden, TemasSolicitud.nombre).all()
    return render_template('admin_temas.html', temas=temas)


@app.route('/admin/temas/nuevo', methods=['POST'])
@login_required
@rol_requerido('administrador')
def admin_nuevo_tema():
    nombre = request.form.get('nombre','').strip()
    if not nombre:
        flash('El nombre es obligatorio.', 'danger')
    elif TemasSolicitud.query.filter_by(nombre=nombre).first():
        flash(f'"{nombre}" ya existe.', 'danger')
    else:
        db.session.add(TemasSolicitud(nombre=nombre))
        db.session.commit()
        flash(f'Tema "{nombre}" agregado.', 'success')
    return redirect(url_for('admin_temas'))


@app.route('/admin/temas/<int:tid>/toggle', methods=['POST'])
@login_required
@rol_requerido('administrador')
def admin_toggle_tema(tid):
    tema = db.session.get(TemasSolicitud, tid)
    if tema:
        tema.activo = not tema.activo
        db.session.commit()
        flash(f'Tema "{tema.nombre}" {"activado" if tema.activo else "desactivado"}.', 'success')
    return redirect(url_for('admin_temas'))


@app.route('/admin/temas/<int:tid>/eliminar', methods=['POST'])
@login_required
@rol_requerido('administrador')
def admin_eliminar_tema(tid):
    tema = db.session.get(TemasSolicitud, tid)
    if tema:
        db.session.delete(tema)
        db.session.commit()
        flash('Tema eliminado.', 'success')
    return redirect(url_for('admin_temas'))


# ── Admin: Usuarios ───────────────────────────────────────────────────────────
@app.route('/admin/usuarios')
@login_required
@rol_requerido('lider_comercial','lider_soluciones','aux_comercial','administrador')
def admin_usuarios():
    if es_admin():
        roles_visibles = list(ROLES_LABEL.keys())
    elif current_user.rol in ('lider_comercial','aux_comercial'):
        roles_visibles = ['comercial','aux_comercial','lider_comercial']
    else:
        roles_visibles = ['ingeniero','lider_soluciones']
    usuarios = User.query.filter(User.rol.in_(roles_visibles)).order_by(User.nombre).all()
    return render_template('admin_usuarios.html', usuarios=usuarios,
                           roles_visibles=roles_visibles, ROLES_LABEL=ROLES_LABEL)


@app.route('/admin/usuarios/nuevo', methods=['GET','POST'])
@login_required
@rol_requerido('lider_comercial','lider_soluciones','aux_comercial','administrador')
def admin_nuevo_usuario():
    if es_admin():
        roles_permitidos = list(ROLES_LABEL.keys())
    elif current_user.rol in ('lider_comercial','aux_comercial'):
        roles_permitidos = ['comercial','aux_comercial']
    else:
        roles_permitidos = ['ingeniero']

    if request.method == 'POST':
        f = request.form
        username = f.get('username','').strip().lower()
        nombre   = f.get('nombre','').strip()
        rol      = f.get('rol','')
        password = f.get('password','')
        confirm  = f.get('confirm','')

        if not all([username, nombre, rol, password]):
            flash('Todos los campos son obligatorios.', 'danger')
        elif rol not in roles_permitidos:
            flash('Rol no permitido.', 'danger')
        elif password != confirm:
            flash('Las contraseñas no coinciden.', 'danger')
        elif len(password) < 6:
            flash('Mínimo 6 caracteres.', 'danger')
        elif User.query.filter_by(username=username).first():
            flash(f'El usuario "{username}" ya existe.', 'danger')
        else:
            db.session.add(User(username=username, nombre=nombre, rol=rol,
                                password_hash=generate_password_hash(password)))
            db.session.commit()
            flash(f'Usuario {nombre} creado.', 'success')
            return redirect(url_for('admin_usuarios'))

    return render_template('admin_form_usuario.html', usuario=None,
                           roles_permitidos=roles_permitidos,
                           ROLES_LABEL=ROLES_LABEL, accion='nuevo')


@app.route('/admin/usuarios/<int:uid>/editar', methods=['GET','POST'])
@login_required
@rol_requerido('lider_comercial','lider_soluciones','aux_comercial','administrador')
def admin_editar_usuario(uid):
    usuario = db.session.get(User, uid)
    if not usuario:
        flash('Usuario no encontrado.', 'danger')
        return redirect(url_for('admin_usuarios'))

    if es_admin():
        roles_permitidos = list(ROLES_LABEL.keys())
    elif current_user.rol in ('lider_comercial','aux_comercial'):
        roles_permitidos = ['comercial','aux_comercial']
    else:
        roles_permitidos = ['ingeniero']

    if not es_admin() and usuario.rol not in roles_permitidos and usuario.id != current_user.id:
        flash('No puedes editar este usuario.', 'danger')
        return redirect(url_for('admin_usuarios'))

    if request.method == 'POST':
        f        = request.form
        nombre   = f.get('nombre','').strip()
        username = f.get('username','').strip().lower()
        password = f.get('password','').strip()
        confirm  = f.get('confirm','').strip()
        activo   = 'activo' in f
        existe   = User.query.filter_by(username=username).first()

        if not nombre:
            flash('El nombre es obligatorio.', 'danger')
        elif not username:
            flash('El usuario es obligatorio.', 'danger')
        elif existe and existe.id != usuario.id:
            flash(f'El usuario "{username}" ya está en uso.', 'danger')
        elif password and password != confirm:
            flash('Las contraseñas no coinciden.', 'danger')
        elif password and len(password) < 6:
            flash('Mínimo 6 caracteres.', 'danger')
        else:
            usuario.nombre = nombre; usuario.username = username; usuario.activo = activo
            if password:
                usuario.password_hash = generate_password_hash(password)
            db.session.commit()
            flash(f'Usuario {nombre} actualizado.', 'success')
            return redirect(url_for('admin_usuarios'))

    return render_template('admin_form_usuario.html', usuario=usuario,
                           roles_permitidos=roles_permitidos,
                           ROLES_LABEL=ROLES_LABEL, accion='editar')


@app.route('/admin/usuarios/<int:uid>/toggle', methods=['POST'])
@login_required
@rol_requerido('lider_comercial','lider_soluciones','aux_comercial','administrador')
def admin_toggle_usuario(uid):
    usuario = db.session.get(User, uid)
    if not usuario or usuario.id == current_user.id:
        flash('Operación no permitida.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if not es_admin():
        if current_user.rol in ('lider_comercial','aux_comercial') and usuario.rol not in ['comercial','aux_comercial']:
            flash('No puedes modificar este usuario.', 'danger')
            return redirect(url_for('admin_usuarios'))
        if current_user.rol == 'lider_soluciones' and usuario.rol not in ['ingeniero']:
            flash('No puedes modificar este usuario.', 'danger')
            return redirect(url_for('admin_usuarios'))
    usuario.activo = not usuario.activo
    db.session.commit()
    flash(f'Usuario {usuario.nombre} {"activado" if usuario.activo else "desactivado"}.', 'success')
    return redirect(url_for('admin_usuarios'))


# ── Documentos ────────────────────────────────────────────────────────────────
def puede_subir_doc(tipo):
    if current_user.rol in ('administrador','lider_comercial','lider_soluciones','aux_comercial'):
        return True
    if tipo == 'comercial' and current_user.rol == 'comercial':
        return True
    if tipo == 'soluciones' and current_user.rol == 'ingeniero':
        return True
    return False


@app.route('/solicitudes/<folio>/documentos/subir', methods=['POST'])
@login_required
def subir_documento(folio):
    sol  = Solicitud.query.filter_by(folio=folio).first_or_404()
    tipo = request.form.get('tipo_documento','')
    if tipo not in ('comercial','soluciones'):
        flash('Tipo inválido.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    if not puede_subir_doc(tipo):
        flash('Sin permiso.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    archivo = request.files.get('archivo')
    if not archivo or archivo.filename == '':
        flash('No seleccionaste archivo.', 'warning')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    if not allowed_file(archivo.filename):
        flash('Formato no permitido.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    nombre_original = secure_filename(archivo.filename)
    ext = nombre_original.rsplit('.', 1)[1].lower()
    nombre_guardado = f"{uuid.uuid4().hex}.{ext}"
    archivo.save(os.path.join(get_upload_path(folio, tipo), nombre_guardado))
    db.session.add(Documento(solicitud_id=sol.id, nombre_original=nombre_original,
        nombre_guardado=nombre_guardado, tipo_documento=tipo,
        usuario_id=current_user.id, version=1, activo=True))
    registrar_bitacora(sol.id, f'{current_user.nombre} cargó "{nombre_original}" ({tipo}).')
    db.session.commit()
    flash(f'"{nombre_original}" subido.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')


@app.route('/solicitudes/<folio>/documentos/<int:doc_id>/reemplazar', methods=['POST'])
@login_required
def reemplazar_documento(folio, doc_id):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    doc_ant = db.session.get(Documento, doc_id)
    if not doc_ant or not doc_ant.activo:
        flash('Documento no encontrado.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    if not puede_subir_doc(doc_ant.tipo_documento):
        flash('Sin permiso.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    archivo = request.files.get('archivo')
    if not archivo or archivo.filename == '' or not allowed_file(archivo.filename):
        flash('Archivo inválido.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    doc_ant.activo = False
    nombre_original = secure_filename(archivo.filename)
    ext = nombre_original.rsplit('.', 1)[1].lower()
    nombre_guardado = f"{uuid.uuid4().hex}.{ext}"
    archivo.save(os.path.join(get_upload_path(folio, doc_ant.tipo_documento), nombre_guardado))
    nueva = Documento(solicitud_id=sol.id, nombre_original=nombre_original,
        nombre_guardado=nombre_guardado, tipo_documento=doc_ant.tipo_documento,
        usuario_id=current_user.id, version=doc_ant.version+1,
        activo=True, documento_padre_id=doc_ant.id)
    db.session.add(nueva)
    registrar_bitacora(sol.id, f'{current_user.nombre} reemplazó "{doc_ant.nombre_original}" → "{nombre_original}" (v{nueva.version}).')
    db.session.commit()
    flash(f'Reemplazado. Nueva versión {nueva.version}.', 'success')
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
        flash('Sin permiso.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    ruta_arch = os.path.join(get_upload_path(folio, doc.tipo_documento), doc.nombre_guardado)
    if os.path.exists(ruta_arch): os.remove(ruta_arch)
    nombre = doc.nombre_original; tipo = doc.tipo_documento
    db.session.delete(doc)
    registrar_bitacora(sol.id, f'{current_user.nombre} eliminó "{nombre}" ({tipo}).')
    db.session.commit()
    flash(f'"{nombre}" eliminado.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')


@app.route('/solicitudes/<folio>/documentos/<int:doc_id>/descargar')
@login_required
def descargar_documento(folio, doc_id):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    doc = db.session.get(Documento, doc_id)
    if not doc: abort(404)
    ruta = get_upload_path(folio, doc.tipo_documento)
    if not os.path.exists(os.path.join(ruta, doc.nombre_guardado)):
        flash('Archivo no encontrado.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio) + '#archivos')
    registrar_bitacora(sol.id, f'{current_user.nombre} descargó "{doc.nombre_original}" ({doc.tipo_documento}).')
    db.session.commit()
    return send_from_directory(ruta, doc.nombre_guardado,
                               as_attachment=True, download_name=doc.nombre_original)





@app.route('/solicitudes/<folio>/prioridad', methods=['POST'])
@login_required
@rol_requerido('lider_comercial','lider_soluciones','administrador')
def actualizar_prioridad(folio):
    sol = Solicitud.query.filter_by(folio=folio).first_or_404()
    valor = request.form.get('prioridad','').strip()
    try:
        num = int(valor) if valor else None
        if num is not None and num < 1:
            raise ValueError
    except ValueError:
        flash('La prioridad debe ser un número mayor a 0.', 'danger')
        return redirect(url_for('detalle_solicitud', folio=folio))

    sol.prioridad = num
    sol.ultima_actualizacion = cdmx_now()
    label = f'#{num}' if num else 'sin prioridad'
    registrar_bitacora(sol.id, f'{current_user.nombre} asignó prioridad {label}.')
    db.session.commit()
    flash(f'Prioridad actualizada a {label}.', 'success')
    return redirect(url_for('detalle_solicitud', folio=folio))

# ── Acciones en lote ──────────────────────────────────────────────────────────
@app.route('/solicitudes/accion-lote', methods=['POST'])
@login_required
@rol_requerido('lider_comercial','lider_soluciones','administrador')
def accion_lote():
    folios  = request.form.getlist('folios')
    accion  = request.form.get('accion')

    if not folios:
        flash('No seleccionaste ninguna solicitud.', 'warning')
        return redirect(url_for('solicitudes'))
    if accion not in ('eliminar','cerrar'):
        flash('Acción no válida.', 'danger')
        return redirect(url_for('solicitudes'))

    # Permiso: solo admin y lider_comercial pueden eliminar
    if accion == 'eliminar' and current_user.rol not in ('administrador','lider_comercial'):
        flash('No tienes permiso para eliminar solicitudes.', 'danger')
        return redirect(url_for('solicitudes'))

    # Permiso: solo admin y lider_soluciones pueden cerrar
    if accion == 'cerrar' and current_user.rol not in ('administrador','lider_soluciones'):
        flash('No tienes permiso para cerrar solicitudes.', 'danger')
        return redirect(url_for('solicitudes'))

    procesadas = 0
    omitidas   = 0

    for folio in folios:
        sol = Solicitud.query.filter_by(folio=folio).first()
        if not sol:
            continue

        if accion == 'eliminar':
            Comentario.query.filter_by(solicitud_id=sol.id).delete()
            Bitacora.query.filter_by(solicitud_id=sol.id).delete()
            Documento.query.filter_by(solicitud_id=sol.id).delete()
            db.session.delete(sol)
            procesadas += 1

        elif accion == 'cerrar':
            if sol.estatus == 'Cerrada':
                omitidas += 1
                continue
            sol.estatus           = 'Cerrada'
            sol.fecha_cierre      = cdmx_now()
            sol.usuario_cierre_id = current_user.id
            sol.ultima_actualizacion = cdmx_now()
            registrar_bitacora(sol.id, f'{current_user.nombre} cerró la solicitud (acción en lote).')
            procesadas += 1

    db.session.commit()

    msg = f'{procesadas} solicitud{"es" if procesadas != 1 else ""} '
    msg += 'eliminada' + ('s' if procesadas != 1 else '') if accion == 'eliminar'           else 'cerrada' + ('s' if procesadas != 1 else '')
    if omitidas:
        msg += f' ({omitidas} ya estaban cerradas, omitidas).'
    flash(msg + '.', 'success')
    return redirect(url_for('solicitudes'))

# ══════════════════════════════════════════════════════════════════════════════
# ── EXPORTACIONES ────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

@app.route('/exportar/solicitudes')
@login_required
@rol_requerido('administrador','lider_comercial','lider_soluciones','aux_comercial')
def exportar_solicitudes():
    q = Solicitud.query

    # Aplicar mismos filtros de rol que en la vista
    if current_user.rol in ('lider_comercial','aux_comercial'):
        ids = [u.id for u in User.query.filter(User.rol.in_(['comercial','aux_comercial'])).all()]
        ids.append(current_user.id)
        q = q.filter(Solicitud.hunter_id.in_(ids))
    elif current_user.rol == 'lider_soluciones':
        pass  # ve todas
    # administrador ve todas sin filtro

    # Filtros opcionales por querystring
    f_estatus   = request.args.get('f_estatus','').strip()
    f_comercial = request.args.get('f_comercial','').strip()
    f_tema      = request.args.get('f_tema','').strip()
    if f_estatus:   q = q.filter(Solicitud.estatus == f_estatus)
    if f_comercial: q = q.join(User, Solicitud.hunter_id == User.id).filter(User.id == int(f_comercial))
    if f_tema:      q = q.filter(Solicitud.tema == f_tema)

    solicitudes = q.order_by(Solicitud.fecha_captura.desc()).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Solicitudes"

    # Estilos
    header_fill = PatternFill("solid", fgColor="0F172A")
    header_font = Font(bold=True, color="4BB8C8", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border_side = Side(style="thin", color="2A3040")
    cell_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    alt_fill = PatternFill("solid", fgColor="161B22")

    headers = [
        "Folio", "Fecha Captura", "Fecha Solicitud", "Cliente", "Tipo",
        "Comercial", "Ingeniero", "Prioridad", "Estatus",
        "Monto Oportunidad", "Días Captura", "Días Sin Mov.",
        "Fecha Compromiso", "Fecha Envío Cliente", "Fecha Cierre",
        "Hist. Surtido", "Inventario", "Maestro Prod.", "Hist. Recepción", "Cuestionario Log."
    ]
    col_widths = [16,16,16,28,22,22,22,10,24,18,12,12,16,18,14,14,12,14,15,16]

    ws.row_dimensions[1].height = 30
    for col_num, (header, width) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = cell_border
        ws.column_dimensions[get_column_letter(col_num)].width = width

    for row_num, s in enumerate(solicitudes, 2):
        fill = alt_fill if row_num % 2 == 0 else PatternFill("solid", fgColor="1A1F27")
        font = Font(color="E2E8F0", size=10)
        row_data = [
            s.folio,
            s.fecha_captura.strftime('%d/%m/%Y %H:%M') if s.fecha_captura else '',
            s.fecha_solicitud.strftime('%d/%m/%Y') if s.fecha_solicitud else '',
            s.cliente,
            s.tema,
            s.hunter.nombre if s.hunter else '',
            s.responsable.nombre if s.responsable else 'Sin asignar',
            s.prioridad,
            s.estatus,
            s.monto_oportunidad or 0,
            s.dias_desde_captura(),
            s.dias_sin_movimiento(),
            s.fecha_compromiso.strftime('%d/%m/%Y') if s.fecha_compromiso else '',
            s.fecha_envio_cliente.strftime('%d/%m/%Y') if s.fecha_envio_cliente else '',
            s.fecha_cierre.strftime('%d/%m/%Y') if s.fecha_cierre else '',
            'Sí' if s.historial_surtido else 'No',
            'Sí' if s.inventario else 'No',
            'Sí' if s.maestro_productos else 'No',
            'Sí' if s.historial_recepcion else 'No',
            'Sí' if s.cuestionario_logistico else 'No',
        ]
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.font = font
            cell.fill = fill
            cell.border = cell_border
            cell.alignment = Alignment(vertical="center")
            # Formato moneda para monto
            if col_num == 10:
                cell.number_format = '"$"#,##0'
                cell.font = Font(color="AECA00", bold=True, size=10)
            # Color estatus
            if col_num == 9:
                colores_estatus = {
                    'Capturada': 'E2E8F0', 'Asignada': '4BB8C8',
                    'En Análisis': 'F59E0B', 'Información Completa': 'AECA00',
                    'Propuesta Enviada': '7FAD00', 'Cerrada': '6B7280',
                    'Pendiente Información Cliente': 'EF4444'
                }
                color = colores_estatus.get(value, 'E2E8F0')
                cell.font = Font(color=color, bold=True, size=10)

    # Freeze header row
    ws.freeze_panes = "A2"

    # Auto-filter
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # Summary row
    total_row = len(solicitudes) + 3
    ws.cell(row=total_row, column=1, value="TOTAL SOLICITUDES").font = Font(bold=True, color="4BB8C8", size=11)
    ws.cell(row=total_row, column=2, value=len(solicitudes)).font = Font(bold=True, color="E2E8F0", size=11)
    ws.cell(row=total_row, column=9, value="MONTO TOTAL").font = Font(bold=True, color="4BB8C8", size=11)
    total_monto = sum(s.monto_oportunidad or 0 for s in solicitudes)
    monto_cell = ws.cell(row=total_row, column=10, value=total_monto)
    monto_cell.font = Font(bold=True, color="AECA00", size=11)
    monto_cell.number_format = '"$"#,##0'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    fecha_str = cdmx_now().strftime('%Y%m%d_%H%M')
    filename = f"SeguimientoSD_Solicitudes_{fecha_str}.xlsx"
    return Response(
        buf.getvalue(),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


@app.route('/exportar/usuarios')
@login_required
@rol_requerido('administrador','lider_comercial','lider_soluciones','aux_comercial')
def exportar_usuarios():
    if es_admin():
        usuarios = User.query.order_by(User.rol, User.nombre).all()
    elif current_user.rol in ('lider_comercial','aux_comercial'):
        usuarios = User.query.filter(User.rol.in_(['comercial','aux_comercial','lider_comercial']))                             .order_by(User.nombre).all()
    else:
        usuarios = User.query.filter(User.rol.in_(['ingeniero','lider_soluciones']))                             .order_by(User.nombre).all()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Nombre', 'Usuario', 'Rol', 'Estatus', 'Fecha Creación',
                     'Solicitudes Creadas', 'Solicitudes Asignadas'])
    for u in usuarios:
        writer.writerow([
            u.nombre,
            u.username,
            ROLES_LABEL.get(u.rol, u.rol),
            'Activo' if u.activo else 'Inactivo',
            u.created_at.strftime('%d/%m/%Y') if u.created_at else '',
            u.solicitudes_creadas.count(),
            u.solicitudes_asignadas.count(),
        ])

    fecha_str = cdmx_now().strftime('%Y%m%d_%H%M')
    filename = f"SeguimientoSD_Usuarios_{fecha_str}.csv"
    return Response(
        '﻿' + output.getvalue(),   # BOM para que Excel lo abra bien
        mimetype='text/csv; charset=utf-8',
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# ── Init DB ───────────────────────────────────────────────────────────────────
def init_db():
    db.create_all()

    # Seed temas
    if TemasSolicitud.query.count() == 0:
        for i, nombre in enumerate(TEMAS_DEFAULT):
            db.session.add(TemasSolicitud(nombre=nombre, orden=i))
        db.session.commit()

    if User.query.count() == 0:
        users = [
            User(username='admin',               password_hash=generate_password_hash('Admin2026!'),     nombre='Administrador',       rol='administrador'),
            User(username='Francisco_Cueva',     password_hash=generate_password_hash('Lcomercial123'),  nombre='Francisco Cueva',     rol='lider_comercial'),
            User(username='Andrés_Toledo',       password_hash=generate_password_hash('Lsoluciones123'), nombre='Andrés Toledo',        rol='lider_soluciones'),
            User(username='Gerardo_Velazquez',   password_hash=generate_password_hash('IngeSD1231'),     nombre='Gerardo Velazquez',   rol='ingeniero'),
            User(username='Elizabeth_Bastida',   password_hash=generate_password_hash('IngeSD1232'),     nombre='Elizabeth Bastida',   rol='ingeniero'),
            User(username='Diego_Arzate',        password_hash=generate_password_hash('IgeSDT123'),      nombre='Diego Arzate',        rol='ingeniero'),
            User(username='Jorge_Camarena',      password_hash=generate_password_hash('IngeSD1233'),     nombre='Jorge Camarena',      rol='ingeniero'),
            User(username='Teresa_Ruiz',         password_hash=generate_password_hash('Hunter1231'),     nombre='Teresa Ruiz',         rol='comercial'),
            User(username='Alejandra_Sanchez',   password_hash=generate_password_hash('Hunter1232'),     nombre='Alejandra Sánchez',   rol='comercial'),
            User(username='Diana_Pelcastre',     password_hash=generate_password_hash('Hunter1233'),     nombre='Diana Pelcastre',     rol='comercial'),
            User(username='Ida_Acosta',          password_hash=generate_password_hash('Hunter1234'),     nombre='Ida Acosta',          rol='comercial'),
            User(username='Malena_Baltazar',     password_hash=generate_password_hash('Hunter1235'),     nombre='Malena Baltazar',     rol='comercial'),
            User(username='José_Ortega',         password_hash=generate_password_hash('Hunter1236'),     nombre='José Ortega',         rol='comercial'),
            User(username='aline_esquivel',      password_hash=generate_password_hash('AuxCom2026!'),    nombre='Aline Esquivel',      rol='aux_comercial'),
            User(username='rubi_arizmendi',      password_hash=generate_password_hash('AuxCom2026!'),    nombre='Rubí Arizmendi',      rol='aux_comercial'),
        ]
        for u in users:
            db.session.add(u)
        db.session.flush()

        demos = [
            dict(hunter_id=8,  responsable_id=4, fecha_solicitud=date(2026, 5, 10),
                 cliente='FEMSA Logística', tema='Transporte',
                 monto_oportunidad=850000, prioridad=1, estatus='En Análisis',
                 historial_surtido=True, inventario=True),
            dict(hunter_id=9,  responsable_id=5, fecha_solicitud=date(2026, 5, 20),
                 cliente='Grupo Bimbo', tema='Almacenaje',
                 monto_oportunidad=1200000, prioridad=1, estatus='Información Completa',
                 historial_surtido=True, inventario=True, maestro_productos=True,
                 historial_recepcion=True, cuestionario_logistico=True),
            dict(hunter_id=10, responsable_id=6, fecha_solicitud=date.today(),
                 cliente='Liverpool', tema='VAS',
                 monto_oportunidad=500000, prioridad=2, estatus='Capturada'),
            dict(hunter_id=11, responsable_id=None, fecha_solicitud=date.today(),
                 cliente='Soriana', tema='Reingeniería',
                 monto_oportunidad=3000000, prioridad=1, estatus='Capturada'),
            dict(hunter_id=12, responsable_id=7, fecha_solicitud=date(2026, 4, 15),
                 cliente='Amazon MX', tema='Transporte y almacenaje',
                 monto_oportunidad=4500000, prioridad=1, estatus='Propuesta Enviada',
                 historial_surtido=True, inventario=True, maestro_productos=True,
                 historial_recepcion=True, cuestionario_logistico=True,
                 fecha_envio_cliente=datetime(2026, 5, 28), usuario_envio_id=7,
                 comentarios_envio='Se envió propuesta económica completa al cliente.'),
        ]
        for d in demos:
            s = Solicitud(folio=generar_folio(), **d)
            db.session.add(s)
            db.session.flush()
            hunter = db.session.get(User, s.hunter_id)
            registrar_bitacora(s.id, f'{hunter.nombre} creó la solicitud.', s.hunter_id)

        db.session.commit()
        print('✅ Base de datos inicializada con usuarios y temas.')


with app.app_context():
    init_db()

if __name__ == '__main__':
    port  = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
