from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime, timezone

db = SQLAlchemy()


def utcnow():
    return datetime.now(timezone.utc).replace(tzinfo=None)


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id            = db.Column(db.Integer, primary_key=True)
    username      = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    nombre        = db.Column(db.String(120), nullable=False)
    rol           = db.Column(db.String(30), nullable=False)
    activo        = db.Column(db.Boolean, default=True)
    created_at    = db.Column(db.DateTime, default=utcnow)

    solicitudes_creadas   = db.relationship('Solicitud', foreign_keys='Solicitud.hunter_id',     backref='hunter',     lazy='dynamic')
    solicitudes_asignadas = db.relationship('Solicitud', foreign_keys='Solicitud.responsable_id', backref='responsable', lazy='dynamic')


# Catálogo de temas — el Administrador puede agregar más
class TemasSolicitud(db.Model):
    __tablename__ = 'temas_solicitud'
    id     = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), unique=True, nullable=False)
    activo = db.Column(db.Boolean, default=True)
    orden  = db.Column(db.Integer, default=0)


class Solicitud(db.Model):
    __tablename__ = 'solicitudes'
    id            = db.Column(db.Integer, primary_key=True)
    folio         = db.Column(db.String(20), unique=True, nullable=False)
    hunter_id     = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    responsable_id= db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    fecha_captura = db.Column(db.DateTime, default=utcnow)
    fecha_solicitud = db.Column(db.Date, nullable=False)
    cliente       = db.Column(db.String(200), nullable=False)
    tema          = db.Column(db.String(300), nullable=False)
    comentarios_comerciales = db.Column(db.Text, nullable=True)
    monto_oportunidad = db.Column(db.Float, nullable=True)
    prioridad     = db.Column(db.String(10), nullable=False, default='Media')
    estatus       = db.Column(db.String(40), nullable=False, default='Capturada')
    ultima_actualizacion = db.Column(db.DateTime, default=utcnow, onupdate=utcnow)

    fecha_info_completa   = db.Column(db.DateTime, nullable=True)
    historial_surtido     = db.Column(db.Boolean, default=False)
    inventario            = db.Column(db.Boolean, default=False)
    maestro_productos     = db.Column(db.Boolean, default=False)
    historial_recepcion   = db.Column(db.Boolean, default=False)
    cuestionario_logistico= db.Column(db.Boolean, default=False)
    fecha_compromiso      = db.Column(db.Date, nullable=True)

    fecha_envio_cliente   = db.Column(db.DateTime, nullable=True)
    usuario_envio_id      = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    comentarios_envio     = db.Column(db.Text, nullable=True)

    fecha_cierre          = db.Column(db.DateTime, nullable=True)
    usuario_cierre_id     = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)

    usuario_envio  = db.relationship('User', foreign_keys=[usuario_envio_id])
    usuario_cierre = db.relationship('User', foreign_keys=[usuario_cierre_id])
    comentarios    = db.relationship('Comentario', backref='solicitud', lazy='dynamic', order_by='Comentario.created_at')
    bitacora       = db.relationship('Bitacora',   backref='solicitud', lazy='dynamic', order_by='Bitacora.created_at')

    def checkboxes_completos(self):
        return all([self.historial_surtido, self.inventario,
                    self.maestro_productos, self.historial_recepcion,
                    self.cuestionario_logistico])

    def dias_sin_movimiento(self):
        ref = self.ultima_actualizacion or self.fecha_captura
        return (datetime.utcnow() - ref).days

    def dias_desde_captura(self):
        return (datetime.utcnow() - self.fecha_captura).days

    def actualizar_estatus_automatico(self):
        if self.estatus == 'Cerrada':
            return
        if self.fecha_envio_cliente:
            self.estatus = 'Propuesta Enviada'
        elif self.checkboxes_completos():
            self.estatus = 'Información Completa'


class Comentario(db.Model):
    __tablename__ = 'comentarios'
    id           = db.Column(db.Integer, primary_key=True)
    solicitud_id = db.Column(db.Integer, db.ForeignKey('solicitudes.id'), nullable=False)
    usuario_id   = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    texto        = db.Column(db.Text, nullable=False)
    created_at   = db.Column(db.DateTime, default=utcnow)
    usuario      = db.relationship('User', foreign_keys=[usuario_id])


class Bitacora(db.Model):
    __tablename__ = 'bitacora'
    id           = db.Column(db.Integer, primary_key=True)
    solicitud_id = db.Column(db.Integer, db.ForeignKey('solicitudes.id'), nullable=False)
    usuario_id   = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    accion       = db.Column(db.Text, nullable=False)
    created_at   = db.Column(db.DateTime, default=utcnow)
    usuario      = db.relationship('User', foreign_keys=[usuario_id])


class Documento(db.Model):
    __tablename__      = 'documentos'
    id                 = db.Column(db.Integer, primary_key=True)
    solicitud_id       = db.Column(db.Integer, db.ForeignKey('solicitudes.id'), nullable=False)
    nombre_original    = db.Column(db.String(255), nullable=False)
    nombre_guardado    = db.Column(db.String(255), nullable=False)
    tipo_documento     = db.Column(db.String(20),  nullable=False)
    usuario_id         = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    fecha_subida       = db.Column(db.DateTime, default=utcnow)
    fecha_modificacion = db.Column(db.DateTime, default=utcnow, onupdate=utcnow)
    version            = db.Column(db.Integer, default=1, nullable=False)
    activo             = db.Column(db.Boolean, default=True)
    documento_padre_id = db.Column(db.Integer, db.ForeignKey('documentos.id'), nullable=True)

    usuario   = db.relationship('User',     foreign_keys=[usuario_id])
    solicitud = db.relationship('Solicitud', foreign_keys=[solicitud_id],
                                backref=db.backref('documentos', lazy='dynamic'))
    versiones_anteriores = db.relationship('Documento', foreign_keys=[documento_padre_id],
                                           backref=db.backref('padre', remote_side=[id]),
                                           lazy='dynamic')
