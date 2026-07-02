from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime, timezone, timedelta

db = SQLAlchemy()

CDMX_OFFSET = timedelta(hours=-6)

def cdmx_now():
    """Hora actual en Ciudad de México (GMT-6, sin DST)."""
    return datetime.now(timezone.utc).replace(tzinfo=None) + CDMX_OFFSET

# Alias para compatibilidad
utcnow = cdmx_now


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
    subtipo             = db.Column(db.String(10),  nullable=True)   # RFQ, RFI, RFP
    solicitud_origen_id = db.Column(db.Integer, db.ForeignKey('solicitudes.id'), nullable=True)  # Retrabajo: ligada a solicitud cerrada

    # Prioridad — asignada directamente por Líder de Soluciones
    prioridad_sugerida   = db.Column(db.Integer, nullable=True)  # DEPRECADO — mantenido por compatibilidad
    prioridad_comercial  = db.Column(db.Integer, nullable=True)  # DEPRECADO — mantenido por compatibilidad
    prioridad            = db.Column(db.Integer, nullable=True)  # Asignada por lider_soluciones
    prioridad_estado     = db.Column(db.String(20), default='pendiente')
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
        if self.estatus == 'Cerrada':
            return 0
        ref = self.ultima_actualizacion or self.fecha_captura
        return (cdmx_now() - ref).days

    def dias_desde_captura(self):
        return (cdmx_now() - self.fecha_captura).days

    # Relación con solicitud origen (retrabajo)
    solicitud_origen = db.relationship('Solicitud', foreign_keys=[solicitud_origen_id],
                                       remote_side='Solicitud.id',
                                       backref=db.backref('retrabajos', lazy='dynamic'))

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

class SolicitudIngeniero(db.Model):
    """Tabla intermedia para múltiples ingenieros por solicitud."""
    __tablename__ = 'solicitud_ingenieros'
    id            = db.Column(db.Integer, primary_key=True)
    solicitud_id  = db.Column(db.Integer, db.ForeignKey('solicitudes.id'), nullable=False)
    ingeniero_id  = db.Column(db.Integer, db.ForeignKey('users.id'),       nullable=False)
    es_principal  = db.Column(db.Boolean, default=False)  # True = responsable principal
    created_at    = db.Column(db.DateTime, default=utcnow)

    solicitud = db.relationship('Solicitud', foreign_keys=[solicitud_id],
                                backref=db.backref('ingenieros_asignados', lazy='dynamic'))
    ingeniero = db.relationship('User', foreign_keys=[ingeniero_id])


class RutaTransporte(db.Model):
    __tablename__ = 'rutas_transporte'
    id                = db.Column(db.Integer, primary_key=True)
    solicitud_id      = db.Column(db.Integer, db.ForeignKey('solicitudes.id'), nullable=False)
    orden             = db.Column(db.Integer, default=1)
    # Origen: estado + ciudad separados para homologar
    origen_estado     = db.Column(db.String(100), nullable=False)
    origen_ciudad     = db.Column(db.String(100), nullable=False)
    # Destino: estado + ciudad separados
    destino_estado    = db.Column(db.String(100), nullable=False)
    destino_ciudad    = db.Column(db.String(100), nullable=False)
    tipo_servicio     = db.Column(db.String(20),  nullable=False)  # FTL, LTL, Mensajería
    # FTL/Mensajería → tipo_unidad; LTL → kg_por_entrega y m3_por_entrega
    tipo_unidad       = db.Column(db.String(50),  nullable=True)   # 1.5 Ton, 3.5T, Rabón, Torton, Trailer
    peso_kg           = db.Column(db.Float,       nullable=True)   # Peso en kg (FTL/Mensajería)
    kg_por_entrega    = db.Column(db.Float,       nullable=True)   # LTL: kg por entrega
    m3_por_entrega    = db.Column(db.Float,       nullable=True)   # LTL: m3 por entrega
    temperatura       = db.Column(db.String(50),  nullable=True)
    custodia          = db.Column(db.String(30),  nullable=True)
    comentarios       = db.Column(db.Text,        nullable=True)
    created_at        = db.Column(db.DateTime, default=utcnow)

    solicitud = db.relationship('Solicitud', foreign_keys=[solicitud_id],
                                backref=db.backref('rutas', lazy='dynamic',
                                                   order_by='RutaTransporte.orden'))

    @property
    def calidad_info(self):
        """Porcentaje de checkboxes completados (0-100)."""
        checks = [self.historial_surtido, self.inventario, self.maestro_productos,
                  self.historial_recepcion, self.cuestionario_logistico]
        return round(sum(1 for c in checks if c) / len(checks) * 100)

    @property
    def origen(self):
        return f"{self.origen_ciudad}, {self.origen_estado}"

    @property
    def destino(self):
        return f"{self.destino_ciudad}, {self.destino_estado}" 
