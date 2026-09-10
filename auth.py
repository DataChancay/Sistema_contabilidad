from __future__ import annotations

import hmac
import os
import re
import secrets
import smtplib
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from functools import wraps
from typing import Iterable
from urllib.parse import urljoin

from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError
from argon2.low_level import Type
from flask import (
    Blueprint,
    abort,
    current_app,
    flash,
    has_request_context,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_login import LoginManager, UserMixin, current_user, login_required, login_user, logout_user
from flask_wtf import CSRFProtect
from sqlalchemy import (
    JSON,
    Boolean,
    Column,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    String,
    Table,
    Text,
    UniqueConstraint,
    Uuid,
    create_engine,
    func,
    select,
)
from sqlalchemy.engine import Engine
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship, scoped_session, sessionmaker
from sqlalchemy.pool import StaticPool
from werkzeug.middleware.proxy_fix import ProxyFix


USER_STATUSES = {"pending", "active", "inactive", "blocked"}
PASSWORD_TOKEN_PURPOSES = {"activation", "password_reset"}
PASSWORD_MIN_LENGTH = 12

DEFAULT_PERMISSIONS = (
    "usuarios.ver",
    "usuarios.crear",
    "usuarios.editar",
    "usuarios.desactivar",
    "roles.ver",
    "roles.editar",
    "reportes.ver",
    "reportes.exportar",
    "configuracion.ver",
    "configuracion.editar",
    "auditoria.ver",
)

DEFAULT_ROLES = {
    "superadmin": DEFAULT_PERMISSIONS,
    "admin": (
        "usuarios.ver",
        "usuarios.crear",
        "usuarios.editar",
        "usuarios.desactivar",
        "roles.ver",
        "reportes.ver",
        "reportes.exportar",
        "auditoria.ver",
    ),
    "supervisor": ("reportes.ver", "reportes.exportar", "auditoria.ver"),
    "user": ("reportes.ver", "reportes.exportar"),
}


class Base(DeclarativeBase):
    pass


user_roles = Table(
    "user_roles",
    Base.metadata,
    Column("user_id", Uuid(as_uuid=False), ForeignKey("users.id", ondelete="CASCADE"), primary_key=True),
    Column("role_id", Uuid(as_uuid=False), ForeignKey("roles.id", ondelete="CASCADE"), primary_key=True),
    Column("created_at", DateTime(timezone=True), nullable=False, server_default=func.now()),
)

role_permissions = Table(
    "role_permissions",
    Base.metadata,
    Column("role_id", Uuid(as_uuid=False), ForeignKey("roles.id", ondelete="CASCADE"), primary_key=True),
    Column("permission_id", Uuid(as_uuid=False), ForeignKey("permissions.id", ondelete="CASCADE"), primary_key=True),
    Column("created_at", DateTime(timezone=True), nullable=False, server_default=func.now()),
)


class User(Base, UserMixin):
    __tablename__ = "users"

    id: Mapped[str] = mapped_column(Uuid(as_uuid=False), primary_key=True, default=lambda: str(uuid.uuid4()))
    name: Mapped[str] = mapped_column(String(120), nullable=False)
    email: Mapped[str] = mapped_column(String(255), nullable=False, unique=True)
    password_hash: Mapped[str | None] = mapped_column(Text, nullable=True)
    status: Mapped[str] = mapped_column(String(20), nullable=False, default="pending")
    last_login_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    session_version: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    mfa_enabled: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    mfa_enrolled_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        onupdate=func.now(),
    )
    roles: Mapped[list["Role"]] = relationship("Role", secondary=user_roles, back_populates="users")

    def is_active(self) -> bool:
        return self.status == "active"


class Role(Base):
    __tablename__ = "roles"

    id: Mapped[str] = mapped_column(Uuid(as_uuid=False), primary_key=True, default=lambda: str(uuid.uuid4()))
    name: Mapped[str] = mapped_column(String(80), nullable=False, unique=True)
    description: Mapped[str | None] = mapped_column(String(255), nullable=True)
    is_system: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        onupdate=func.now(),
    )
    users: Mapped[list[User]] = relationship("User", secondary=user_roles, back_populates="roles")
    permissions: Mapped[list["Permission"]] = relationship(
        "Permission",
        secondary=role_permissions,
        back_populates="roles",
    )


class Permission(Base):
    __tablename__ = "permissions"

    id: Mapped[str] = mapped_column(Uuid(as_uuid=False), primary_key=True, default=lambda: str(uuid.uuid4()))
    code: Mapped[str] = mapped_column(String(120), nullable=False, unique=True)
    description: Mapped[str | None] = mapped_column(String(255), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, server_default=func.now())
    roles: Mapped[list[Role]] = relationship("Role", secondary=role_permissions, back_populates="permissions")


class PasswordToken(Base):
    __tablename__ = "password_tokens"
    __table_args__ = (
        UniqueConstraint("token_hash", name="uq_password_tokens_token_hash"),
        Index("ix_password_tokens_user_id", "user_id"),
        Index("ix_password_tokens_expires_at", "expires_at"),
    )

    id: Mapped[str] = mapped_column(Uuid(as_uuid=False), primary_key=True, default=lambda: str(uuid.uuid4()))
    user_id: Mapped[str] = mapped_column(Uuid(as_uuid=False), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    token_hash: Mapped[str] = mapped_column(String(128), nullable=False)
    purpose: Mapped[str] = mapped_column(String(30), nullable=False)
    expires_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    used_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, server_default=func.now())
    user: Mapped[User] = relationship("User")


class AuditLog(Base):
    __tablename__ = "audit_logs"
    __table_args__ = (
        Index("ix_audit_logs_user_id", "user_id"),
        Index("ix_audit_logs_created_at", "created_at"),
        Index("ix_audit_logs_action", "action"),
    )

    id: Mapped[str] = mapped_column(Uuid(as_uuid=False), primary_key=True, default=lambda: str(uuid.uuid4()))
    user_id: Mapped[str | None] = mapped_column(Uuid(as_uuid=False), ForeignKey("users.id", ondelete="SET NULL"), nullable=True)
    action: Mapped[str] = mapped_column(String(120), nullable=False)
    entity_type: Mapped[str | None] = mapped_column(String(80), nullable=True)
    entity_id: Mapped[str | None] = mapped_column(String(80), nullable=True)
    ip_address: Mapped[str | None] = mapped_column(String(80), nullable=True)
    user_agent: Mapped[str | None] = mapped_column(String(300), nullable=True)
    audit_metadata: Mapped[dict | None] = mapped_column("metadata", JSON, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, server_default=func.now())
    user: Mapped[User | None] = relationship("User")


SessionLocal = scoped_session(sessionmaker(autoflush=False, expire_on_commit=False))
_engine: Engine | None = None

login_manager = LoginManager()
csrf = CSRFProtect()
limiter = Limiter(key_func=get_remote_address)
auth_bp = Blueprint("auth", __name__)

password_hasher = PasswordHasher(
    time_cost=3,
    memory_cost=65536,
    parallelism=4,
    hash_len=32,
    salt_len=16,
    type=Type.ID,
)


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


def get_database_url() -> str:
    explicit = os.getenv("DATABASE_URL", "").strip()
    if explicit:
        return explicit
    host = os.getenv("DB_HOST", "").strip()
    name = os.getenv("DB_NAME", "").strip()
    user = os.getenv("DB_USER", "").strip()
    password = os.getenv("DB_PASSWORD", "").strip()
    port = os.getenv("DB_PORT", "5432").strip() or "5432"
    if not all((host, name, user, password)):
        return "sqlite:///auth-dev.sqlite3"
    return f"postgresql+psycopg://{user}:{password}@{host}:{port}/{name}"


def configure_database(database_url: str | None = None) -> Engine:
    global _engine
    if _engine is not None:
        SessionLocal.remove()
        _engine.dispose()
    url = database_url or get_database_url()
    kwargs: dict[str, object] = {"future": True, "pool_pre_ping": True}
    if url.startswith("sqlite://"):
        kwargs["connect_args"] = {"check_same_thread": False}
        if url in {"sqlite://", "sqlite:///:memory:"}:
            kwargs["poolclass"] = StaticPool
    _engine = create_engine(url, **kwargs)
    SessionLocal.configure(bind=_engine)
    return _engine


def get_engine() -> Engine:
    global _engine
    if _engine is None:
        return configure_database()
    return _engine


def db_session():
    get_engine()
    return SessionLocal()


@contextmanager
def transaction():
    db = db_session()
    try:
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


def create_schema() -> None:
    Base.metadata.create_all(get_engine())
    with transaction() as db:
        seed_rbac(db)


def seed_rbac(db) -> None:
    permissions = {permission.code: permission for permission in db.scalars(select(Permission)).all()}
    for code in DEFAULT_PERMISSIONS:
        if code not in permissions:
            permission = Permission(code=code, description=code)
            db.add(permission)
            permissions[code] = permission
    roles = {role.name: role for role in db.scalars(select(Role)).all()}
    for name, role_permissions_ in DEFAULT_ROLES.items():
        role = roles.get(name)
        if role is None:
            role = Role(name=name, description=name, is_system=True)
            db.add(role)
            roles[name] = role
        role.permissions = [permissions[code] for code in role_permissions_ if code in permissions]


def normalize_email(raw_email: str) -> str:
    return raw_email.strip().lower()


def validate_email(raw_email: str) -> str:
    email = normalize_email(raw_email)
    if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
        raise ValueError("Ingresa un correo valido.")
    if len(email) > 255:
        raise ValueError("El correo es demasiado largo.")
    return email


def validate_name(raw_name: str) -> str:
    name = " ".join(raw_name.strip().split())
    if len(name) < 2 or len(name) > 120:
        raise ValueError("El nombre debe tener entre 2 y 120 caracteres.")
    return name


def validate_password(raw_password: str) -> str:
    if len(raw_password) < PASSWORD_MIN_LENGTH:
        raise ValueError(f"La contrasena debe tener al menos {PASSWORD_MIN_LENGTH} caracteres.")
    if len(raw_password) > 1024:
        raise ValueError("La contrasena es demasiado larga.")
    return raw_password


def hash_password(password: str) -> str:
    return password_hasher.hash(validate_password(password))


def verify_password(password_hash: str | None, password: str) -> bool:
    if not password_hash:
        return False
    try:
        return password_hasher.verify(password_hash, password)
    except VerifyMismatchError:
        return False
    except Exception:
        return False


def token_digest(token: str) -> str:
    return hmac.digest(b"auth-token-v1", token.encode("utf-8"), "sha256").hex()


def create_password_token(db, user: User, purpose: str, expires_in: timedelta | None = None) -> str:
    if purpose not in PASSWORD_TOKEN_PURPOSES:
        raise ValueError("Proposito de token invalido.")
    expires_delta = expires_in or timedelta(hours=24)
    db.query(PasswordToken).filter(
        PasswordToken.user_id == user.id,
        PasswordToken.purpose == purpose,
        PasswordToken.used_at.is_(None),
    ).update({"used_at": utcnow()})
    token = secrets.token_urlsafe(48)
    db.add(
        PasswordToken(
            user_id=user.id,
            token_hash=token_digest(token),
            purpose=purpose,
            expires_at=utcnow() + expires_delta,
        )
    )
    return token


def consume_password_token(db, token: str, purpose: str) -> User | None:
    if purpose not in PASSWORD_TOKEN_PURPOSES:
        return None
    stored = db.scalar(
        select(PasswordToken).where(
            PasswordToken.token_hash == token_digest(token),
            PasswordToken.purpose == purpose,
            PasswordToken.used_at.is_(None),
            PasswordToken.expires_at > utcnow(),
        )
    )
    if stored is None:
        return None
    stored.used_at = utcnow()
    return stored.user


def request_ip() -> str | None:
    return request.headers.get("X-Forwarded-For", request.remote_addr or "").split(",")[0].strip() or None


def audit(
    db,
    action: str,
    user: User | None = None,
    entity_type: str | None = None,
    entity_id: str | None = None,
    metadata: dict | None = None,
) -> None:
    clean_metadata = metadata or None
    in_request = has_request_context()
    db.add(
        AuditLog(
            user_id=user.id if user else None,
            action=action,
            entity_type=entity_type,
            entity_id=entity_id,
            ip_address=request_ip() if in_request else None,
            user_agent=(request.headers.get("User-Agent", "")[:300] if in_request else None),
            audit_metadata=clean_metadata,
        )
    )


def has_permission(user: User, permission_code: str) -> bool:
    if not user or not getattr(user, "is_authenticated", False) or user.status != "active":
        return False
    for role in user.roles:
        if role.name == "superadmin":
            return True
        if any(permission.code == permission_code for permission in role.permissions):
            return True
    return False


def permission_required(permission_code: str):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if current_app.config.get("AUTH_REQUIRED", True) is False:
                return view(*args, **kwargs)
            if not current_user.is_authenticated:
                return login_manager.unauthorized()
            if not has_permission(current_user, permission_code):
                abort(403)
            if session.get("auth_version") != current_user.session_version or current_user.status != "active":
                logout_user()
                session.clear()
                return redirect(url_for("auth.login", next=request.full_path))
            return view(*args, **kwargs)

        return wrapped

    return decorator


def role_choices(db) -> list[Role]:
    return db.scalars(select(Role).order_by(Role.name)).all()


def permission_choices(db) -> list[Permission]:
    return db.scalars(select(Permission).order_by(Permission.code)).all()


def selected_roles(db, role_ids: Iterable[str]) -> list[Role]:
    clean_ids = [role_id for role_id in role_ids if re.fullmatch(r"[0-9a-fA-F-]{36}", role_id or "")]
    if not clean_ids:
        return []
    return db.scalars(select(Role).where(Role.id.in_(clean_ids))).all()


def safe_next_url(raw_next: str | None) -> str:
    if raw_next and raw_next.startswith("/") and not raw_next.startswith("//"):
        return raw_next
    return url_for("index")


def public_base_url() -> str:
    configured = os.getenv("APP_BASE_URL", "").strip()
    if configured:
        return configured.rstrip("/") + "/"
    return request.url_root


def send_token_email(user: User, token: str, purpose: str) -> bool:
    smtp_host = os.getenv("SMTP_HOST", "").strip()
    smtp_from = os.getenv("SMTP_FROM", "").strip()
    if not smtp_host or not smtp_from:
        return False
    path = f"activate/{token}" if purpose == "activation" else f"password/reset/{token}"
    link = urljoin(public_base_url(), path)
    message = EmailMessage()
    message["From"] = smtp_from
    message["To"] = user.email
    message["Subject"] = "Acceso a Contabilidad Chancay"
    message.set_content(
        "Hola,\n\nUsa el siguiente enlace para definir tu contrasena. "
        "El enlace es de un solo uso y expira automaticamente:\n\n"
        f"{link}\n"
    )
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER", "").strip()
    smtp_password = os.getenv("SMTP_PASSWORD", "")
    with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as smtp:
        smtp.starttls()
        if smtp_user and smtp_password:
            smtp.login(smtp_user, smtp_password)
        smtp.send_message(message)
    return True


@login_manager.user_loader
def load_user(user_id: str) -> User | None:
    if not re.fullmatch(r"[0-9a-fA-F-]{36}", user_id or ""):
        return None
    return db_session().get(User, user_id)


@auth_bp.get("/login")
def login():
    if current_user.is_authenticated:
        return redirect(safe_next_url(request.args.get("next")))
    return render_template("login.html", next_url=safe_next_url(request.args.get("next")))


def login_email_key() -> str:
    if request.form:
        raw_email = request.form.get("email", "")
    else:
        raw_email = (request.get_json(silent=True) or {}).get("email", "")
    email = normalize_email(raw_email)
    return f"{get_remote_address()}:{email or 'empty'}"


@auth_bp.post("/login")
@limiter.limit("10 per minute", key_func=get_remote_address)
@limiter.limit("5 per minute", key_func=login_email_key)
def login_post():
    email = normalize_email(request.form.get("email", ""))
    password = request.form.get("password", "")
    with transaction() as db:
        user = db.scalar(select(User).where(User.email == email))
        valid = bool(user and user.status == "active" and verify_password(user.password_hash, password))
        if not valid:
            audit(db, "auth.login_failed", user=user if user else None, entity_type="user", entity_id=user.id if user else None)
            flash("Credenciales incorrectas.", "error")
            return render_template("login.html", next_url=safe_next_url(request.form.get("next"))), 401
        session.clear()
        login_user(user)
        session.permanent = True
        session["auth_version"] = user.session_version
        user.last_login_at = utcnow()
        audit(db, "auth.login_success", user=user, entity_type="user", entity_id=user.id)
        return redirect(safe_next_url(request.form.get("next")))


@auth_bp.post("/logout")
@login_required
def logout():
    with transaction() as db:
        user = db.get(User, current_user.id)
        audit(db, "auth.logout", user=user, entity_type="user", entity_id=user.id if user else None)
    logout_user()
    session.clear()
    return redirect(url_for("auth.login"))


@auth_bp.get("/password/forgot")
def forgot_password():
    return render_template("forgot_password.html")


@auth_bp.post("/password/forgot")
@limiter.limit("5 per hour", key_func=login_email_key)
def forgot_password_post():
    email = normalize_email(request.form.get("email", ""))
    with transaction() as db:
        user = db.scalar(select(User).where(User.email == email, User.status == "active"))
        if user:
            token = create_password_token(db, user, "password_reset", timedelta(hours=1))
            send_token_email(user, token, "password_reset")
            audit(db, "auth.password_reset_requested", user=user, entity_type="user", entity_id=user.id)
    flash("Si existe una cuenta asociada a ese correo, recibiras las instrucciones.", "success")
    return redirect(url_for("auth.forgot_password"))


@auth_bp.get("/password/reset/<token>")
def reset_password(token: str):
    return render_template("set_password.html", token=token, purpose="password_reset")


@auth_bp.post("/password/reset/<token>")
def reset_password_post(token: str):
    return set_password_from_token(token, "password_reset")


@auth_bp.get("/activate/<token>")
def activate(token: str):
    return render_template("set_password.html", token=token, purpose="activation")


@auth_bp.post("/activate/<token>")
def activate_post(token: str):
    return set_password_from_token(token, "activation")


def set_password_from_token(token: str, purpose: str):
    password = request.form.get("password", "")
    password_confirmation = request.form.get("password_confirmation", "")
    if password != password_confirmation:
        flash("Las contrasenas no coinciden.", "error")
        return render_template("set_password.html", token=token, purpose=purpose), 400
    try:
        password_hash = hash_password(password)
    except ValueError as exc:
        flash(str(exc), "error")
        return render_template("set_password.html", token=token, purpose=purpose), 400
    with transaction() as db:
        user = consume_password_token(db, token, purpose)
        if user is None:
            flash("El enlace no es valido o ya expiro.", "error")
            return render_template("set_password.html", token=token, purpose=purpose), 400
        user.password_hash = password_hash
        user.status = "active"
        user.session_version += 1
        audit(db, "auth.password_changed", user=user, entity_type="user", entity_id=user.id, metadata={"purpose": purpose})
    flash("Tu contrasena fue actualizada. Ya puedes iniciar sesion.", "success")
    return redirect(url_for("auth.login"))


@auth_bp.get("/admin")
@permission_required("usuarios.ver")
def admin_home():
    return redirect(url_for("auth.admin_users"))


@auth_bp.get("/admin/usuarios")
@permission_required("usuarios.ver")
def admin_users():
    q = request.args.get("q", "").strip()
    with transaction() as db:
        query = select(User).order_by(User.created_at.desc())
        if q:
            like = f"%{q.lower()}%"
            query = query.where((func.lower(User.email).like(like)) | (func.lower(User.name).like(like)))
        users = db.scalars(query.limit(100)).all()
        return render_template("admin_users.html", users=users, roles=role_choices(db), q=q)


@auth_bp.post("/admin/usuarios")
@permission_required("usuarios.crear")
def admin_users_create():
    try:
        name = validate_name(request.form.get("name", ""))
        email = validate_email(request.form.get("email", ""))
        roles = request.form.getlist("roles")
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("auth.admin_users")), 400
    with transaction() as db:
        if db.scalar(select(User).where(User.email == email)):
            flash("El correo ya esta registrado.", "error")
            return redirect(url_for("auth.admin_users")), 409
        user = User(name=name, email=email, status="pending")
        user.roles = selected_roles(db, roles)
        db.add(user)
        db.flush()
        token = create_password_token(db, user, "activation", timedelta(hours=24))
        emailed = send_token_email(user, token, "activation")
        audit(db, "admin.user_created", user=current_user, entity_type="user", entity_id=user.id, metadata={"email_sent": emailed})
        if current_app.config.get("AUTH_SHOW_DEV_TOKENS"):
            flash(f"Usuario creado. Enlace temporal: {urljoin(public_base_url(), 'activate/' + token)}", "success")
        else:
            flash("Usuario creado. Configura SMTP para enviar el enlace de activacion automaticamente." if not emailed else "Usuario creado y enlace enviado.", "success")
    return redirect(url_for("auth.admin_users"))


@auth_bp.get("/admin/usuarios/<user_id>")
@permission_required("usuarios.editar")
def admin_user_edit(user_id: str):
    with transaction() as db:
        user = db.get(User, user_id)
        if user is None:
            abort(404)
        return render_template("admin_user_edit.html", edited_user=user, roles=role_choices(db), statuses=sorted(USER_STATUSES))


@auth_bp.post("/admin/usuarios/<user_id>")
@permission_required("usuarios.editar")
def admin_user_update(user_id: str):
    try:
        name = validate_name(request.form.get("name", ""))
        email = validate_email(request.form.get("email", ""))
        status = request.form.get("status", "")
        if status not in USER_STATUSES:
            raise ValueError("Estado invalido.")
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("auth.admin_user_edit", user_id=user_id)), 400
    with transaction() as db:
        user = db.get(User, user_id)
        if user is None:
            abort(404)
        duplicate = db.scalar(select(User).where(User.email == email, User.id != user.id))
        if duplicate:
            flash("El correo ya esta registrado.", "error")
            return redirect(url_for("auth.admin_user_edit", user_id=user_id)), 409
        previous_roles = sorted(role.name for role in user.roles)
        previous_status = user.status
        user.name = name
        user.email = email
        user.status = status
        user.roles = selected_roles(db, request.form.getlist("roles"))
        if status != "active" or previous_status != status:
            user.session_version += 1
        audit(
            db,
            "admin.user_updated",
            user=current_user,
            entity_type="user",
            entity_id=user.id,
            metadata={"previous_roles": previous_roles, "new_roles": sorted(role.name for role in user.roles), "previous_status": previous_status, "new_status": status},
        )
    flash("Usuario actualizado.", "success")
    return redirect(url_for("auth.admin_user_edit", user_id=user_id))


@auth_bp.post("/admin/usuarios/<user_id>/restablecer-acceso")
@permission_required("usuarios.editar")
def admin_user_reset_access(user_id: str):
    with transaction() as db:
        user = db.get(User, user_id)
        if user is None:
            abort(404)
        token = create_password_token(db, user, "password_reset", timedelta(hours=1))
        emailed = send_token_email(user, token, "password_reset")
        user.session_version += 1
        audit(db, "admin.user_access_reset", user=current_user, entity_type="user", entity_id=user.id, metadata={"email_sent": emailed})
        if current_app.config.get("AUTH_SHOW_DEV_TOKENS"):
            flash(f"Enlace temporal: {urljoin(public_base_url(), 'password/reset/' + token)}", "success")
        else:
            flash("Acceso restablecido. Configura SMTP para enviar el enlace automaticamente." if not emailed else "Acceso restablecido y enlace enviado.", "success")
    return redirect(url_for("auth.admin_user_edit", user_id=user_id))


@auth_bp.get("/admin/roles")
@permission_required("roles.ver")
def admin_roles():
    with transaction() as db:
        return render_template("admin_roles.html", roles=role_choices(db), permissions=permission_choices(db))


@auth_bp.post("/admin/roles/<role_id>")
@permission_required("roles.editar")
def admin_role_update(role_id: str):
    with transaction() as db:
        role = db.get(Role, role_id)
        if role is None:
            abort(404)
        if role.name == "superadmin":
            flash("El rol superadmin conserva todos los permisos.", "error")
            return redirect(url_for("auth.admin_roles")), 400
        permission_ids = [value for value in request.form.getlist("permissions") if re.fullmatch(r"[0-9a-fA-F-]{36}", value or "")]
        role.permissions = db.scalars(select(Permission).where(Permission.id.in_(permission_ids))).all() if permission_ids else []
        audit(db, "admin.role_permissions_updated", user=current_user, entity_type="role", entity_id=role.id, metadata={"role": role.name})
    flash("Permisos actualizados.", "success")
    return redirect(url_for("auth.admin_roles"))


@auth_bp.get("/admin/auditoria")
@permission_required("auditoria.ver")
def admin_audit():
    with transaction() as db:
        logs = db.scalars(select(AuditLog).order_by(AuditLog.created_at.desc()).limit(200)).all()
        return render_template("admin_audit.html", logs=logs)


@auth_bp.errorhandler(403)
def forbidden(_error):
    if request.path.startswith("/api/"):
        return jsonify({"error": "No autorizado."}), 403
    return render_template("error.html", title="No autorizado", message="No tienes permisos para acceder a esta seccion."), 403


def create_superadmin(name: str, email: str, password: str) -> User:
    with transaction() as db:
        seed_rbac(db)
        email = validate_email(email)
        existing = db.scalar(select(User).where(User.email == email))
        if existing:
            raise ValueError("Ya existe un usuario con ese correo.")
        role = db.scalar(select(Role).where(Role.name == "superadmin"))
        user = User(name=validate_name(name), email=email, password_hash=hash_password(password), status="active")
        user.roles = [role] if role else []
        db.add(user)
        db.flush()
        audit(db, "admin.superadmin_created", user=user, entity_type="user", entity_id=user.id)
        return user


def init_auth(app) -> None:
    secure_cookies = os.getenv("SESSION_COOKIE_SECURE", "").lower() in {"1", "true", "yes"}
    if os.getenv("APP_ENV", "development").lower() == "production":
        secure_cookies = True
    secret_key = os.getenv("APP_SECRET_KEY") or os.getenv("SECRET_KEY")
    if os.getenv("APP_ENV", "development").lower() == "production" and not secret_key:
        raise RuntimeError("APP_SECRET_KEY o SECRET_KEY es obligatorio en produccion.")
    if not secret_key:
        secret_key = "dev-only-change-me"
    app.config.update(
        SECRET_KEY=secret_key,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE=os.getenv("SESSION_COOKIE_SAMESITE", "Lax"),
        SESSION_COOKIE_SECURE=secure_cookies,
        PERMANENT_SESSION_LIFETIME=timedelta(hours=int(os.getenv("SESSION_HOURS", "8"))),
        WTF_CSRF_TIME_LIMIT=int(os.getenv("CSRF_TIME_LIMIT_SECONDS", "3600")),
        AUTH_REQUIRED=os.getenv("AUTH_REQUIRED", "true").lower() not in {"0", "false", "no"},
        AUTH_SHOW_DEV_TOKENS=os.getenv("AUTH_SHOW_DEV_TOKENS", "false").lower() in {"1", "true", "yes"},
        RATELIMIT_STORAGE_URI=os.getenv("RATELIMIT_STORAGE_URL", "memory://"),
    )
    trusted_proxies = int(os.getenv("TRUSTED_PROXY_COUNT", "0"))
    if trusted_proxies > 0:
        app.wsgi_app = ProxyFix(app.wsgi_app, x_for=trusted_proxies, x_proto=trusted_proxies, x_host=trusted_proxies)
    configure_database()
    login_manager.login_view = "auth.login"
    login_manager.login_message = "Inicia sesion para continuar."
    login_manager.init_app(app)
    csrf.init_app(app)
    limiter.init_app(app)
    app.register_blueprint(auth_bp)

    @app.teardown_appcontext
    def remove_session(_exc=None):
        SessionLocal.remove()

    @app.after_request
    def security_headers(response):
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
        response.headers.setdefault("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; object-src 'none'; base-uri 'self'; frame-ancestors 'none'")
        if request.is_secure:
            response.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
        return response

    @app.cli.command("init-db")
    def init_db_command():
        create_schema()
        print("Base de datos inicializada.")

    @app.cli.command("create-superadmin")
    def create_superadmin_command():
        import getpass

        name = input("Nombre: ")
        email = input("Correo: ")
        password = getpass.getpass("Contrasena: ")
        confirmation = getpass.getpass("Confirmar contrasena: ")
        if password != confirmation:
            raise SystemExit("Las contrasenas no coinciden.")
        create_superadmin(name, email, password)
        print("Superadministrador creado.")
