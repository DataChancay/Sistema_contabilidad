import re
import unittest
from datetime import timedelta
from unittest.mock import patch

from sqlalchemy import select

from app import app
from auth import (
    AuditLog,
    PasswordToken,
    Permission,
    Role,
    User,
    configure_database,
    create_password_token,
    create_schema,
    create_superadmin,
    db_session,
    hash_password,
    limiter,
    transaction,
)


PASSWORD = "valid-password-123"


class AuthRoutesTestCase(unittest.TestCase):
    def setUp(self):
        app.config.update(
            TESTING=True,
            AUTH_REQUIRED=True,
            WTF_CSRF_ENABLED=True,
            AUTH_SHOW_DEV_TOKENS=True,
            RATELIMIT_ENABLED=False,
        )
        limiter.reset()
        configure_database("sqlite:///:memory:")
        create_schema()
        self.client = app.test_client()

    def csrf_token(self, path="/login"):
        response = self.client.get(path)
        match = re.search(rb'name="csrf_token" value="([^"]+)"', response.data)
        self.assertIsNotNone(match, response.data.decode("utf-8", errors="ignore"))
        return match.group(1).decode()

    def create_user(self, email="user@example.com", status="active", roles=("user",)):
        with transaction() as db:
            user = User(name="Usuario Prueba", email=email, password_hash=hash_password(PASSWORD), status=status)
            if roles:
                user.roles = db.scalars(select(Role).where(Role.name.in_(roles))).all()
            db.add(user)
            db.flush()
            return user.id

    def login_as(self, email, password=PASSWORD):
        token = self.csrf_token("/login")
        return self.client.post(
            "/login",
            data={"email": email, "password": password, "csrf_token": token, "next": "/"},
            follow_redirects=False,
        )

    def test_visitor_cannot_access_private_system(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/login", response.headers["Location"])

    def test_valid_user_can_login(self):
        self.create_user()
        response = self.login_as("user@example.com")
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers["Location"], "/")

    def test_wrong_password_is_rejected(self):
        self.create_user()
        response = self.login_as("user@example.com", "wrong-password")
        self.assertEqual(response.status_code, 401)
        self.assertIn(b"Credenciales incorrectas", response.data)

    def test_inactive_user_cannot_login(self):
        self.create_user(status="inactive")
        response = self.login_as("user@example.com")
        self.assertEqual(response.status_code, 401)

    def test_blocked_user_cannot_login(self):
        self.create_user(status="blocked")
        response = self.login_as("user@example.com")
        self.assertEqual(response.status_code, 401)

    def test_normal_user_cannot_access_admin(self):
        self.create_user()
        self.login_as("user@example.com")
        response = self.client.get("/admin")
        self.assertEqual(response.status_code, 403)

    def test_admin_access_is_limited_by_permissions(self):
        self.create_user(email="admin@example.com", roles=("admin",))
        self.login_as("admin@example.com")
        self.assertEqual(self.client.get("/admin/roles").status_code, 200)
        role_id = db_session().scalar(select(Role).where(Role.name == "user")).id
        token = self.csrf_token("/admin/roles")
        response = self.client.post(f"/admin/roles/{role_id}", data={"csrf_token": token})
        self.assertEqual(response.status_code, 403)

    def test_superadmin_can_manage_users(self):
        create_superadmin("Super Admin", "super@example.com", PASSWORD)
        self.login_as("super@example.com")
        response = self.client.get("/admin/usuarios")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Usuarios", response.data)

    def test_user_without_permission_cannot_create_users_directly(self):
        self.create_user()
        self.login_as("user@example.com")
        token = self.csrf_token("/")
        response = self.client.post(
            "/admin/usuarios",
            data={"name": "Nuevo", "email": "nuevo@example.com", "csrf_token": token},
        )
        self.assertEqual(response.status_code, 403)

    def test_public_register_route_does_not_exist(self):
        response = self.client.get("/register")
        self.assertEqual(response.status_code, 404)

    def test_rate_limiting_rejects_repeated_login_attempts(self):
        app.config["RATELIMIT_ENABLED"] = True
        self.create_user()
        status_codes = []
        for _ in range(6):
            token = self.csrf_token("/login")
            response = self.client.post(
                "/login",
                data={"email": "user@example.com", "password": "bad", "csrf_token": token},
            )
            status_codes.append(response.status_code)
        self.assertIn(429, status_codes)

    def test_logout_invalidates_session(self):
        self.create_user()
        self.login_as("user@example.com")
        token = self.csrf_token("/")
        self.client.post("/logout", data={"csrf_token": token})
        response = self.client.get("/")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/login", response.headers["Location"])

    def test_password_reset_token_expires(self):
        user_id = self.create_user()
        with transaction() as db:
            user = db.get(User, user_id)
            token = create_password_token(db, user, "password_reset", timedelta(seconds=-1))
        csrf = self.csrf_token(f"/password/reset/{token}")
        response = self.client.post(
            f"/password/reset/{token}",
            data={"password": PASSWORD, "password_confirmation": PASSWORD, "csrf_token": csrf},
        )
        self.assertEqual(response.status_code, 400)

    def test_password_reset_token_cannot_be_reused(self):
        user_id = self.create_user()
        with transaction() as db:
            user = db.get(User, user_id)
            token = create_password_token(db, user, "password_reset")
        csrf = self.csrf_token(f"/password/reset/{token}")
        first = self.client.post(
            f"/password/reset/{token}",
            data={"password": PASSWORD, "password_confirmation": PASSWORD, "csrf_token": csrf},
        )
        self.assertEqual(first.status_code, 302)
        csrf = self.csrf_token(f"/password/reset/{token}")
        second = self.client.post(
            f"/password/reset/{token}",
            data={"password": PASSWORD, "password_confirmation": PASSWORD, "csrf_token": csrf},
        )
        self.assertEqual(second.status_code, 400)

    def test_duplicate_email_is_rejected(self):
        create_superadmin("Super Admin", "super@example.com", PASSWORD)
        self.create_user(email="user@example.com")
        self.login_as("super@example.com")
        csrf = self.csrf_token("/admin/usuarios")
        response = self.client.post(
            "/admin/usuarios",
            data={"name": "Duplicado", "email": "user@example.com", "roles": [], "csrf_token": csrf},
        )
        self.assertEqual(response.status_code, 409)

    def test_role_change_is_audited(self):
        create_superadmin("Super Admin", "super@example.com", PASSWORD)
        user_id = self.create_user(email="target@example.com")
        self.login_as("super@example.com")
        with transaction() as db:
            admin_role = db.scalar(select(Role).where(Role.name == "admin"))
        csrf = self.csrf_token(f"/admin/usuarios/{user_id}")
        self.client.post(
            f"/admin/usuarios/{user_id}",
            data={
                "name": "Target User",
                "email": "target@example.com",
                "status": "active",
                "roles": [admin_role.id],
                "csrf_token": csrf,
            },
        )
        actions = [log.action for log in db_session().scalars(select(AuditLog)).all()]
        self.assertIn("admin.user_updated", actions)

    def test_deactivated_user_cannot_open_new_session(self):
        user_id = self.create_user()
        with transaction() as db:
            user = db.get(User, user_id)
            user.status = "inactive"
            user.session_version += 1
        response = self.login_as("user@example.com")
        self.assertEqual(response.status_code, 401)

    @patch("app.download_xafiro_report")
    def test_backend_authorization_protects_report_endpoint(self, download_report):
        app.config["WTF_CSRF_ENABLED"] = False
        self.create_user(email="limited@example.com", roles=())
        self.login_as("limited@example.com")
        response = self.client.post(
            "/api/xafiro/export",
            json={"fecha_inicio": "2026-08-01", "fecha_fin": "2026-08-02"},
        )
        self.assertEqual(response.status_code, 403)
        download_report.assert_not_called()


if __name__ == "__main__":
    unittest.main()
