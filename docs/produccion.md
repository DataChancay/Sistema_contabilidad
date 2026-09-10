# Produccion y seguridad

## PostgreSQL

Usa un rol dedicado para la aplicacion. No uses `postgres` ni un superusuario. Revisa `docs/postgresql-produccion.sql` y reemplaza los placeholders.

La aplicacion debe conectarse por `DATABASE_URL` o por `DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`, `DB_PASSWORD`.

Si la aplicacion y PostgreSQL estan en el mismo servidor, prioriza `DB_HOST=127.0.0.1` o socket local. En `postgresql.conf`, evita `listen_addresses='*'` si no es estrictamente necesario. En `pg_hba.conf`, permite solo localhost o la red privada requerida. En el firewall, no publiques `5432` a internet.

## HTTPS y proxy

Termina TLS en Nginx o Apache. Define:

```env
APP_ENV=production
APP_BASE_URL=https://tu-dominio.example
SESSION_COOKIE_SECURE=true
TRUSTED_PROXY_COUNT=1
```

Solo configura `TRUSTED_PROXY_COUNT` con el numero real de proxies controlados por ti.

## Correo

Para activacion y recuperacion de contrasena configura:

```env
SMTP_HOST=
SMTP_PORT=587
SMTP_USER=
SMTP_PASSWORD=
SMTP_FROM=
```

Sin SMTP, el sistema crea tokens de un solo uso, pero no los envia por correo. En desarrollo puedes usar `AUTH_SHOW_DEV_TOKENS=true`; no lo habilites en produccion.

## MFA

La tabla `users` ya incluye `mfa_enabled` y `mfa_enrolled_at`. La siguiente fase debe usar una libreria TOTP mantenida, almacenar secretos protegidos con una clave de aplicacion dedicada y exigir MFA primero a `superadmin` y `admin`. No hay criptografia propia en esta implementacion.

## Migraciones

Ejecuta la migracion SQL con el usuario de la aplicacion despues de crear base de datos y permisos:

```powershell
psql "postgresql://USUARIO_APP:PASSWORD_SEGURA@HOST:5432/NOMBRE_BD?options=-csearch_path%3Dapp" -f migrations/001_auth_rbac.sql
```

Luego crea el primer superadministrador:

```powershell
flask --app app create-superadmin
```
