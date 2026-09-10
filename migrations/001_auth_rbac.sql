CREATE TABLE IF NOT EXISTS users (
    id uuid PRIMARY KEY,
    name varchar(120) NOT NULL,
    email varchar(255) NOT NULL,
    password_hash text,
    status varchar(20) NOT NULL DEFAULT 'pending',
    last_login_at timestamptz,
    session_version integer NOT NULL DEFAULT 0,
    mfa_enabled boolean NOT NULL DEFAULT false,
    mfa_enrolled_at timestamptz,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT uq_users_email UNIQUE (email),
    CONSTRAINT ck_users_status CHECK (status IN ('pending', 'active', 'inactive', 'blocked'))
);

CREATE INDEX IF NOT EXISTS ix_users_status ON users(status);

CREATE TABLE IF NOT EXISTS roles (
    id uuid PRIMARY KEY,
    name varchar(80) NOT NULL,
    description varchar(255),
    is_system boolean NOT NULL DEFAULT true,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT uq_roles_name UNIQUE (name)
);

CREATE TABLE IF NOT EXISTS permissions (
    id uuid PRIMARY KEY,
    code varchar(120) NOT NULL,
    description varchar(255),
    created_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT uq_permissions_code UNIQUE (code)
);

CREATE TABLE IF NOT EXISTS user_roles (
    user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    role_id uuid NOT NULL REFERENCES roles(id) ON DELETE CASCADE,
    created_at timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (user_id, role_id)
);

CREATE TABLE IF NOT EXISTS role_permissions (
    role_id uuid NOT NULL REFERENCES roles(id) ON DELETE CASCADE,
    permission_id uuid NOT NULL REFERENCES permissions(id) ON DELETE CASCADE,
    created_at timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (role_id, permission_id)
);

CREATE TABLE IF NOT EXISTS password_tokens (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    token_hash varchar(128) NOT NULL,
    purpose varchar(30) NOT NULL,
    expires_at timestamptz NOT NULL,
    used_at timestamptz,
    created_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT uq_password_tokens_token_hash UNIQUE (token_hash),
    CONSTRAINT ck_password_tokens_purpose CHECK (purpose IN ('activation', 'password_reset'))
);

CREATE INDEX IF NOT EXISTS ix_password_tokens_user_id ON password_tokens(user_id);
CREATE INDEX IF NOT EXISTS ix_password_tokens_expires_at ON password_tokens(expires_at);

CREATE TABLE IF NOT EXISTS audit_logs (
    id uuid PRIMARY KEY,
    user_id uuid REFERENCES users(id) ON DELETE SET NULL,
    action varchar(120) NOT NULL,
    entity_type varchar(80),
    entity_id varchar(80),
    ip_address varchar(80),
    user_agent varchar(300),
    metadata jsonb,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_audit_logs_user_id ON audit_logs(user_id);
CREATE INDEX IF NOT EXISTS ix_audit_logs_created_at ON audit_logs(created_at);
CREATE INDEX IF NOT EXISTS ix_audit_logs_action ON audit_logs(action);

INSERT INTO permissions (id, code, description)
VALUES
    ('10000000-0000-0000-0000-000000000001', 'usuarios.ver', 'Ver usuarios'),
    ('10000000-0000-0000-0000-000000000002', 'usuarios.crear', 'Crear usuarios'),
    ('10000000-0000-0000-0000-000000000003', 'usuarios.editar', 'Editar usuarios'),
    ('10000000-0000-0000-0000-000000000004', 'usuarios.desactivar', 'Desactivar usuarios'),
    ('10000000-0000-0000-0000-000000000005', 'roles.ver', 'Ver roles'),
    ('10000000-0000-0000-0000-000000000006', 'roles.editar', 'Editar roles'),
    ('10000000-0000-0000-0000-000000000007', 'reportes.ver', 'Ver reportes'),
    ('10000000-0000-0000-0000-000000000008', 'reportes.exportar', 'Exportar reportes'),
    ('10000000-0000-0000-0000-000000000009', 'configuracion.ver', 'Ver configuracion'),
    ('10000000-0000-0000-0000-000000000010', 'configuracion.editar', 'Editar configuracion'),
    ('10000000-0000-0000-0000-000000000011', 'auditoria.ver', 'Ver auditoria')
ON CONFLICT (code) DO NOTHING;

INSERT INTO roles (id, name, description, is_system)
VALUES
    ('20000000-0000-0000-0000-000000000001', 'superadmin', 'Acceso administrativo completo', true),
    ('20000000-0000-0000-0000-000000000002', 'admin', 'Administracion operativa', true),
    ('20000000-0000-0000-0000-000000000003', 'supervisor', 'Supervision de reportes', true),
    ('20000000-0000-0000-0000-000000000004', 'user', 'Usuario de reportes', true)
ON CONFLICT (name) DO NOTHING;

INSERT INTO role_permissions (role_id, permission_id)
SELECT r.id, p.id
FROM roles r
CROSS JOIN permissions p
WHERE r.name = 'superadmin'
ON CONFLICT DO NOTHING;

INSERT INTO role_permissions (role_id, permission_id)
SELECT r.id, p.id
FROM roles r
JOIN permissions p ON p.code IN (
    'usuarios.ver',
    'usuarios.crear',
    'usuarios.editar',
    'usuarios.desactivar',
    'roles.ver',
    'reportes.ver',
    'reportes.exportar',
    'auditoria.ver'
)
WHERE r.name = 'admin'
ON CONFLICT DO NOTHING;

INSERT INTO role_permissions (role_id, permission_id)
SELECT r.id, p.id
FROM roles r
JOIN permissions p ON p.code IN ('reportes.ver', 'reportes.exportar', 'auditoria.ver')
WHERE r.name = 'supervisor'
ON CONFLICT DO NOTHING;

INSERT INTO role_permissions (role_id, permission_id)
SELECT r.id, p.id
FROM roles r
JOIN permissions p ON p.code IN ('reportes.ver', 'reportes.exportar')
WHERE r.name = 'user'
ON CONFLICT DO NOTHING;
