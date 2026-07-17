"""reset_admin_password

Revision ID: 289f2c6b9b36
Revises: 8f3b77c623b8
Create Date: 2026-07-17 10:53:22.313803

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


import uuid

# revision identifiers, used by Alembic.
revision: str = '289f2c6b9b36'
down_revision: Union[str, None] = '8f3b77c623b8'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    connection = op.get_bind()
    result = connection.execute(
        sa.text("SELECT id FROM users WHERE username = 'admin'")
    ).fetchone()
    
    hashed_pwd = "$2b$12$.Kn.j9ttYRNBg4/f9OZkLeh4bLwbGNrViYx.VYNvkovDX8/xmNxAS"  # Allbegood8*
    
    if result:
        connection.execute(
            sa.text("UPDATE users SET hashed_password = :pwd WHERE username = 'admin'"),
            {"pwd": hashed_pwd}
        )
    else:
        new_id = str(uuid.uuid4())
        connection.execute(
            sa.text(
                "INSERT INTO users (id, username, hashed_password, full_name, role, created_at, updated_at) "
                "VALUES (:id, 'admin', :pwd, 'Local Admin', 'admin', NOW(), NOW())"
            ),
            {"id": new_id, "pwd": hashed_pwd}
        )


def downgrade() -> None:
    pass
