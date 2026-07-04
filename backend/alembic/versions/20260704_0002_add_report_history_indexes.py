"""add report history indexes

Revision ID: 20260704_0002
Revises: 20260704_0001
Create Date: 2026-07-04 00:00:00.000000
"""

from alembic import op


revision = "20260704_0002"
down_revision = "20260704_0001"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_index("ix_reports_report_type", "reports", ["report_type"])
    op.create_index("ix_reports_created_at", "reports", ["created_at"])


def downgrade() -> None:
    op.drop_index("ix_reports_created_at", table_name="reports")
    op.drop_index("ix_reports_report_type", table_name="reports")
