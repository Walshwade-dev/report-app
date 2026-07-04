"""create report metadata tables

Revision ID: 20260704_0001
Revises:
Create Date: 2026-07-04 00:00:00.000000
"""

from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import postgresql


revision = "20260704_0001"
down_revision = None
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "reports",
        sa.Column("id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("title", sa.String(length=255), nullable=True),
        sa.Column("report_type", sa.String(length=80), nullable=False),
        sa.Column("report_date", sa.String(length=32), nullable=True),
        sa.Column("station", sa.String(length=255), nullable=True),
        sa.Column("weighbridge_name", sa.String(length=255), nullable=True),
        sa.Column("bound_name", sa.String(length=255), nullable=True),
        sa.Column("status", sa.String(length=80), nullable=False),
        sa.Column("state_payload", postgresql.JSONB(astext_type=sa.Text()), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("completed_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("error_message", sa.Text(), nullable=True),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_reports_report_date", "reports", ["report_date"])
    op.create_index("ix_reports_status", "reports", ["status"])

    op.create_table(
        "report_uploads",
        sa.Column("id", sa.BigInteger(), autoincrement=True, nullable=False),
        sa.Column("report_id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("upload_type", sa.String(length=80), nullable=False),
        sa.Column("original_filename", sa.Text(), nullable=True),
        sa.Column("stored_filename", sa.Text(), nullable=True),
        sa.Column("file_path", sa.Text(), nullable=False),
        sa.Column("file_size_bytes", sa.BigInteger(), nullable=True),
        sa.Column("mime_type", sa.String(length=255), nullable=True),
        sa.Column("uploaded_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["report_id"], ["reports.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("report_id", "upload_type", name="uq_report_uploads_report_type"),
    )
    op.create_index("ix_report_uploads_report_id", "report_uploads", ["report_id"])

    op.create_table(
        "report_manual_inputs",
        sa.Column("id", sa.BigInteger(), autoincrement=True, nullable=False),
        sa.Column("report_id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("prepared_by", sa.String(length=255), nullable=True),
        sa.Column("approved_by", sa.String(length=255), nullable=True),
        sa.Column("weighbridge_name", sa.String(length=255), nullable=True),
        sa.Column("bound_name", sa.String(length=255), nullable=True),
        sa.Column("cases_cleared_in_court", sa.Integer(), nullable=False),
        sa.Column("total_transgressions", sa.Integer(), nullable=False),
        sa.Column("buses_3500kg", sa.Integer(), nullable=False),
        sa.Column("vehicles_3500_to_7000kg", sa.Integer(), nullable=False),
        sa.Column("vehicles_above_7000kg", sa.Integer(), nullable=False),
        sa.Column("traffic_total", sa.Integer(), nullable=False),
        sa.Column("payload", postgresql.JSONB(astext_type=sa.Text()), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["report_id"], ["reports.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("report_id"),
    )

    op.create_table(
        "report_previews",
        sa.Column("id", sa.BigInteger(), autoincrement=True, nullable=False),
        sa.Column("report_id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("section_number", sa.Integer(), nullable=True),
        sa.Column("section_name", sa.String(length=120), nullable=False),
        sa.Column("format", sa.String(length=24), nullable=False),
        sa.Column("page", sa.Integer(), nullable=True),
        sa.Column("file_path", sa.Text(), nullable=False),
        sa.Column("generated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["report_id"], ["reports.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint(
            "report_id",
            "section_name",
            "format",
            "page",
            name="uq_report_previews_report_section_format_page",
        ),
    )
    op.create_index("ix_report_previews_report_id", "report_previews", ["report_id"])

    op.create_table(
        "report_outputs",
        sa.Column("id", sa.BigInteger(), autoincrement=True, nullable=False),
        sa.Column("report_id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("final_docx_path", sa.Text(), nullable=True),
        sa.Column("final_pdf_path", sa.Text(), nullable=True),
        sa.Column("generated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["report_id"], ["reports.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("report_id"),
    )
    op.create_index("ix_report_outputs_report_id", "report_outputs", ["report_id"])


def downgrade() -> None:
    op.drop_index("ix_report_outputs_report_id", table_name="report_outputs")
    op.drop_table("report_outputs")
    op.drop_index("ix_report_previews_report_id", table_name="report_previews")
    op.drop_table("report_previews")
    op.drop_table("report_manual_inputs")
    op.drop_index("ix_report_uploads_report_id", table_name="report_uploads")
    op.drop_table("report_uploads")
    op.drop_index("ix_reports_status", table_name="reports")
    op.drop_index("ix_reports_report_date", table_name="reports")
    op.drop_table("reports")
