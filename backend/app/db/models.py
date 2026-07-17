from datetime import datetime, timezone
from uuid import uuid4

from sqlalchemy import (
    BigInteger,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.dialects.postgresql import JSONB, UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.core.database import Base


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


class Report(Base):
    __tablename__ = "reports"

    id: Mapped[UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid4)
    title: Mapped[str | None] = mapped_column(String(255), nullable=True)
    report_type: Mapped[str] = mapped_column(String(80), nullable=False, default="static_weighbridge")
    report_date: Mapped[str | None] = mapped_column(String(32), nullable=True)
    station: Mapped[str | None] = mapped_column(String(255), nullable=True)
    weighbridge_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    bound_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    status: Mapped[str] = mapped_column(String(80), nullable=False, default="draft")
    state_payload: Mapped[dict | None] = mapped_column(JSONB, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        default=utc_now,
        onupdate=utc_now,
    )
    completed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    error_message: Mapped[str | None] = mapped_column(Text, nullable=True)

    uploads: Mapped[list["ReportUpload"]] = relationship(
        back_populates="report",
        cascade="all, delete-orphan",
    )
    manual_input: Mapped["ReportManualInput | None"] = relationship(
        back_populates="report",
        cascade="all, delete-orphan",
        uselist=False,
    )
    previews: Mapped[list["ReportPreview"]] = relationship(
        back_populates="report",
        cascade="all, delete-orphan",
    )
    output: Mapped["ReportOutput | None"] = relationship(
        back_populates="report",
        cascade="all, delete-orphan",
        uselist=False,
    )

    __table_args__ = (
        Index("ix_reports_status", "status"),
        Index("ix_reports_report_type", "report_type"),
        Index("ix_reports_created_at", "created_at"),
        Index("ix_reports_report_date", "report_date"),
    )


class ReportUpload(Base):
    __tablename__ = "report_uploads"

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    report_id: Mapped[UUID] = mapped_column(
        UUID(as_uuid=True),
        ForeignKey("reports.id", ondelete="CASCADE"),
        nullable=False,
    )
    upload_type: Mapped[str] = mapped_column(String(80), nullable=False)
    original_filename: Mapped[str | None] = mapped_column(Text, nullable=True)
    stored_filename: Mapped[str | None] = mapped_column(Text, nullable=True)
    file_path: Mapped[str] = mapped_column(Text, nullable=False)
    file_size_bytes: Mapped[int | None] = mapped_column(BigInteger, nullable=True)
    mime_type: Mapped[str | None] = mapped_column(String(255), nullable=True)
    uploaded_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    report: Mapped[Report] = relationship(back_populates="uploads")

    __table_args__ = (
        UniqueConstraint("report_id", "upload_type", name="uq_report_uploads_report_type"),
        Index("ix_report_uploads_report_id", "report_id"),
    )


class ReportManualInput(Base):
    __tablename__ = "report_manual_inputs"

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    report_id: Mapped[UUID] = mapped_column(
        UUID(as_uuid=True),
        ForeignKey("reports.id", ondelete="CASCADE"),
        nullable=False,
        unique=True,
    )
    prepared_by: Mapped[str | None] = mapped_column(String(255), nullable=True)
    approved_by: Mapped[str | None] = mapped_column(String(255), nullable=True)
    weighbridge_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    bound_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    cases_cleared_in_court: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    total_transgressions: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    buses_3500kg: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    vehicles_3500_to_7000kg: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    vehicles_above_7000kg: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    traffic_total: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    payload: Mapped[dict | None] = mapped_column(JSONB, nullable=True)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        default=utc_now,
        onupdate=utc_now,
    )

    report: Mapped[Report] = relationship(back_populates="manual_input")


class ReportPreview(Base):
    __tablename__ = "report_previews"

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    report_id: Mapped[UUID] = mapped_column(
        UUID(as_uuid=True),
        ForeignKey("reports.id", ondelete="CASCADE"),
        nullable=False,
    )
    section_number: Mapped[int | None] = mapped_column(Integer, nullable=True)
    section_name: Mapped[str] = mapped_column(String(120), nullable=False)
    format: Mapped[str] = mapped_column(String(24), nullable=False)
    page: Mapped[int | None] = mapped_column(Integer, nullable=True)
    file_path: Mapped[str] = mapped_column(Text, nullable=False)
    generated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    report: Mapped[Report] = relationship(back_populates="previews")

    __table_args__ = (
        UniqueConstraint(
            "report_id",
            "section_name",
            "format",
            "page",
            name="uq_report_previews_report_section_format_page",
        ),
        Index("ix_report_previews_report_id", "report_id"),
    )


class ReportOutput(Base):
    __tablename__ = "report_outputs"

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    report_id: Mapped[UUID] = mapped_column(
        UUID(as_uuid=True),
        ForeignKey("reports.id", ondelete="CASCADE"),
        nullable=False,
        unique=True,
    )
    final_docx_path: Mapped[str | None] = mapped_column(Text, nullable=True)
    final_pdf_path: Mapped[str | None] = mapped_column(Text, nullable=True)
    generated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    report: Mapped[Report] = relationship(back_populates="output")

    __table_args__ = (Index("ix_report_outputs_report_id", "report_id"),)


class User(Base):
    __tablename__ = "users"

    id: Mapped[UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid4)
    username: Mapped[str] = mapped_column(String(80), nullable=False, unique=True)
    hashed_password: Mapped[str] = mapped_column(String(255), nullable=False)
    full_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    role: Mapped[str] = mapped_column(String(32), nullable=False, default="user")
    station: Mapped[str | None] = mapped_column(String(255), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        default=utc_now,
        onupdate=utc_now,
    )

    __table_args__ = (
        Index("ix_users_username", "username"),
    )

