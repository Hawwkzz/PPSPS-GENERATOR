from __future__ import annotations
from datetime import datetime, date, timezone
from typing import Optional
from sqlalchemy import ForeignKey
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship

class Base(DeclarativeBase):
    pass

class UserDB(Base):
    __tablename__ = "userdb"

    id: Mapped[int] = mapped_column(primary_key=True)
    email: Mapped[str] = mapped_column(unique=True, index=True)
    password_hash: Mapped[str]
    is_admin: Mapped[bool] = mapped_column(default=False)
    created_at: Mapped[datetime] = mapped_column(default=lambda: datetime.now(timezone.utc))  # <- CHANGÃ‰

    projects: Mapped[list["ProjectDB"]] = relationship(back_populates="owner")

class ProjectDB(Base):
    __tablename__ = "projectdb"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str]
    address: Mapped[str]
    works_csv: Mapped[str] = mapped_column(default="")
    duration_weeks: Mapped[int] = mapped_column(default=0)
    workforce: Mapped[int] = mapped_column(default=0)
    companies_csv: Mapped[str] = mapped_column(default="")

    owner_id: Mapped[int] = mapped_column(ForeignKey("userdb.id"), index=True)
    owner: Mapped[Optional["UserDB"]] = relationship(back_populates="projects")

    site_env: Mapped[str] = mapped_column(default="")
    zones_csv: Mapped[str] = mapped_column(default="")
    eu_name: Mapped[str] = mapped_column(default="")
    csps_name: Mapped[str] = mapped_column(default="")
    work_hours: Mapped[str] = mapped_column(default="")
    start_date: Mapped[Optional[date]] = mapped_column(nullable=True, default=None)
    end_date: Mapped[Optional[date]] = mapped_column(nullable=True, default=None)
    facts_json: Mapped[Optional[str]] = mapped_column(nullable=True, default=None)

    created_at: Mapped[datetime] = mapped_column(default=lambda: datetime.now(timezone.utc))   # <- CHANGÃ‰
    updated_at: Mapped[datetime] = mapped_column(default=lambda: datetime.now(timezone.utc))   # <- CHANGÃ‰

class DocumentDB(Base):
    __tablename__ = "documentdb"

    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projectdb.id"), index=True)
    doc_type: Mapped[str]
    content_md: Mapped[str]
    version: Mapped[int] = mapped_column(default=1)
    status: Mapped[str] = mapped_column(default="draft")
    created_at: Mapped[datetime] = mapped_column(default=lambda: datetime.now(timezone.utc))   # <- CHANGÃ‰
    updated_at: Mapped[datetime] = mapped_column(default=lambda: datetime.now(timezone.utc))   # <- CHANGÃ‰

class AttachmentDB(Base):
    __tablename__ = "attachmentdb"

    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projectdb.id"), index=True)
    filename: Mapped[str]
    stored_path: Mapped[str]
    mime_type: Mapped[str]
    size_bytes: Mapped[int]
    extracted_text: Mapped[Optional[str]] = mapped_column(nullable=True, default=None)
    created_at: Mapped[datetime] = mapped_column(default=lambda: datetime.now(timezone.utc))   # <- CHANGÃ‰