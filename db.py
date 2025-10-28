# app/db.py
from sqlalchemy import create_engine
from sqlmodel import Session  # <- on utilise la Session de sqlmodel pour garder .exec()
from typing import Generator

# Choisis ton URL (SQLite fichier local)
DATABASE_URL = "sqlite:///./sps.db"

# Pour SQLite + FastAPI (multi-threads) :
engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False},
    future=True,
)

def get_session() -> Generator[Session, None, None]:
    """Dépendance FastAPI : retourne une Session compatible .exec()/.get()"""
    with Session(engine) as session:
        yield session

def init_db() -> None:
    """Crée les tables au démarrage."""
    # import tardif pour éviter l'import circulaire (db -> models -> db)
    from app.models import Base
    Base.metadata.create_all(bind=engine)
