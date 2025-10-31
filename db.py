# app/db.py
import os
from sqlalchemy import create_engine
from sqlmodel import Session
from typing import Generator

# 🔧 CONFIGURATION AUTOMATIQUE : PostgreSQL (Railway) ou SQLite (local)
DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    # Fallback pour développement local avec SQLite
    DATABASE_URL = "sqlite:///./sps.db"
    print("⚠️ Utilisation de SQLite (développement local)")
    
    engine = create_engine(
        DATABASE_URL,
        connect_args={"check_same_thread": False},  # Requis pour SQLite
        future=True,
    )
else:
    # Production avec PostgreSQL (Railway)
    print(f"✅ Connexion à PostgreSQL")
    
    # Railway fournit parfois postgres:// au lieu de postgresql://
    # SQLAlchemy 1.4+ requiert postgresql://
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
    
    engine = create_engine(
        DATABASE_URL,
        pool_pre_ping=True,  # Vérifie la connexion avant utilisation
        pool_size=10,        # Pool de connexions
        max_overflow=20,     # Connexions supplémentaires si besoin
        future=True,
    )

def get_session() -> Generator[Session, None, None]:
    """Dépendance FastAPI : retourne une Session compatible .exec()/.get()"""
    with Session(engine) as session:
        yield session

def init_db() -> None:
    """Crée les tables au démarrage."""
    from app.models import Base
    Base.metadata.create_all(bind=engine)
    print("✅ Tables créées/vérifiées")
