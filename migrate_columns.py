from app.db import engine
from sqlalchemy import text

with engine.connect() as conn:
    conn.execute(text("""
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS project_reference VARCHAR DEFAULT '';
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS site_phone VARCHAR DEFAULT '';
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS company_name VARCHAR DEFAULT '';
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS company_address VARCHAR DEFAULT '';
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS company_phone VARCHAR DEFAULT '';
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS company_email VARCHAR DEFAULT '';
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS site_manager_name VARCHAR DEFAULT '';
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS owner_name VARCHAR DEFAULT '';
        ALTER TABLE projectdb ADD COLUMN IF NOT EXISTS architect_name VARCHAR DEFAULT '';
    """))
    conn.commit()

print("✅ Colonnes ajoutées")
