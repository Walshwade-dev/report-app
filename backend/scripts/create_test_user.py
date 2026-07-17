import sys
import os

# Add backend root to python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv
load_dotenv()

from app.core.database import SessionLocal
from app.db.models import User
from app.core.security import get_password_hash

def create_user():
    if SessionLocal is None:
        print("Error: Database not configured.")
        return

    admin_username = os.getenv("ADMIN_USERNAME", "admin")
    admin_password = os.getenv("ADMIN_PASSWORD", "Allbegood8*")

    db = SessionLocal()
    try:
        # Check if user exists
        existing = db.query(User).filter(User.username == admin_username).first()
        if existing:
            existing.hashed_password = get_password_hash(admin_password)
            db.commit()
            print(f"Updated existing user '{admin_username}' with password '{admin_password}'.")
            return
        
        user = User(
            username=admin_username,
            hashed_password=get_password_hash(admin_password),
            full_name="Local Admin",
            role="admin"
        )
        db.add(user)
        db.commit()
        print(f"Created test user '{admin_username}' with password '{admin_password}'.")
    finally:
        db.close()

if __name__ == "__main__":
    create_user()
