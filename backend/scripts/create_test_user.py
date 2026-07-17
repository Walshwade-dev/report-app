import sys
import os

# Add backend root to python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.core.database import SessionLocal
from app.db.models import User
from app.core.security import get_password_hash

def create_user():
    if SessionLocal is None:
        print("Error: Database not configured.")
        return

    db = SessionLocal()
    try:
        # Check if user exists
        existing = db.query(User).filter(User.username == "admin").first()
        if existing:
            existing.hashed_password = get_password_hash("Allbegood8*")
            db.commit()
            print("Updated existing user 'admin' with password 'Allbegood8*'.")
            return
        
        user = User(
            username="admin",
            hashed_password=get_password_hash("Allbegood8*"),
            full_name="Local Admin",
            role="admin"
        )
        db.add(user)
        db.commit()
        print("Created test user 'admin' with password 'Allbegood8*'.")
    finally:
        db.close()

if __name__ == "__main__":
    create_user()
