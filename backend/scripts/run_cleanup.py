import logging
from app.services.report_session_store import ReportSessionStore, DEFAULT_SESSION_RETENTION_HOURS

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def main():
    logger.info("Starting storage and database cleanup...")
    store = ReportSessionStore()
    deleted_ids = store.cleanup_expired_sessions(DEFAULT_SESSION_RETENTION_HOURS)
    logger.info(f"Cleanup complete. Deleted {len(deleted_ids)} expired sessions.")

if __name__ == "__main__":
    main()
