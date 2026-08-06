import os
import logging
import asyncio
from concurrent.futures import ProcessPoolExecutor
from fastapi import BackgroundTasks
from app.services.report_session_store import report_session_store
from app.services.report_session_metrics import get_wideload_count_from_session
from app.services.final_report_builder import build_final_report

logger = logging.getLogger(__name__)

_process_pool = ProcessPoolExecutor(max_workers=2)

def _build_final_report_in_process(report_id: str) -> tuple[bytes | None, str | None]:
    logger.info("Starting background build for report session %s in separate process", report_id)
    try:
        session = report_session_store.get(report_id)
        if not session:
            logger.error("Session %s not found", report_id)
            return None, f"Session {report_id} not found"

        wideload_count = get_wideload_count_from_session(session)
        if wideload_count is None:
            wideload_count = len(session.dataframes.get("wideload", []))

        file_stream = build_final_report(
            daily_df=session.dataframes["daily_hour"],
            wideload_df=session.dataframes["wideload"],
            impounded_prohibited_df=session.dataframes["impounded_prohibited"],
            overloaded_df=session.dataframes["overloaded"],
            report_date=session.report_date,
            station=session.station,
            bound=session.bound,
            prepared_by=session.prepared_by,
            confirmed_by=session.confirmed_by,
            traffic_census=session.manual_inputs.get("traffic_census"),
            daily_summary=session.sections.get("daily_summary", {}).get("values"),
            transgressions=session.manual_inputs.get("transgressions"),
            wideload_count=wideload_count,
        )
        return file_stream.read(), None
    except Exception as exc:
        logger.exception("Failed background build for report session %s in separate process", report_id)
        return None, str(exc)


def run_build_final_report(report_id: str):
    logger.info("Starting background build for report session %s", report_id)
    content, error = _build_final_report_in_process(report_id)
    if error:
        report_session_store.set_final_report_error(report_id, error)
    elif content:
        report_session_store.set_final_report(report_id, content)
        logger.info("Successfully completed background build for report session %s", report_id)


async def _async_build_and_save(report_id: str):
    loop = asyncio.get_running_loop()
    content, error = await loop.run_in_executor(_process_pool, _build_final_report_in_process, report_id)
    if error:
        report_session_store.set_final_report_error(report_id, error)
    elif content:
        report_session_store.set_final_report(report_id, content)
        logger.info("Successfully completed async background build for report session %s", report_id)


def enqueue_build_final_report(report_id: str, background_tasks: BackgroundTasks):
    redis_url = os.getenv("REDIS_URL")
    if redis_url:
        try:
            from redis import Redis
            from rq import Queue
            redis_conn = Redis.from_url(redis_url)
            q = Queue("reports", connection=redis_conn)
            q.enqueue(run_build_final_report, report_id)
            logger.info("Enqueued report build task in Redis Queue for report_id: %s", report_id)
            return
        except Exception as e:
            logger.warning("Failed to enqueue task in Redis Queue, falling back to background tasks: %s", e)

    # Fallback to local async ProcessPoolExecutor
    background_tasks.add_task(_async_build_and_save, report_id)
    logger.info("Enqueued report build task in ProcessPoolExecutor for report_id: %s", report_id)
