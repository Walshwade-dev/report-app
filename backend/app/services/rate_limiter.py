import os
import time
import logging

logger = logging.getLogger(__name__)

# Simple in-memory rate limiter fallback
_login_attempts = {}  # structure: {key: (count, reset_time)}

def _check_in_memory_rate_limit(username: str, client_ip: str) -> bool:
    now = time.time()
    
    # Clean old records
    for k in list(_login_attempts.keys()):
        _, reset_time = _login_attempts[k]
        if now > reset_time:
            _login_attempts.pop(k, None)
            
    for limit, key in [(5, f"user:{username}"), (10, f"ip:{client_ip}")]:
        record = _login_attempts.get(key)
        if record:
            count, reset_time = record
            if now < reset_time:
                if count >= limit:
                    logger.warning("In-memory rate limit hit for key: %s (count: %s/%s)", key, count, limit)
                    return False
                _login_attempts[key] = (count + 1, reset_time)
            else:
                _login_attempts[key] = (1, now + 60)
        else:
            _login_attempts[key] = (1, now + 60)
            
    return True

def check_login_rate_limit(username: str, client_ip: str) -> bool:
    redis_url = os.getenv("REDIS_URL")
    if not redis_url:
        return _check_in_memory_rate_limit(username, client_ip)
        
    try:
        from redis import Redis
        redis_conn = Redis.from_url(redis_url)
        
        # Check username limit (max 5 login attempts per minute per user)
        user_key = f"rate_limit:login:user:{username}"
        user_attempts = redis_conn.get(user_key)
        if user_attempts and int(user_attempts) >= 5:
            logger.warning("Redis rate limit hit for user: %s", username)
            return False
            
        # Check IP limit (max 10 login attempts per minute per IP)
        ip_key = f"rate_limit:login:ip:{client_ip}"
        ip_attempts = redis_conn.get(ip_key)
        if ip_attempts and int(ip_attempts) >= 10:
            logger.warning("Redis rate limit hit for IP: %s", client_ip)
            return False
            
        # Increment counters and set expirations
        pipe = redis_conn.pipeline()
        pipe.incr(user_key)
        pipe.expire(user_key, 60)
        pipe.incr(ip_key)
        pipe.expire(ip_key, 60)
        pipe.execute()
        return True
    except Exception as e:
        logger.warning("Redis connection error in rate limiter, falling back to in-memory: %s", e)
        return _check_in_memory_rate_limit(username, client_ip)
