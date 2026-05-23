from __future__ import annotations

import logging
import time
from typing import Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from tenacity import Retrying, retry_if_exception, retry_if_exception_type, stop_after_attempt, wait_exponential


RETRYABLE_STATUS_CODES = {408, 429, 500, 502, 503, 504}


class EmptyResponseError(RuntimeError):
    """Raised when an HTTP response body is unexpectedly empty."""


class RetryableHttpError(requests.HTTPError):
    """HTTP error class used to force tenacity retries for transient statuses."""


def _is_retryable_http_error(exc: BaseException) -> bool:
    if not isinstance(exc, requests.HTTPError):
        return False
    response = getattr(exc, "response", None)
    if response is None:
        return False
    return response.status_code in RETRYABLE_STATUS_CODES


def create_retry_session(
    *,
    total_retries: int = 5,
    backoff_factor: float = 1.0,
    pool_maxsize: int = 10,
) -> requests.Session:
    session = requests.Session()

    retry = Retry(
        total=total_retries,
        connect=total_retries,
        read=total_retries,
        status=total_retries,
        allowed_methods=frozenset({"GET", "HEAD"}),
        status_forcelist=sorted(RETRYABLE_STATUS_CODES),
        backoff_factor=backoff_factor,
        raise_on_status=False,
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=pool_maxsize, pool_maxsize=pool_maxsize)

    session.mount("https://", adapter)
    session.mount("http://", adapter)

    return session


def fetch_text(
    session: requests.Session,
    url: str,
    *,
    headers: Optional[dict] = None,
    delay: float = 0.0,
    timeout: int = 90,
    logger: Optional[logging.Logger] = None,
    context: str = "",
    max_attempts: int = 5,
    validate_html: bool = True,
) -> Optional[str]:
    if delay > 0:
        time.sleep(delay)

    def before_sleep(retry_state) -> None:
        if logger is None:
            return
        exc = retry_state.outcome.exception()
        logger.warning(
            "Retry %s/%s for %s | context=%s | reason=%s",
            retry_state.attempt_number,
            max_attempts,
            url,
            context or "request",
            exc,
        )

    def request_once() -> str:
        response = session.get(url, headers=headers, timeout=timeout)
        if response.status_code in RETRYABLE_STATUS_CODES:
            raise RetryableHttpError(
                f"Transient HTTP status {response.status_code} for {url}",
                response=response,
            )
        response.raise_for_status()
        apparent_encoding = getattr(response, "apparent_encoding", None)
        response_encoding = (response.encoding or "").lower()
        if apparent_encoding and response_encoding in {"", "iso-8859-1", "latin-1", "ascii"}:
            response.encoding = apparent_encoding
        text = response.text or ""
        if validate_html and not text.strip():
            raise EmptyResponseError(f"Empty response body for {url}")
        return text

    retrying = Retrying(
        stop=stop_after_attempt(max_attempts),
        wait=wait_exponential(multiplier=1, min=2, max=60),
        retry=(
            retry_if_exception_type((requests.Timeout, requests.ConnectionError, EmptyResponseError))
            | retry_if_exception(_is_retryable_http_error)
        ),
        before_sleep=before_sleep,
        reraise=True,
    )

    try:
        return retrying(request_once)
    except requests.HTTPError as exc:
        if logger:
            response = getattr(exc, "response", None)
            if response is not None:
                logger.error(
                    "HTTP error for %s | context=%s | status=%s",
                    url,
                    context or "request",
                    response.status_code,
                )
            else:
                logger.error("HTTP error for %s | context=%s | %s", url, context or "request", exc)
        return None
    except Exception as exc:
        if logger:
            logger.error("Request failed for %s | context=%s | %s", url, context or "request", exc)
        return None
