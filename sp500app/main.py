"""Bloomberg AI monetization research — standalone script.

公司电脑只需复制本文件；无需复制其他本地 Python 文件或安装本项目。
Requires Python 3.10+ and a logged-in Bloomberg Desktop API session.

Install once:
    python -m pip install "pandas>=2.2,<3" "numpy>=1.26,<3" "matplotlib>=3.8,<4"
    python -m pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi
Optional Parquet output:
    python -m pip install "pyarrow>=15"
Run:
    python main.py --test --validate-only
    python main.py --test
    python main.py

Edit CONFIGURATION below to change fields, fiscal years, or Software filters.
Data/output folders are created automatically. Test output goes to runs/test.
All real data come from Bloomberg; this script contains no synthetic data.
Period comparisons describe labor leverage, not a causal AI effect.
"""

# ruff: noqa: E402
from __future__ import annotations

import sys

if sys.version_info < (3, 10):
    raise SystemExit("Python 3.10+ is required. Activate a compatible Python environment first.")

import argparse
import base64
import importlib
import importlib.metadata
import json
import logging
import math
import os
import platform
import re
import tempfile
import time
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import asdict, dataclass, field, replace
from datetime import date, datetime, timezone
from datetime import time as datetime_time
from pathlib import Path
from types import TracebackType
from typing import Any, Protocol
from uuid import uuid4

import numpy as np
import pandas as pd

__version__ = "0.2.0"
LOGGER = logging.getLogger(__name__)


# ========================================================================
# CONFIGURATION — edit fields, years and Software definition here
# ========================================================================

START_FY = 2015
END_FY = 2025
PRE_AI_YEARS = (2017, 2018, 2019, 2020, 2021, 2022)
POST_AI_YEARS = (2023, 2024, 2025)
TEST_SECURITIES = ("MSFT US Equity", "CRM US Equity", "ADBE US Equity", "NOW US Equity", "ORCL US Equity")
VALIDATION_SECURITIES = ("MSFT US Equity", "CRM US Equity", "ORCL US Equity")
FIELD_CANDIDATES = {
    "revenue": ["SALES_REV_TURN"],
    "gross_profit": ["GROSS_PROFIT", "IS_GROSS_PROFIT"],
    "operating_income": ["IS_OPER_INC", "EBIT"],
    "ebitda": ["EBITDA"],
    "rd_expense": ["IS_RD_EXPEND"],
    "sga_expense": ["IS_SGA_EXPENSE", "IS_SGA_OTHER_OP_DEPR_OP_MAINT"],
    "personnel_expense": ["IS_PERSONNEL_EXP"],
    "employees": ["NUM_OF_EMPLOYEES"],
    "free_cash_flow": ["CF_FREE_CASH_FLOW"],
    "operating_cash_flow": ["CF_CASH_FROM_OPER"],
    "capex": ["CF_CAP_EXPEND_PRPTY_ADD"],
    "market_cap": ["HISTORICAL_MARKET_CAP"],
}
CLASSIFICATION_CANDIDATES = {
    "company_name": ["NAME"],
    "gics_sector": ["GICS_SECTOR_NAME"],
    "gics_industry_group": ["GICS_INDUSTRY_GROUP_NAME"],
    "gics_industry": ["GICS_INDUSTRY_NAME"],
    "gics_sub_industry": ["GICS_SUB_INDUSTRY_NAME"],
}
METADATA_CANDIDATES = {
    "reported_fiscal_year": ["EQY_FUND_YEAR"],
    "reported_period": ["FUND_PER"],
    "reported_currency": ["EQY_FUND_CRNCY"],
}
SNAPSHOT_CANDIDATES = {
    key: list(FIELD_CANDIDATES[key])
    for key in ("revenue", "gross_profit", "operating_income", "ebitda", "free_cash_flow", "employees")
}
SNAPSHOT_CANDIDATES["market_cap"] = ["CUR_MKT_CAP"]
MONETARY_VARIABLES = tuple(key for key in FIELD_CANDIDATES if key != "employees")
SOFTWARE_TERMS = ("Software", "Application Software", "Systems Software", "IT Services")
FISCAL_PERIOD_CANDIDATES = ("Y", "FY")
ANNUAL_PERIOD_LABELS = frozenset(("Y", "FY", "AY", "ANNUAL", "YEARLY"))
TARGET_CURRENCY = "USD"
CURRENCY_CONTROL = "EUR"
SCALING_FORMAT = "MLN"
MONETARY_MULTIPLIER = 1_000_000.0
CONTROL_SCALING_FORMAT = "BLN"
CONTROL_MONETARY_MULTIPLIER = 1_000_000_000.0
SCALING_RTOL = 0.02
MIN_PERIOD_OBSERVATIONS = 2
WINSOR_QUANTILES = (0.01, 0.99)


@dataclass(frozen=True)
class Settings:
    output_dir: Path = Path(".")
    host: str = "localhost"
    port: int = 8194
    batch_size: int = 100
    timeout_seconds: float = 60.0
    retries: int = 2
    test: bool = False
    parquet: bool = True

    @property
    def raw_dir(self) -> Path:
        return self.output_dir / "data" / "raw"

    @property
    def processed_dir(self) -> Path:
        return self.output_dir / "data" / "processed"

    @property
    def diagnostics_dir(self) -> Path:
        return self.output_dir / "output" / "diagnostics"

    def make_dirs(self) -> None:
        for path in (
            self.raw_dir,
            self.processed_dir,
            self.diagnostics_dir,
            self.output_dir / "output" / "charts",
            self.output_dir / "output" / "tables",
        ):
            path.mkdir(parents=True, exist_ok=True)


# ========================================================================
# Atomic local output
# ========================================================================


def write_json(value: Any, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=path.parent, suffix=".tmp", delete=False
        ) as handle:
            tmp = Path(handle.name)
            json.dump(value, handle, indent=2, default=str, allow_nan=False)
            handle.write("\n")
        os.replace(tmp, path)
    finally:
        if tmp is not None:
            tmp.unlink(missing_ok=True)


def write_frame(frame: pd.DataFrame, path: Path, parquet: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".tmp", delete=False) as handle:
        tmp = Path(handle.name)
    try:
        frame.to_csv(tmp, index=False)
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)
    target = path.with_suffix(".parquet")
    if parquet:
        with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".parquet", delete=False) as handle:
            tmp = Path(handle.name)
        try:
            frame.to_parquet(tmp, index=False)
            os.replace(tmp, target)
        except (ImportError, ValueError, TypeError, NotImplementedError) as exc:
            target.unlink(missing_ok=True)
            LOGGER.warning("Optional Parquet export unavailable (%s). CSV saved: %s", exc, path)
        finally:
            tmp.unlink(missing_ok=True)
    else:
        target.unlink(missing_ok=True)


# ========================================================================
# Bloomberg Desktop API client
# ========================================================================

SERVICE = "//blp/refdata"


class BloombergError(RuntimeError):
    """Session, transport, timeout, malformed response, or archive failure."""

    def __init__(self, message: str, *, retryable: bool = False) -> None:
        super().__init__(message)
        self.retryable = retryable


@dataclass(frozen=True)
class BloombergIssue:
    security: str | None
    field: str | None
    kind: str
    category: str = ""
    subcategory: str = ""
    code: int | None = None
    message: str = ""


@dataclass
class BloombergResult:
    data: dict[str, dict[str, Any]]
    issues: list[BloombergIssue]
    request_id: str = ""


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _element_value(element: Any) -> Any:
    """Convert SDK elements, including nested bulk tables, without flattening."""
    if element.isArray():
        if element.isComplexType():
            return [_element_value(element.getValueAsElement(i)) for i in range(element.numValues())]
        return [
            None if element.isNullValue(i) else _scalar_value(element.getValue(i))
            for i in range(element.numValues())
        ]
    if element.isNull():
        return None
    if element.isComplexType():
        result = {}
        for i in range(element.numElements()):
            child = element.getElement(i)
            result[str(child.name())] = _element_value(child)
        return result
    return _scalar_value(element.getValue())


def _scalar_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, bytes, bool, int, float, date, datetime_time)):
        return value
    return str(value)


def _json_ready(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _json_ready(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_ready(item) for item in value]
    if isinstance(value, (date, datetime_time)):
        return value.isoformat()
    if isinstance(value, bytes):
        return {"encoding": "base64", "value": base64.b64encode(value).decode("ascii")}
    if isinstance(value, float) and (not math.isfinite(value)):
        return str(value)
    return value


def _issue(
    payload: dict[str, Any],
    *,
    security: str | None = None,
    field: str | None = None,
    fallback: str = "response_error",
) -> BloombergIssue:
    category = str(payload.get("category") or "")
    subcategory = str(payload.get("subcategory") or "")
    message = str(payload.get("message") or payload.get("description") or "")
    tokens = f"{category} {subcategory} {message}".upper()
    if any(
        (
            token in tokens
            for token in (
                "NOT_ENTITLED",
                "NO_AUTH",
                "NOT_AUTHORIZED",
                "NOT_AUTHORISED",
                "PERMISSION_DENIED",
                "NOT ENTITLED",
                "NOT AUTHORIZED",
            )
        )
    ):
        kind = "entitlement"
    elif any(
        (
            token in tokens
            for token in (
                "BAD_FLD",
                "INVALID_FIELD",
                "UNKNOWN_FIELD",
                "FIELD_NOT_FOUND",
                "INVALID FIELD",
                "UNKNOWN FIELD",
            )
        )
    ):
        kind = "invalid_field"
    else:
        kind = fallback
    try:
        code = int(payload["code"]) if payload.get("code") is not None else None
    except (TypeError, ValueError):
        code = None
    return BloombergIssue(security, field, kind, category, subcategory, code, message)


def _is_transient(issue: BloombergIssue) -> bool:
    """Retry an explicit, narrow set of transport/server failures only."""
    if issue.kind != "response_error":
        return False
    categories = {issue.category.upper(), issue.subcategory.upper()}
    return bool(
        categories
        & {
            "TIMEOUT",
            "TIME_OUT",
            "CONNECTION_TIMEOUT",
            "CONNECTION_DOWN",
            "SERVICE_UNAVAILABLE",
            "TEMPORARY_UNAVAILABLE",
            "TEMPORARILY_UNAVAILABLE",
            "INTERNAL_ERROR",
            "SERVER_ERROR",
        }
    )


class BloombergClient:
    """Reference-data client using the logged-in local Bloomberg Desktop session.

    ``retries`` counts additional attempts after the first attempt. Field and
    security errors are returned as issues rather than raised. Timeout and
    explicitly transient request failures are retried; invalid fields,
    entitlements, bad overrides and usage limits are never retried.
    """

    def __init__(
        self,
        host: str = "localhost",
        port: int = 8194,
        timeout_seconds: float = 60,
        retries: int = 2,
        batch_size: int = 100,
        raw_dir: Path | None = None,
    ) -> None:
        if not host or not isinstance(host, str):
            raise ValueError("host must be a nonempty string")
        if isinstance(port, bool) or not isinstance(port, int) or (not 1 <= port <= 65535):
            raise ValueError("port must be an integer between 1 and 65535")
        if not math.isfinite(timeout_seconds) or timeout_seconds <= 0:
            raise ValueError("timeout_seconds must be positive and finite")
        if isinstance(retries, bool) or not isinstance(retries, int) or retries < 0:
            raise ValueError("retries must be a nonnegative integer")
        if isinstance(batch_size, bool) or not isinstance(batch_size, int) or batch_size <= 0:
            raise ValueError("batch_size must be a positive integer")
        self.host = host
        self.port = port
        self.timeout_seconds = timeout_seconds
        self.retries = retries
        self.batch_size = batch_size
        self.raw_dir = Path(raw_dir) if raw_dir is not None else None
        self._sdk: Any = None
        self._session: Any = None
        self._service: Any = None

    def __enter__(self) -> BloombergClient:
        return self.start()

    def __exit__(
        self, exc_type: type[BaseException] | None, exc: BaseException | None, traceback: TracebackType | None
    ) -> None:
        try:
            self.stop()
        except BloombergError:
            if exc is None:
                raise
            LOGGER.exception("Bloomberg session cleanup failed after an earlier error")

    def start(self) -> BloombergClient:
        if self._session is not None:
            return self
        try:
            self._sdk = importlib.import_module("blpapi")
        except (ImportError, OSError) as exc:
            raise BloombergError(
                "The official blpapi package/native library is unavailable. Install Bloomberg's Python SDK on the Bloomberg Desktop host and ensure Terminal is running and logged in."
            ) from exc
        session = None
        try:
            options = self._sdk.SessionOptions()
            options.setServerHost(self.host)
            options.setServerPort(self.port)
            options.setClientMode(self._sdk.SessionOptions.DAPI)
            options.setConnectTimeout(max(1, min(120000, math.ceil(self.timeout_seconds * 1000))))
            session = self._sdk.Session(options)
            if not session.start():
                raise BloombergError(
                    f"Cannot start Bloomberg Desktop session at {self.host}:{self.port}; check that Terminal is running, logged in, and API access is enabled."
                )
            if not session.openService(SERVICE):
                raise BloombergError(f"Bloomberg could not open {SERVICE}; check Desktop API access.")
            service = session.getService(SERVICE)
        except Exception as exc:
            if session is not None:
                try:
                    session.stop()
                except Exception:
                    LOGGER.exception("Bloomberg session cleanup failed during startup")
            if isinstance(exc, BloombergError):
                raise
            raise BloombergError(f"Bloomberg Desktop session setup failed: {exc}") from exc
        self._session, self._service = (session, service)
        return self

    def stop(self) -> None:
        session, self._session, self._service = (self._session, None, None)
        if session is not None:
            try:
                if session.stop() is False:
                    raise BloombergError("Bloomberg Desktop session did not stop successfully")
            except BloombergError:
                raise
            except Exception as exc:
                raise BloombergError(f"Bloomberg session cleanup failed: {exc}") from exc

    def reference(
        self, securities: list[str], fields: list[str], overrides: dict[str, str] | None = None
    ) -> BloombergResult:
        if not fields or any((not isinstance(field, str) or not field.strip() for field in fields)):
            raise ValueError("fields must contain nonempty Bloomberg field names")
        if any((not isinstance(security, str) or not security.strip() for security in securities)):
            raise ValueError("securities must contain nonempty Bloomberg identifiers")
        override_values = dict(overrides or {})
        if any(
            (
                not isinstance(key, str) or not key or (not isinstance(value, str))
                for key, value in override_values.items()
            )
        ):
            raise ValueError("overrides must map nonempty field names to strings")
        unique_securities = list(dict.fromkeys(securities))
        unique_fields = list(dict.fromkeys(fields))
        request_id = uuid4().hex
        result = BloombergResult({}, [], request_id)
        if not unique_securities:
            return result
        self.start()
        for start in range(0, len(unique_securities), self.batch_size):
            batch = unique_securities[start : start + self.batch_size]
            batch_id = uuid4().hex
            for attempt in range(1, self.retries + 2):
                try:
                    response = self._attempt(
                        batch,
                        unique_fields,
                        override_values,
                        request_id=request_id,
                        batch_id=batch_id,
                        attempt=attempt,
                    )
                except BloombergError as exc:
                    if not exc.retryable or attempt > self.retries:
                        raise
                    LOGGER.warning("Bloomberg attempt %d failed; retrying: %s", attempt, exc)
                else:
                    transient = response.issues and all((_is_transient(issue) for issue in response.issues))
                    if not transient or attempt > self.retries:
                        for security, values in response.data.items():
                            result.data.setdefault(security, {}).update(values)
                        for issue in response.issues:
                            if issue.security is None:
                                result.issues.extend(
                                    (replace(issue, security=security) for security in batch)
                                )
                            else:
                                result.issues.append(issue)
                        break
                    LOGGER.warning(
                        "Bloomberg returned transient response errors; retrying attempt %d", attempt
                    )
                time.sleep(min(0.5 * 2 ** (attempt - 1), 5.0))
        return result

    def _attempt(
        self,
        securities: list[str],
        fields: list[str],
        overrides: dict[str, str],
        *,
        request_id: str,
        batch_id: str,
        attempt: int,
    ) -> BloombergResult:
        attempt_id = uuid4().hex
        archive: dict[str, Any] = {
            "request_id": request_id,
            "batch_id": batch_id,
            "attempt_id": attempt_id,
            "attempt": attempt,
            "started_at": _utc_now(),
            "service": SERVICE,
            "request_type": "ReferenceDataRequest",
            "host": self.host,
            "port": self.port,
            "securities": securities,
            "fields": fields,
            "overrides": overrides,
            "messages": [],
            "timeout_seconds": self.timeout_seconds,
        }
        result = BloombergResult({}, [], request_id)
        sent, complete = (False, False)
        correlation = self._sdk.CorrelationId(attempt_id)
        queue = self._sdk.EventQueue()
        try:
            request = self._service.createRequest("ReferenceDataRequest")
            for security in securities:
                request.getElement("securities").appendValue(security)
            for field in fields:
                request.getElement("fields").appendValue(field)
            for field, value in overrides.items():
                override = request.getElement("overrides").appendElement()
                override.setElement("fieldId", field)
                override.setElement("value", value)
            archive["request_text"] = str(request)
            self._session.sendRequest(request, correlationId=correlation, eventQueue=queue)
            sent = True
            deadline = time.monotonic() + self.timeout_seconds
            while not complete:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise BloombergError(
                        f"Bloomberg request {attempt_id} timed out after {self.timeout_seconds:g}s",
                        retryable=True,
                    )
                event = queue.nextEvent(max(1, math.ceil(remaining * 1000)))
                event_type = event.eventType()
                if event_type == self._sdk.Event.TIMEOUT:
                    raise BloombergError(
                        f"Bloomberg request {attempt_id} timed out after {self.timeout_seconds:g}s",
                        retryable=True,
                    )
                matched_response = False
                for message in event:
                    ids = message.correlationIds()
                    matched = not ids or any((item.value() == attempt_id for item in ids))
                    entry: dict[str, Any] = {
                        "received_at": _utc_now(),
                        "event_type": event_type,
                        "message_type": str(message.messageType()),
                        "correlation_ids": [str(item.value()) for item in ids],
                        "matched_request": matched,
                        "raw_text": str(message),
                    }
                    archive["messages"].append(entry)
                    if not matched:
                        continue
                    payload = _element_value(message.asElement())
                    entry["data"] = payload
                    if not isinstance(payload, dict):
                        raise BloombergError("Bloomberg returned a non-object response")
                    if event_type in (self._sdk.Event.PARTIAL_RESPONSE, self._sdk.Event.RESPONSE):
                        matched_response = True
                        self._merge_payload(payload, result, securities)
                    elif event_type == self._sdk.Event.REQUEST_STATUS:
                        if str(message.messageType()) in (
                            "RequestFailure",
                            "RequestCanceled",
                            "RequestTerminated",
                        ):
                            reason = payload.get("reason") or payload.get("responseError") or payload
                            result.issues.append(_issue(reason))
                            complete = True
                    elif event_type == self._sdk.Event.SESSION_STATUS:
                        if str(message.messageType()) in ("SessionTerminated", "SessionStartupFailure"):
                            raise BloombergError(f"Bloomberg session failed: {message}")
                if event_type == self._sdk.Event.RESPONSE and matched_response:
                    complete = True
            if any((issue.security is None and issue.field is None for issue in result.issues)):
                result.data.clear()
            archive["status"] = "completed_with_issues" if result.issues else "completed"
            archive["issues"] = [asdict(issue) for issue in result.issues]
            return result
        except Exception as exc:
            archive["status"] = "failed"
            archive["error"] = f"{type(exc).__name__}: {exc}"
            if isinstance(exc, BloombergError):
                raise
            raise BloombergError(f"Bloomberg request {attempt_id} failed: {exc}") from exc
        finally:
            if sent and (not complete):
                try:
                    self._session.cancel(correlation)
                except Exception as exc:
                    archive["cancellation_error"] = str(exc)
                    LOGGER.warning(
                        "Could not cancel Bloomberg request %s: %s", attempt_id, exc, exc_info=True
                    )
            archive["completed_at"] = _utc_now()
            self._save_archive(archive)

    @staticmethod
    def _merge_payload(
        payload: dict[str, Any], result: BloombergResult, requested_securities: list[str]
    ) -> None:
        if payload.get("responseError"):
            result.issues.append(_issue(payload["responseError"]))
        securities = payload.get("securityData") or []
        if isinstance(securities, dict):
            securities = [securities]
        for entry in securities:
            response_security = str(entry.get("security") or "")
            sequence = entry.get("sequenceNumber")
            if (
                isinstance(sequence, int)
                and (not isinstance(sequence, bool))
                and (0 <= sequence < len(requested_securities))
            ):
                security = requested_securities[sequence]
            elif response_security in requested_securities:
                security = response_security
            else:
                result.issues.append(
                    BloombergIssue(
                        None,
                        None,
                        "response_error",
                        category="UNMAPPED_SECURITY",
                        message=f"Cannot map returned security {response_security!r} and sequenceNumber {sequence!r} to the request",
                    )
                )
                continue
            if entry.get("securityError"):
                result.issues.append(
                    _issue(entry["securityError"], security=security, fallback="security_error")
                )
            if security is not None:
                values = entry.get("fieldData") or {}
                if not isinstance(values, dict):
                    raise BloombergError(f"Unexpected reference fieldData format for {security}")
                result.data.setdefault(security, {}).update(values)
            for exception in entry.get("fieldExceptions") or []:
                result.issues.append(
                    _issue(
                        exception.get("errorInfo") or {},
                        security=security,
                        field=str(exception.get("fieldId") or "") or None,
                        fallback="field_exception",
                    )
                )

    def _save_archive(self, archive: dict[str, Any]) -> None:
        if self.raw_dir is None:
            return
        target = self.raw_dir / f"{archive['request_id']}_{archive['batch_id']}_{archive['attempt']}.json"
        temporary = target.with_suffix(f".{uuid4().hex}.tmp")
        try:
            self.raw_dir.mkdir(parents=True, exist_ok=True)
            with temporary.open("x", encoding="utf-8") as handle:
                json.dump(_json_ready(archive), handle, ensure_ascii=False, allow_nan=False, indent=2)
                handle.write("\n")
            temporary.replace(target)
        except (OSError, TypeError, ValueError) as exc:
            raise BloombergError(f"Cannot save Bloomberg raw response archive {target}: {exc}") from exc
        finally:
            try:
                temporary.unlink()
            except FileNotFoundError:
                pass
            except OSError:
                LOGGER.warning("Could not remove incomplete archive %s", temporary, exc_info=True)


# ========================================================================
# Index universe and classifications
# ========================================================================

CLASSIFICATION_COLUMNS = (
    "company_name",
    "gics_sector",
    "gics_industry_group",
    "gics_industry",
    "gics_sub_industry",
)
UNIVERSE_COLUMNS = ("ticker", "security", *CLASSIFICATION_COLUMNS)
_MEMBER_KEYS = {"membertickerandexchangecode", "tickerandexchangecode", "memberticker", "security", "member"}


class ReferenceClient(Protocol):
    """The small portion of the Bloomberg client needed by this module."""

    def reference(
        self, securities: Sequence[str], fields: Sequence[str], overrides: Mapping[str, str] | None = None
    ) -> Any: ...


def _normalize_security(value: object) -> str | None:
    """Keep Bloomberg's exchange and share-class tokens without guessing either."""
    if not isinstance(value, str):
        return None
    tokens = value.split()
    if tokens and tokens[-1].lower() == "equity":
        tokens.pop()
    if len(tokens) < 2 or not re.fullmatch("[A-Za-z]{2,3}", tokens[-1]):
        return None
    if any((not re.fullmatch("[A-Za-z0-9./&+\\-]+", part) for part in tokens[:-1])):
        return None
    return " ".join(tokens).upper() + " Equity"


def _member_security(row: object) -> str | None:
    if isinstance(row, str):
        return _normalize_security(row)
    if isinstance(row, Mapping):
        for key, value in row.items():
            normalized_key = re.sub("[^a-z0-9]", "", str(key).lower())
            if normalized_key in _MEMBER_KEYS:
                security = _normalize_security(value)
                if security is not None:
                    return security
    return None


def _log_issues(issues: Sequence[Any], context: str) -> None:
    for issue in issues:
        LOGGER.warning(
            "%s: kind=%s security=%s field=%s category=%s message=%s",
            context,
            getattr(issue, "kind", "unknown"),
            getattr(issue, "security", ""),
            getattr(issue, "field", ""),
            getattr(issue, "category", ""),
            getattr(issue, "message", str(issue)),
        )


def _issue_blocks(issues: Sequence[Any], security: str, field: str) -> bool:
    """Discard values implicated in an exception, even if a payload contains one."""
    for issue in issues:
        issue_security = getattr(issue, "security", None)
        issue_field = getattr(issue, "field", None)
        if issue_security and issue_security != security:
            continue
        if not issue_field or issue_field == field:
            return True
    return False


def _classification_value(value: object) -> str | float:
    if not isinstance(value, str) or not value.strip():
        return np.nan
    return value.strip()


def fetch_universe(
    client: ReferenceClient, classification_fields: dict[str, str | None], test: bool = False
) -> pd.DataFrame:
    """Fetch current SPX members (or the fixed five-name live debug universe).

    ``classification_fields`` contains the field validation routine's selected
    Bloomberg mnemonics. Missing or rejected fields remain NaN. This function
    never substitutes a mnemonic or fabricates an exchange code. The client is
    responsible for batching reference requests and persisting API responses.
    """
    if test:
        securities = list(TEST_SECURITIES)
    else:
        members = client.reference(["SPX Index"], ["INDX_MEMBERS"])
        _log_issues(members.issues, "Index membership")
        if _issue_blocks(members.issues, "SPX Index", "INDX_MEMBERS"):
            raise ValueError(
                "Bloomberg could not provide SPX Index INDX_MEMBERS; inspect the logged Bloomberg exceptions and raw responses."
            )
        index_data = members.data.get("SPX Index", {})
        bulk = index_data.get("INDX_MEMBERS") if isinstance(index_data, Mapping) else None
        if not isinstance(bulk, (list, tuple)):
            raise ValueError("SPX Index INDX_MEMBERS is missing or is not a bulk sequence.")
        securities = []
        seen = set()
        rejected = 0
        for row_number, row in enumerate(bulk, start=1):
            security = _member_security(row)
            if security is None:
                rejected += 1
                LOGGER.warning("Skipping unusable INDX_MEMBERS row %d: %r", row_number, row)
                continue
            if security not in seen:
                securities.append(security)
                seen.add(security)
        if not securities:
            raise ValueError("SPX Index INDX_MEMBERS returned no usable equity securities.")
        if rejected:
            LOGGER.warning(
                "Universe has %d usable securities; %d malformed membership rows were excluded.",
                len(securities),
                rejected,
            )
    rows = [
        {"ticker": security.removesuffix(" Equity").rsplit(" ", 1)[0], "security": security}
        for security in securities
    ]
    fields = list(
        dict.fromkeys(
            (field for column in CLASSIFICATION_COLUMNS if (field := classification_fields.get(column)))
        )
    )
    classifications = client.reference(securities, fields) if fields else None
    if classifications is not None:
        _log_issues(classifications.issues, "Universe classification")
    for row in rows:
        security = str(row["security"])
        values = classifications.data.get(security, {}) if classifications is not None else {}
        if not isinstance(values, Mapping):
            values = {}
        for column in CLASSIFICATION_COLUMNS:
            field = classification_fields.get(column)
            blocked = (
                classifications is not None
                and field is not None
                and _issue_blocks(classifications.issues, security, field)
            )
            row[column] = _classification_value(values.get(field)) if field and (not blocked) else np.nan
    return pd.DataFrame(rows, columns=list(UNIVERSE_COLUMNS))


def print_classifications(universe: pd.DataFrame) -> None:
    """Print the observed current GICS taxonomy for reviewing Software filters."""
    for column in CLASSIFICATION_COLUMNS[1:]:
        values = sorted(universe[column].dropna().astype(str).unique())
        print(f"{column}: {(', '.join(values) if values else '(no available classifications)')}")


# ========================================================================
# Live field and override validation
# ========================================================================


def numeric(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    try:
        number = float(value)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def relevant_issues(result: BloombergResult, security: str, mnemonic: str) -> list:
    return [
        issue
        for issue in result.issues
        if issue.security in (None, "", security) and issue.field in (None, "", mnemonic)
    ]


def clean_value(result: BloombergResult, security: str, mnemonic: str | None) -> Any:
    if not mnemonic or relevant_issues(result, security, mnemonic):
        return None
    return result.data.get(security, {}).get(mnemonic)


def reported_year(result: BloombergResult, security: str, metadata: dict) -> int | None:
    value = numeric(clean_value(result, security, metadata.get("reported_fiscal_year")))
    return int(value) if value is not None and value.is_integer() and (1900 <= value <= 2200) else None


def fiscal_identity(result: BloombergResult, security: str, metadata: dict, year: int | None = None) -> bool:
    observed = reported_year(result, security, metadata)
    period = str(clean_value(result, security, metadata.get("reported_period"))).strip().upper()
    return observed is not None and (year is None or observed == year) and (period in ANNUAL_PERIOD_LABELS)


@dataclass
class ValidationReport:
    selected: dict[str, str | None] = field(default_factory=dict)
    classifications: dict[str, str | None] = field(default_factory=dict)
    metadata: dict[str, str | None] = field(default_factory=dict)
    snapshot_selected: dict[str, str | None] = field(default_factory=dict)
    annual_verified: dict[str, bool] = field(default_factory=dict)
    snapshot_verified: dict[str, bool] = field(default_factory=dict)
    fiscal_period: str | None = None
    rows: list[dict] = field(default_factory=list)
    checks: list[dict] = field(default_factory=list)
    selections: list[dict] = field(default_factory=list)

    def save(self, directory: Path) -> None:
        write_frame(
            pd.DataFrame(
                self.rows,
                columns=[
                    "context",
                    "security",
                    "field",
                    "status",
                    "value",
                    "exceptions",
                    "overrides",
                    "request_id",
                ],
            ),
            directory / "field_validation.csv",
        )
        write_frame(
            pd.DataFrame(
                self.checks, columns=["scope", "variable", "field", "verified", "witnesses", "reason"]
            ),
            directory / "override_validation.csv",
        )
        write_frame(pd.DataFrame(self.selections), directory / "selected_fields.csv")
        write_json(
            {
                "fiscal_period": self.fiscal_period,
                "annual_fields": self.selected,
                "annual_verified": self.annual_verified,
                "snapshot_fields": self.snapshot_selected,
                "snapshot_verified": self.snapshot_verified,
                "classifications": self.classifications,
                "metadata": self.metadata,
                "target_currency": TARGET_CURRENCY,
                "scaling_format": SCALING_FORMAT,
                "monetary_multiplier": MONETARY_MULTIPLIER,
            },
            directory / "validation_summary.json",
        )


def record_result(
    report: ValidationReport,
    result: BloombergResult,
    securities: list[str],
    fields: list[str],
    context: str,
    overrides: dict | None = None,
) -> None:
    for security in securities:
        for mnemonic in fields:
            issues = relevant_issues(result, security, mnemonic)
            value = result.data.get(security, {}).get(mnemonic)
            missing = value is None or (isinstance(value, float) and (not math.isfinite(value)))
            status = (
                "|".join(sorted({i.kind for i in issues}))
                if issues
                else "missing_security"
                if security not in result.data
                else "all_null"
                if missing
                else "valid"
            )
            report.rows.append(
                {
                    "context": context,
                    "security": security,
                    "field": mnemonic,
                    "status": status,
                    "value": json.dumps(value, default=str),
                    "exceptions": json.dumps([vars(i) for i in issues], default=str),
                    "overrides": json.dumps(overrides or {}, sort_keys=True),
                    "request_id": result.request_id,
                }
            )


def _select(
    report: ValidationReport, candidates: dict, result: BloombergResult, securities: list[str], scope: str
) -> dict:
    selected = {}
    for variable, options in candidates.items():
        selected[variable] = None
        reason = "All candidates returned invalid-field errors"
        for position, mnemonic in enumerate(options):
            statuses = []
            for security in securities:
                issues = relevant_issues(result, security, mnemonic)
                statuses.append(
                    {i.kind for i in issues}
                    or {"supported" if security in result.data else "missing_security"}
                )
            if all((status == {"invalid_field"} for status in statuses)):
                continue
            if any(("supported" in status for status in statuses)):
                selected[variable] = mnemonic
                reason = (
                    "Primary candidate" if position == 0 else "Fallback after explicit invalid-field errors"
                )
                if not any((clean_value(result, s, mnemonic) is not None for s in securities)):
                    reason += "; all probe values null; no substitution for coverage"
                if position:
                    LOGGER.warning(
                        "%s.%s explicitly selects %s after invalid %s; review economic definition",
                        scope,
                        variable,
                        mnemonic,
                        options[:position],
                    )
            else:
                reason = (
                    "Non-field error blocks selection; no fallback on entitlements/security/null coverage"
                )
            break
        report.selections.append(
            {
                "scope": scope,
                "variable": variable,
                "selected_field": selected[variable],
                "candidates": "|".join(options),
                "reason": reason,
            }
        )
    return selected


def annual_overrides(year: int, period: str, monetary: bool) -> dict[str, str]:
    overrides = {"EQY_FUND_YEAR": str(year), "FUND_PER": period}
    if monetary:
        overrides.update(EQY_FUND_CRNCY=TARGET_CURRENCY, SCALING_FORMAT=SCALING_FORMAT)
    return overrides


def _request(client, report, securities, fields, context, overrides=None) -> BloombergResult:
    fields = list(dict.fromkeys((f for f in fields if f)))
    if not fields:
        return BloombergResult(data={}, issues=[])
    result = client.reference(securities, fields, overrides=overrides)
    record_result(report, result, securities, fields, context, overrides)
    return result


def _unit_witnesses(
    base: BloombergResult,
    scale: BloombergResult,
    currency: BloombergResult,
    securities: list[str],
    mnemonic: str,
    metadata: dict,
    *,
    fiscal: bool = False,
    year: int | None = None,
) -> list[str]:
    witnesses = []
    ratio = CONTROL_MONETARY_MULTIPLIER / MONETARY_MULTIPLIER
    for security in securities:
        if fiscal:
            expected_year = year if year is not None else reported_year(base, security, metadata)
            if expected_year is None or not all(
                (
                    fiscal_identity(result, security, metadata, expected_year)
                    for result in (base, scale, currency)
                )
            ):
                continue
        values = [numeric(clean_value(result, security, mnemonic)) for result in (base, scale, currency)]
        if any((v is None or v == 0 for v in values)):
            continue
        base_value, scale_value, currency_value = values
        base_currency = str(clean_value(base, security, metadata.get("reported_currency"))).upper()
        scale_currency = str(clean_value(scale, security, metadata.get("reported_currency"))).upper()
        control_currency = str(clean_value(currency, security, metadata.get("reported_currency"))).upper()
        if (
            math.isclose(base_value / scale_value, ratio, rel_tol=SCALING_RTOL)
            and (not math.isclose(base_value, currency_value, rel_tol=1e-06))
            and (base_currency == TARGET_CURRENCY)
            and (scale_currency == TARGET_CURRENCY)
            and (control_currency == CURRENCY_CONTROL)
        ):
            witnesses.append(security)
    return witnesses


def _controls(client, report, securities, fields, context, overrides):
    scale_overrides = {**overrides, "SCALING_FORMAT": CONTROL_SCALING_FORMAT}
    currency_overrides = {**overrides, "EQY_FUND_CRNCY": CURRENCY_CONTROL}
    scale = _request(client, report, securities, fields, context + ":scale_control", scale_overrides)
    currency = _request(client, report, securities, fields, context + ":currency_control", currency_overrides)
    return (scale, currency)


def validate_fields(
    client, diagnostics_dir: Path, securities: tuple[str, ...] = VALIDATION_SECURITIES
) -> ValidationReport:
    """Test all candidate mnemonics, then certify each intended request context.

    The controls deliberately fail closed: fields with no usable witness stay
    null. Read the CSVs and inspect FLDS on the Terminal to resolve a failed check.
    """
    securities = list(securities)
    report = ValidationReport()
    maps = (FIELD_CANDIDATES, CLASSIFICATION_CANDIDATES, METADATA_CANDIDATES, SNAPSHOT_CANDIDATES)
    all_fields = list(
        dict.fromkeys((f for mapping in maps for candidates in mapping.values() for f in candidates))
    )
    try:
        available = _request(client, report, securities, all_fields, "candidate_availability")
        report.selected = _select(report, maps[0], available, securities, "annual")
        report.classifications = _select(report, maps[1], available, securities, "classification")
        report.metadata = _select(report, maps[2], available, securities, "metadata")
        report.snapshot_selected = _select(report, maps[3], available, securities, "snapshot")
        report.annual_verified = dict.fromkeys(FIELD_CANDIDATES, False)
        report.snapshot_verified = dict.fromkeys(SNAPSHOT_CANDIDATES, False)
        metadata = [f for f in report.metadata.values() if f]
        monetary = [report.selected[k] for k in MONETARY_VARIABLES if report.selected[k]]
        employee = report.selected.get("employees")
        probes = {}
        for period in FISCAL_PERIOD_CANDIDATES:
            probes = {}
            for year in (START_FY, END_FY):
                probes[year, "money"] = _request(
                    client,
                    report,
                    securities,
                    monetary + metadata,
                    f"annual:{period}:{year}:money",
                    annual_overrides(year, period, True),
                )
                probes[year, "employees"] = _request(
                    client,
                    report,
                    securities,
                    [employee] + metadata,
                    f"annual:{period}:{year}:employees",
                    annual_overrides(year, period, False),
                )
            old, new = (probes[START_FY, "money"], probes[END_FY, "money"])
            fiscal_witnesses = []
            for security in securities:
                if not (
                    fiscal_identity(old, security, report.metadata, START_FY)
                    and fiscal_identity(new, security, report.metadata, END_FY)
                ):
                    continue
                for mnemonic in monetary:
                    a, b = (
                        numeric(clean_value(old, security, mnemonic)),
                        numeric(clean_value(new, security, mnemonic)),
                    )
                    if a is not None and b is not None and (not math.isclose(a, b, rel_tol=1e-09)):
                        fiscal_witnesses.append(security)
                        break
            report.checks.append(
                {
                    "scope": "fiscal_period",
                    "variable": period,
                    "field": "FUND_PER",
                    "verified": bool(fiscal_witnesses),
                    "witnesses": "|".join(fiscal_witnesses),
                    "reason": "Requires annual period and exact fiscal-year metadata plus changing historical values",
                }
            )
            if fiscal_witnesses:
                report.fiscal_period = period
                break
        if report.fiscal_period:
            period = report.fiscal_period
            old, new = (probes[START_FY, "money"], probes[END_FY, "money"])
            scale, currency = _controls(
                client,
                report,
                securities,
                monetary + metadata,
                "annual",
                annual_overrides(END_FY, period, True),
            )
            for variable, mnemonic in report.selected.items():
                if not mnemonic:
                    continue
                left = old if variable != "employees" else probes[START_FY, "employees"]
                right = new if variable != "employees" else probes[END_FY, "employees"]
                witnesses = []
                unit_witnesses = (
                    securities
                    if variable == "employees"
                    else _unit_witnesses(
                        new, scale, currency, securities, mnemonic, report.metadata, fiscal=True, year=END_FY
                    )
                )
                for security in unit_witnesses:
                    a = numeric(clean_value(left, security, mnemonic))
                    b = numeric(clean_value(right, security, mnemonic))
                    if (
                        fiscal_identity(left, security, report.metadata, START_FY)
                        and fiscal_identity(right, security, report.metadata, END_FY)
                        and (a is not None)
                        and (b is not None)
                        and (not math.isclose(a, b, rel_tol=1e-09))
                    ):
                        witnesses.append(security)
                report.annual_verified[variable] = bool(witnesses)
                report.checks.append(
                    {
                        "scope": "annual",
                        "variable": variable,
                        "field": mnemonic,
                        "verified": bool(witnesses),
                        "witnesses": "|".join(witnesses),
                        "reason": "Fiscal identity and year response"
                        + (
                            "; currency metadata, USD/EUR response and MLN/BLN scale ratio"
                            if variable != "employees"
                            else "; headcount requested without monetary overrides"
                        ),
                    }
                )
        _validate_snapshot(client, report, securities, metadata)
    finally:
        report.save(diagnostics_dir)
    for scope, selected, verified in (
        ("annual", report.selected, report.annual_verified),
        ("snapshot", report.snapshot_selected, report.snapshot_verified),
    ):
        for variable, mnemonic in selected.items():
            LOGGER.info(
                "Validation %-8s %-20s %-32s %s",
                scope,
                variable,
                mnemonic or "unavailable",
                "VERIFIED" if verified.get(variable) else "NULL (unverified/unavailable; see diagnostics)",
            )
    LOGGER.info("Fiscal period selected: %s; diagnostics: %s", report.fiscal_period, diagnostics_dir)
    return report


def snapshot_overrides(report: ValidationReport, group: str) -> dict:
    if group == "employees":
        return {}
    result = {"EQY_FUND_CRNCY": TARGET_CURRENCY, "SCALING_FORMAT": SCALING_FORMAT}
    if group == "statements":
        result["FUND_PER"] = report.fiscal_period
    return result


def _validate_snapshot(client, report, securities, metadata):
    for group in ("statements", "market_cap", "employees"):
        variables = (
            [k for k in report.snapshot_selected if k not in ("employees", "market_cap")]
            if group == "statements"
            else [group]
        )
        if group == "statements" and (not report.fiscal_period):
            continue
        fields = [report.snapshot_selected[k] for k in variables if report.snapshot_selected[k]]
        fields += metadata if group != "employees" else []
        overrides = snapshot_overrides(report, group)
        base = _request(client, report, securities, fields, "snapshot:" + group, overrides)
        if group != "employees":
            scale, currency = _controls(client, report, securities, fields, "snapshot:" + group, overrides)
        for variable in variables:
            mnemonic = report.snapshot_selected[variable]
            if not mnemonic:
                continue
            if group == "employees":
                witnesses = [s for s in securities if numeric(clean_value(base, s, mnemonic)) is not None]
            else:
                witnesses = _unit_witnesses(
                    base, scale, currency, securities, mnemonic, report.metadata, fiscal=group == "statements"
                )
                if group == "statements":
                    witnesses = [s for s in witnesses if fiscal_identity(base, s, report.metadata)]
            report.snapshot_verified[variable] = bool(witnesses)
            report.checks.append(
                {
                    "scope": "snapshot:" + group,
                    "variable": variable,
                    "field": mnemonic,
                    "verified": bool(witnesses),
                    "witnesses": "|".join(witnesses),
                    "reason": "Latest annual actual with fiscal metadata"
                    if group == "statements"
                    else "Current, not a historical fiscal-period observation",
                }
            )


# ========================================================================
# Annual fundamentals and separate current snapshot
# ========================================================================

CLASSIFICATION_RENAMES = {
    "gics_sector": "sector",
    "gics_industry_group": "industry_group",
    "gics_industry": "industry",
    "gics_sub_industry": "sub_industry",
}


def _request_fields(report: ValidationReport, variables: list[str], snapshot=False) -> list[str]:
    selected = report.snapshot_selected if snapshot else report.selected
    fields = [selected[v] for v in variables if selected[v]]
    fields.extend((f for f in report.metadata.values() if f))
    return list(dict.fromkeys(fields))


def _observation(
    result: BloombergResult,
    security: str,
    mnemonic: str | None,
    report: ValidationReport,
    *,
    verified: bool,
    monetary: bool,
    fiscal: bool,
    year: int | None = None,
) -> tuple[float, str]:
    if not verified or not mnemonic:
        return (np.nan, "not_requested_unverified")
    issues = relevant_issues(result, security, mnemonic)
    if issues:
        return (np.nan, "|".join(sorted({i.kind for i in issues})))
    if security not in result.data:
        return (np.nan, "missing_security")
    if fiscal and (not fiscal_identity(result, security, report.metadata, year)):
        return (np.nan, "fiscal_metadata_mismatch")
    if monetary:
        currency = clean_value(result, security, report.metadata.get("reported_currency"))
        if str(currency).upper() != TARGET_CURRENCY:
            return (np.nan, "currency_metadata_mismatch")
    raw = clean_value(result, security, mnemonic)
    value = numeric(raw)
    if value is None:
        return (np.nan, "null" if raw is None else "non_numeric_or_nonfinite")
    return (value, "ok")


def pull_annual_panel(
    client, members: pd.DataFrame, report: ValidationReport, settings: Settings
) -> pd.DataFrame:
    if not report.fiscal_period:
        raise ValueError(
            "No annual fiscal-period override passed live validation; inspect field/override diagnostics"
        )
    if members.empty or members.security.duplicated().any():
        raise ValueError("Universe must be nonempty with unique securities")
    securities = members.security.tolist()
    metadata = members.rename(columns=CLASSIFICATION_RENAMES).copy()
    frames = []
    for year in range(START_FY, END_FY + 1):
        LOGGER.info("Retrieving FY%s for %s securities", year, len(securities))
        panel = metadata.copy()
        panel["fiscal_year"] = year
        panel["monetary_currency"] = TARGET_CURRENCY
        panel["monetary_multiplier"] = MONETARY_MULTIPLIER
        raw_rows = []
        for monetary in (True, False):
            all_variables = list(MONETARY_VARIABLES) if monetary else ["employees"]
            variables = [v for v in all_variables if report.annual_verified.get(v) and report.selected.get(v)]
            overrides = annual_overrides(year, report.fiscal_period, monetary)
            fields = _request_fields(report, variables)
            result = (
                client.reference(securities, fields, overrides=overrides)
                if variables
                else BloombergResult(data={}, issues=[])
            )
            record_result(
                report,
                result,
                securities,
                fields if variables else [],
                f"download:annual:{year}:{('money' if monetary else 'employees')}",
                overrides,
            )
            for variable in all_variables:
                mnemonic = report.selected.get(variable)
                values = []
                for security in securities:
                    value, status = _observation(
                        result,
                        security,
                        mnemonic,
                        report,
                        verified=variable in variables,
                        monetary=monetary,
                        fiscal=True,
                        year=year,
                    )
                    values.append(value)
                    raw_rows.append(
                        {
                            "security": security,
                            "fiscal_year": year,
                            "variable": variable,
                            "field": mnemonic,
                            "raw_value": result.data.get(security, {}).get(mnemonic),
                            "accepted_value": value,
                            "status": status,
                            "reported_fiscal_year": reported_year(result, security, report.metadata),
                            "reported_period": clean_value(
                                result, security, report.metadata.get("reported_period")
                            ),
                            "reported_currency": clean_value(
                                result, security, report.metadata.get("reported_currency")
                            ),
                            "overrides": json.dumps(overrides, sort_keys=True),
                            "request_id": result.request_id,
                        }
                    )
                panel[variable] = values
        write_frame(pd.DataFrame(raw_rows), settings.raw_dir / f"fundamentals_{year}.csv", settings.parquet)
        frames.append(panel)
    return pd.concat(frames, ignore_index=True)


def pull_current_snapshot(
    client, members: pd.DataFrame, report: ValidationReport, settings: Settings
) -> pd.DataFrame:
    """Latest annual actual statements plus current reported headcount/market cap.

    Fiscal years may differ by company and can be later than END_FY. This frame
    is never appended to the historical panel and has no historical YoY metrics.
    """
    securities = members.security.tolist()
    panel = members.rename(columns=CLASSIFICATION_RENAMES).copy()
    panel["snapshot_asof_utc"] = datetime.now(timezone.utc).isoformat()
    panel["statement_basis"] = "latest_annual_actual"
    panel["employees_basis"] = "latest_reported_count"
    panel["market_cap_basis"] = "current"
    panel["monetary_currency"] = TARGET_CURRENCY
    panel["monetary_multiplier"] = MONETARY_MULTIPLIER
    panel["statement_fiscal_year"] = pd.Series(pd.NA, index=panel.index, dtype="Int64")
    raw_rows = []
    for group in ("statements", "employees", "market_cap"):
        all_variables = (
            [k for k in SNAPSHOT_CANDIDATES if k not in ("employees", "market_cap")]
            if group == "statements"
            else [group]
        )
        variables = [
            v for v in all_variables if report.snapshot_verified.get(v) and report.snapshot_selected.get(v)
        ]
        overrides = snapshot_overrides(report, group)
        fields = (
            _request_fields(report, variables, snapshot=True)
            if group != "employees"
            else [report.snapshot_selected[v] for v in variables]
        )
        result = (
            client.reference(securities, fields, overrides=overrides)
            if variables
            else BloombergResult(data={}, issues=[])
        )
        record_result(
            report, result, securities, fields if variables else [], "download:snapshot:" + group, overrides
        )
        if group == "statements":
            panel["statement_fiscal_year"] = pd.array(
                [
                    reported_year(result, s, report.metadata)
                    if fiscal_identity(result, s, report.metadata)
                    else None
                    for s in securities
                ],
                dtype="Int64",
            )
        for variable in all_variables:
            mnemonic = report.snapshot_selected.get(variable)
            values, statuses = ([], [])
            for security in securities:
                value, status = _observation(
                    result,
                    security,
                    mnemonic,
                    report,
                    verified=variable in variables,
                    monetary=group != "employees",
                    fiscal=group == "statements",
                )
                values.append(value)
                statuses.append(status)
                raw_rows.append(
                    {
                        "security": security,
                        "variable": variable,
                        "field": mnemonic,
                        "basis": group,
                        "raw_value": result.data.get(security, {}).get(mnemonic),
                        "accepted_value": value,
                        "status": status,
                        "overrides": json.dumps(overrides, sort_keys=True),
                        "request_id": result.request_id,
                    }
                )
            panel[variable] = values
            panel[f"{variable}_status"] = statuses
            panel[f"{variable}_source_field"] = mnemonic
    write_frame(pd.DataFrame(raw_rows), settings.raw_dir / "current_snapshot.csv", settings.parquet)
    return panel


def quality_diagnostics(panel: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for variable in FIELD_CANDIDATES:
        values = pd.to_numeric(panel[variable], errors="coerce")
        by_security = panel.assign(_value=values).groupby("security")["_value"]
        constant = by_security.agg(["count", "nunique"])
        rows.append(
            {
                "variable": variable,
                "observations": len(panel),
                "non_null": int(values.notna().sum()),
                "missing_fraction": float(values.isna().mean()),
                "negative": int(values.lt(0).sum()),
                "zero": int(values.eq(0).sum()),
                "constant_securities_3plus_years": int(
                    ((constant["count"] >= 3) & (constant["nunique"] == 1)).sum()
                ),
            }
        )
    return pd.DataFrame(rows)


# ========================================================================
# Productivity, growth and labor leverage
# ========================================================================

PER_EMPLOYEE_SOURCES = {
    "revenue_per_employee": "revenue",
    "gross_profit_per_employee": "gross_profit",
    "operating_income_per_employee": "operating_income",
    "ebitda_per_employee": "ebitda",
    "fcf_per_employee": "free_cash_flow",
    "personnel_cost_per_employee": "personnel_expense",
}
RATIO_SOURCES = {
    "revenue_per_personnel_dollar": ("revenue", "personnel_expense"),
    "gross_profit_per_personnel_dollar": ("gross_profit", "personnel_expense"),
    "fcf_per_personnel_dollar": ("free_cash_flow", "personnel_expense"),
    "rd_to_revenue": ("rd_expense", "revenue"),
    "sga_to_revenue": ("sga_expense", "revenue"),
    "personnel_expense_to_revenue": ("personnel_expense", "revenue"),
    "capex_to_revenue": ("capex", "revenue"),
    "gross_margin": ("gross_profit", "revenue"),
    "operating_margin": ("operating_income", "revenue"),
    "ebitda_margin": ("ebitda", "revenue"),
    "fcf_margin": ("free_cash_flow", "revenue"),
}
GROWTH_SOURCES = {
    "revenue_growth": "revenue",
    "gross_profit_growth": "gross_profit",
    "employee_growth": "employees",
    "operating_income_growth": "operating_income",
    "fcf_growth": "free_cash_flow",
    "revenue_per_employee_growth": "revenue_per_employee",
    "gross_profit_per_employee_growth": "gross_profit_per_employee",
    "operating_income_per_employee_growth": "operating_income_per_employee",
    "fcf_per_employee_growth": "fcf_per_employee",
}
LEVERAGE_SOURCES = {
    "labor_leverage": "revenue_growth",
    "gross_profit_labor_leverage": "gross_profit_growth",
    "fcf_labor_leverage": "fcf_growth",
}
PERCENTAGE_POINT_METRICS = set(RATIO_SOURCES) | set(GROWTH_SOURCES) | set(LEVERAGE_SOURCES)
METRIC_UNITS = {
    **{name: "USD per employee" for name in PER_EMPLOYEE_SOURCES},
    **{name: "fraction" for name in RATIO_SOURCES},
    **{name: "fraction" for name in GROWTH_SOURCES},
    **{
        name: "growth-fraction difference; multiply by 100 for percentage points" for name in LEVERAGE_SOURCES
    },
    "delta_revenue_to_rd": "USD revenue change per USD R&D expense (multiple)",
}


def validate_company_years(panel: pd.DataFrame) -> pd.DataFrame:
    """Return an independent, sorted panel with validated security/FY keys."""
    missing = {"security", "fiscal_year"}.difference(panel.columns)
    if missing:
        raise ValueError(f"Panel is missing required key columns: {sorted(missing)}")
    result = panel.copy(deep=True)
    if not result["security"].map(lambda item: isinstance(item, str) and bool(item.strip())).all():
        raise ValueError("Every company-year must have a non-empty security identifier")
    years = pd.to_numeric(result["fiscal_year"], errors="coerce")
    if (years.isna() | ~np.isfinite(years) | (years % 1 != 0) | (years < 1) | (years > 9999)).any():
        raise ValueError("fiscal_year must contain finite integer years between 1 and 9999")
    result["fiscal_year"] = years.astype("int64")
    duplicates = result.duplicated(["security", "fiscal_year"], keep=False)
    if duplicates.any():
        examples = result.loc[duplicates, ["security", "fiscal_year"]].head(6).to_dict("records")
        raise ValueError(f"Duplicate security/fiscal_year observations: {examples}")
    return result.sort_values(["security", "fiscal_year"], kind="stable").reset_index(drop=True)


def numeric_column(panel: pd.DataFrame, name: str) -> pd.Series:
    """Read a numeric calculation view without changing the raw input column."""
    if name not in panel:
        return pd.Series(np.nan, index=panel.index, dtype=float)
    values = pd.to_numeric(panel[name], errors="coerce").astype(float)
    return values.where(np.isfinite(values))


def _divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
        result = numerator.div(denominator.where(denominator > 0))
    return result.where(np.isfinite(result))


def transform_panel(panel: pd.DataFrame, monetary_multiplier: float = 1000000) -> pd.DataFrame:
    """Calculate productivity, margins, YoY growth, and labor leverage.

    Monetary source values must be in one verified currency (USD) and unit.
    ``monetary_multiplier`` converts that unit to dollars; use 1 for dollars
    or 1_000_000 for millions. It is never applied to employee counts.
    Negative profit/FCF levels remain meaningful numerators, but growth from
    a zero or negative base is undefined. Employee levels must be positive
    for both per-employee calculations and employee growth. Missing fiscal
    years break the growth chain; there is no interpolation or forward fill.
    Expense and capex signs are preserved as downloaded.
    Incremental revenue / R&D uses the current year's positive R&D expense
    and adjacent-year revenue difference, in matching monetary units. Negative
    revenue changes remain negative; this is not a causal return on R&D or AI.
    """
    if (
        isinstance(monetary_multiplier, bool)
        or not math.isfinite(monetary_multiplier)
        or monetary_multiplier <= 0
    ):
        raise ValueError("monetary_multiplier must be a finite positive number")
    result = validate_company_years(panel)
    employees = numeric_column(result, "employees").where(lambda values: values > 0)
    for metric, source in PER_EMPLOYEE_SOURCES.items():
        result[metric] = _divide(numeric_column(result, source) * monetary_multiplier, employees)
    for metric, (numerator, denominator) in RATIO_SOURCES.items():
        result[metric] = _divide(numeric_column(result, numerator), numeric_column(result, denominator))
    previous_year = result.groupby("security", sort=False)["fiscal_year"].shift()
    adjacent = result["fiscal_year"].sub(previous_year).eq(1)
    revenue = numeric_column(result, "revenue")
    revenue_delta = revenue.sub(revenue.groupby(result["security"], sort=False).shift()).where(adjacent)
    result["delta_revenue_to_rd"] = _divide(revenue_delta, numeric_column(result, "rd_expense"))
    for metric, source in GROWTH_SOURCES.items():
        current = employees if source == "employees" else numeric_column(result, source)
        previous = current.groupby(result["security"], sort=False).shift()
        growth = _divide(current.sub(previous), previous)
        result[metric] = growth.where(adjacent)
    for metric, growth in LEVERAGE_SOURCES.items():
        result[metric] = result[growth].sub(result["employee_growth"])
    result.attrs.update(panel.attrs)
    result.attrs["monetary_multiplier"] = monetary_multiplier
    result.attrs["metric_units"] = METRIC_UNITS.copy()
    return result


def select_software(
    panel: pd.DataFrame,
    terms: Iterable[str] = ("Software", "Application Software", "Systems Software", "IT Services"),
) -> pd.DataFrame:
    """Select literal, case-insensitive classification substrings.

    Current GICS metadata are usually applied to the complete history; this
    is a configurable descriptive subset, not historical GICS membership.
    Company names/tickers are deliberately excluded from the filter.
    """
    terms = tuple((term.strip().casefold() for term in terms if isinstance(term, str) and term.strip()))
    matched = pd.Series(False, index=panel.index)
    classifications = ("sector", "industry_group", "industry", "sub_industry")
    for column in classifications + tuple((f"gics_{column}" for column in classifications)):
        if column not in panel:
            continue
        values = panel[column].astype("string").str.casefold()
        for term in terms:
            matched |= values.str.contains(term, regex=False, na=False).astype(bool)
    result = panel.loc[matched].copy()
    result.attrs.update(panel.attrs)
    result.attrs["software_classification_terms"] = terms
    return result


# ========================================================================
# Descriptive comparisons and summaries
# ========================================================================

PRE_LABEL = "Pre-2023 period"
POST_LABEL = "Post-2022 period"
REGIME_METRICS = (
    "revenue_per_employee",
    "gross_profit_per_employee",
    "operating_income_per_employee",
    "fcf_per_employee",
    "personnel_cost_per_employee",
    "rd_to_revenue",
    "delta_revenue_to_rd",
    "sga_to_revenue",
    "personnel_expense_to_revenue",
    "operating_margin",
    "fcf_margin",
    "employee_growth",
    "labor_leverage",
)
IDENTIFIERS = ("security", "ticker", "company_name", "sector", "industry_group", "industry", "sub_industry")
NAMED_CHANGES = {
    "post_vs_pre_revenue_per_employee_change": ("revenue_per_employee", "relative_change"),
    "post_vs_pre_gp_per_employee_change": ("gross_profit_per_employee", "relative_change"),
    "post_vs_pre_labor_leverage_change": ("labor_leverage", "absolute_change"),
    "post_vs_pre_operating_margin_change": ("operating_margin", "absolute_change"),
    "post_vs_pre_fcf_margin_change": ("fcf_margin", "absolute_change"),
    "post_vs_pre_delta_revenue_to_rd_change": ("delta_revenue_to_rd", "absolute_change"),
}
SUMMARY_COLUMNS = (
    "metric",
    "unit",
    "observations",
    "universe_securities",
    "coverage_fraction",
    "mean",
    "median",
    "p25",
    "p75",
)


def _years(values: Iterable[int], label: str) -> tuple[int, ...]:
    result = tuple(values)
    if not result or any(
        (isinstance(year, bool) or not isinstance(year, (int, np.integer)) for year in result)
    ):
        raise ValueError(f"{label} must contain integer fiscal years")
    if len(set(result)) != len(result):
        raise ValueError(f"{label} contains duplicate fiscal years")
    return tuple(sorted(result))


def _company_comparisons(
    panel: pd.DataFrame,
    metrics: list[str],
    pre_years: tuple[int, ...],
    post_years: tuple[int, ...],
    minimum: int,
) -> pd.DataFrame:
    rows = []
    for security, company in panel.groupby("security", sort=True):
        row = {"security": security}
        for column in IDENTIFIERS[1:]:
            values = company[column].dropna() if column in company else pd.Series(dtype=object)
            row[column] = values.iloc[-1] if len(values) else np.nan
        row.update(
            pre_period_label=PRE_LABEL,
            post_period_label=POST_LABEL,
            pre_fiscal_years=",".join(map(str, pre_years)),
            post_fiscal_years=",".join(map(str, post_years)),
        )
        for metric in metrics:
            values = numeric_column(company, metric)
            pre = values.loc[company["fiscal_year"].isin(pre_years)].dropna()
            post = values.loc[company["fiscal_year"].isin(post_years)].dropna()
            supported = len(pre) >= minimum and len(post) >= minimum
            pre_mean, post_mean = (pre.mean(), post.mean())
            difference = post_mean - pre_mean if supported else np.nan
            row.update(
                {
                    f"{metric}_pre_mean": pre_mean,
                    f"{metric}_post_mean": post_mean,
                    f"{metric}_pre_n": len(pre),
                    f"{metric}_post_n": len(post),
                    f"{metric}_matched_support": supported,
                    f"{metric}_absolute_change": difference,
                }
            )
            if metric in PER_EMPLOYEE_SOURCES:
                row[f"{metric}_relative_change"] = (
                    difference / pre_mean if supported and pre_mean > 0 else np.nan
                )
            elif metric in PERCENTAGE_POINT_METRICS:
                row[f"{metric}_change_pp"] = difference * 100
        for name, (metric, kind) in NAMED_CHANGES.items():
            row[name] = row.get(f"{metric}_{kind}", np.nan)
            if kind == "absolute_change" and metric in PERCENTAGE_POINT_METRICS:
                row[f"{name}_pp"] = row[name] * 100
        rows.append(row)
    if rows:
        return pd.DataFrame(rows)
    columns = list(IDENTIFIERS) + [
        "pre_period_label",
        "post_period_label",
        "pre_fiscal_years",
        "post_fiscal_years",
    ]
    for metric in metrics:
        columns.extend(
            (
                f"{metric}_{suffix}"
                for suffix in (
                    "pre_mean",
                    "post_mean",
                    "pre_n",
                    "post_n",
                    "matched_support",
                    "absolute_change",
                )
            )
        )
        if metric in PER_EMPLOYEE_SOURCES:
            columns.append(f"{metric}_relative_change")
        elif metric in PERCENTAGE_POINT_METRICS:
            columns.append(f"{metric}_change_pp")
    columns.extend(NAMED_CHANGES)
    columns.extend(
        f"{name}_pp"
        for name, (metric, kind) in NAMED_CHANGES.items()
        if kind == "absolute_change" and metric in PERCENTAGE_POINT_METRICS
    )
    return pd.DataFrame(columns=columns)


def _cross_section(panel: pd.DataFrame, metrics: list[str], groups: list[str]) -> pd.DataFrame:
    rows = []
    for key, group in panel.groupby(groups, dropna=False, sort=True, observed=True):
        keys = key if isinstance(key, tuple) else (key,)
        universe = int(group["security"].nunique())
        for metric in metrics:
            values = numeric_column(group, metric).dropna()
            rows.append(
                {
                    **dict(zip(groups, keys)),
                    "metric": metric,
                    "unit": METRIC_UNITS[metric],
                    "observations": len(values),
                    "universe_securities": universe,
                    "coverage_fraction": len(values) / universe if universe else np.nan,
                    "mean": values.mean(),
                    "median": values.median(),
                    "p25": values.quantile(0.25),
                    "p75": values.quantile(0.75),
                }
            )
    return pd.DataFrame(rows, columns=groups + list(SUMMARY_COLUMNS))


def _comparison_summary(
    company: pd.DataFrame, metrics: list[str], quantiles: tuple[float, float] | None
) -> pd.DataFrame:
    rows = []
    for metric in metrics:
        kinds = ["absolute_change"] + (["relative_change"] if metric in PER_EMPLOYEE_SOURCES else [])
        for kind in kinds:
            values = numeric_column(company, f"{metric}_{kind}").dropna()
            lower = values.quantile(quantiles[0]) if quantiles and len(values) else np.nan
            upper = values.quantile(quantiles[1]) if quantiles and len(values) else np.nan
            sample = values.clip(lower=lower, upper=upper) if quantiles else values
            rows.append(
                {
                    "metric": metric,
                    "change_kind": kind,
                    "unit": "fraction" if kind == "relative_change" else METRIC_UNITS[metric],
                    "matched_securities": len(values),
                    "universe_securities": len(company),
                    "matched_coverage_fraction": len(values) / len(company) if len(company) else np.nan,
                    "mean": sample.mean(),
                    "median": sample.median(),
                    "p25": sample.quantile(0.25),
                    "p75": sample.quantile(0.75),
                    "winsorized": quantiles is not None,
                    "lower_clip_value": lower,
                    "upper_clip_value": upper,
                    "winsor_lower_quantile": quantiles[0] if quantiles else np.nan,
                    "winsor_upper_quantile": quantiles[1] if quantiles else np.nan,
                }
            )
    return pd.DataFrame(rows)


def _definitions(metrics: list[str]) -> pd.DataFrame:
    rows = [
        {
            "column_or_pattern": metric,
            "unit": METRIC_UNITS[metric],
            "definition": (
                "(Revenue in FY t minus revenue in FY t-1) / R&D expense in FY t. "
                "Requires adjacent fiscal years, both revenue values, and positive current R&D. "
                "Matching monetary units cancel; negative/zero revenue changes are preserved. "
                "A descriptive multiple, not causal R&D returns or proof of AI impact."
                if metric == "delta_revenue_to_rd"
                else "Per security/fiscal year. Source signs preserved; no imputation or clipping."
            ),
        }
        for metric in metrics
    ]
    rows.extend(
        [
            {
                "column_or_pattern": "*_pre_mean / *_post_mean",
                "unit": "same as underlying metric",
                "definition": "Arithmetic mean of non-null fiscal-year observations within the named period.",
            },
            {
                "column_or_pattern": "*_pre_n / *_post_n",
                "unit": "fiscal-year observations",
                "definition": "Metric-specific non-null count; no common balanced panel is imposed across metrics.",
            },
            {
                "column_or_pattern": "*_matched_support",
                "unit": "boolean",
                "definition": "Both periods meet minimum observations for this security and metric.",
            },
            {
                "column_or_pattern": "*_absolute_change",
                "unit": "same as underlying metric",
                "definition": "Post-period mean minus pre-period mean, only with matched support.",
            },
            {
                "column_or_pattern": "*_relative_change",
                "unit": "fraction",
                "definition": "Post/pre - 1 for per-employee metrics; requires positive pre-period mean and matched support.",
            },
            {
                "column_or_pattern": "*_change_pp",
                "unit": "percentage points",
                "definition": "100 times post-minus-pre difference for ratios, growth fractions and labor leverage.",
            },
        ]
    )
    for name, (metric, kind) in NAMED_CHANGES.items():
        rows.append(
            {
                "column_or_pattern": name,
                "unit": "fraction" if kind == "relative_change" else METRIC_UNITS[metric],
                "definition": f"Alias of {metric}_{kind}.",
            }
        )
    return pd.DataFrame(rows)


def analyze(
    panel: pd.DataFrame,
    output_dir: Path | str,
    pre_years: Iterable[int] = (2017, 2018, 2019, 2020, 2021, 2022),
    post_years: Iterable[int] = (2023, 2024, 2025),
    winsor_quantiles: tuple[float, float] | None = (0.01, 0.99),
    min_period_observations: int = 2,
) -> dict[str, pd.DataFrame]:
    """Save/return untrimmed company, year, sector, and coverage tables.

    Optional winsorization applies ONLY to a separately named summary of
    company-level changes. Input data, period means, and all primary outputs
    remain untrimmed. A security is the observational unit: distinct share
    classes are not deduplicated without a verified issuer mapping.
    """
    pre_years, post_years = (_years(pre_years, "pre_years"), _years(post_years, "post_years"))
    if set(pre_years) & set(post_years):
        raise ValueError("Pre- and post-period fiscal years must not overlap")
    if (
        isinstance(min_period_observations, bool)
        or not isinstance(min_period_observations, int)
        or min_period_observations < 1
    ):
        raise ValueError("min_period_observations must be a positive integer")
    if winsor_quantiles is not None:
        if len(winsor_quantiles) != 2 or not 0 <= winsor_quantiles[0] < winsor_quantiles[1] <= 1:
            raise ValueError("winsor_quantiles must satisfy 0 <= lower < upper <= 1, or be None")
    data = validate_company_years(panel)
    if "sector" not in data:
        data["sector"] = np.nan
    metrics = list(dict.fromkeys([*REGIME_METRICS, *(metric for metric in METRIC_UNITS if metric in data)]))
    company = _company_comparisons(data, metrics, pre_years, post_years, min_period_observations)
    annual = _cross_section(data, metrics, ["fiscal_year"])
    sector = _cross_section(data, metrics, ["fiscal_year", "sector"])
    annual_coverage = annual[
        ["fiscal_year", "metric", "observations", "universe_securities", "coverage_fraction"]
    ].copy()
    annual_coverage.insert(0, "scope", "All securities")
    annual_coverage.insert(2, "sector", pd.NA)
    sector_coverage = sector[
        ["fiscal_year", "sector", "metric", "observations", "universe_securities", "coverage_fraction"]
    ].copy()
    sector_coverage.insert(0, "scope", "Sector")
    tables = {
        "company_regime_comparison": company,
        "annual_summary": annual,
        "sector_summary": sector,
        "coverage_summary": pd.concat([annual_coverage, sector_coverage], ignore_index=True),
        "comparison_summary": _comparison_summary(company, metrics, None),
        "metric_definitions": _definitions(metrics),
        "analysis_notes": pd.DataFrame(
            [
                {
                    "topic": "interpretation",
                    "note": "Descriptive association only. Period differences are not causal estimates of AI productivity or AI impact.",
                },
                {
                    "topic": "universe",
                    "note": "Current index securities, not historical constituents. Survivorship bias and historical GICS reclassification can affect comparisons.",
                },
                {
                    "topic": "observational_unit",
                    "note": "Securities; multiple share classes of one issuer can count more than once. Counts must not be interpreted as deduplicated issuers.",
                },
                {
                    "topic": "aggregation",
                    "note": "Equal-security cross-sectional statistics of company metrics; not ratios of aggregate monetary values and aggregate employees.",
                },
                {
                    "topic": "coverage",
                    "note": "Non-null values divided by securities present in each fiscal-year/group. Coverage does not include securities absent from the input panel.",
                },
                {
                    "topic": "period_support",
                    "note": f"At least {min_period_observations} observed fiscal years per period and metric; all available years contribute to each period mean. Fiscal years need not share calendar dates.",
                },
                {
                    "topic": "growth_boundary",
                    "note": "Growth is assigned to its ending fiscal year: 2023 growth uses the 2022 base even though 2022 belongs to the earlier period.",
                },
                {
                    "topic": "winsorization",
                    "note": "Only winsorized_comparison_summary clips supported company-level differences across securities. Primary tables and company means remain untrimmed.",
                },
                {
                    "topic": "accounting",
                    "note": "Source expense/capex signs are preserved. Acquisitions, disposals, outsourcing, wages, accounting and reporting changes can affect these metrics.",
                },
            ]
        ),
    }
    if winsor_quantiles is not None:
        tables["winsorized_comparison_summary"] = _comparison_summary(company, metrics, winsor_quantiles)
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    for name, table in tables.items():
        write_frame(table, directory / f"{name}.csv")
    return tables


# ========================================================================
# Charts
# ========================================================================

plt: Any = None
FuncFormatter: Any = None

NAVY = "#22445c"
TEAL = "#138a86"
GRAY = "#657380"
STYLE = {
    "font.family": "sans-serif",
    "font.size": 10,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.edgecolor": "#bcc5cc",
    "axes.labelcolor": NAVY,
    "axes.titleweight": "bold",
    "text.color": NAVY,
    "xtick.color": GRAY,
    "ytick.color": GRAY,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "savefig.facecolor": "white",
}


def _finish(fig: plt.Figure, path: Path, footer: str) -> None:
    fig.text(0.07, 0.045, footer, ha="left", va="bottom", fontsize=8, color=GRAY)
    fig.tight_layout(rect=(0, 0.15, 1, 0.94))
    fig.savefig(path, dpi=180, bbox_inches="tight")
    fig.savefig(path.with_suffix(".svg"), bbox_inches="tight")
    plt.close(fig)


def _line(ax: plt.Axes, panel: pd.DataFrame, metric: str, label: str, color: str, factor: float = 1) -> bool:
    values = numeric_column(panel, metric)
    frame = pd.DataFrame({"year": panel["fiscal_year"], "value": values * factor})
    medians = frame.groupby("year")["value"].median()
    if medians.notna().sum() == 0:
        return False
    medians = medians.reindex(range(int(frame["year"].min()), int(frame["year"].max()) + 1))
    ax.plot(medians.index, medians.values, color=color, marker="o", markersize=3.5, linewidth=2, label=label)
    return True


def _style_axis(ax: plt.Axes, panel: pd.DataFrame, post_years: tuple[int, ...], has_data: bool) -> None:
    ax.grid(axis="y", color="#e4e9ed", linewidth=0.7)
    ax.set_axisbelow(True)
    if not panel.empty:
        first, last = (int(panel["fiscal_year"].min()), int(panel["fiscal_year"].max()))
        ax.set_xlim(first - 0.35, last + 0.35)
        if first <= max(post_years) and last >= min(post_years):
            ax.axvspan(
                max(first - 0.35, min(post_years) - 0.5),
                min(last + 0.35, max(post_years) + 0.5),
                color=TEAL,
                alpha=0.055,
                linewidth=0,
            )
        ticks = list(range(first, last + 1, 2 if last - first > 8 else 1))
        ax.set_xticks(ticks)
    if has_data:
        ax.legend(frameon=False, fontsize=9, loc="best")
    else:
        ax.text(
            0.5,
            0.5,
            "No valid observations available",
            ha="center",
            va="center",
            transform=ax.transAxes,
            color=GRAY,
        )
    ax.set_xlabel("Fiscal year")


def create_charts(
    panel: pd.DataFrame,
    software_panel: pd.DataFrame,
    output_dir: Path | str,
    pre_years: Iterable[int] = (2017, 2018, 2019, 2020, 2021, 2022),
    post_years: Iterable[int] = (2023, 2024, 2025),
    min_period_observations: int = 2,
) -> dict[str, Path]:
    """Create untrimmed trends and matched-period Software comparison.

    No observations produces an annotated figure, never a fabricated series.
    Return PNG paths; editable SVG versions are written alongside each PNG.
    """
    global plt, FuncFormatter
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

    data, software = (validate_company_years(panel), validate_company_years(software_panel))
    pre, post = (_years(pre_years, "pre_years"), _years(post_years, "post_years"))
    if set(pre) & set(post):
        raise ValueError("Pre- and post-period fiscal years must not overlap")
    if (
        isinstance(min_period_observations, bool)
        or not isinstance(min_period_observations, int)
        or min_period_observations < 1
    ):
        raise ValueError("min_period_observations must be a positive integer")
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    result = {
        name: directory / f"{name}.png"
        for name in ("productivity_trends", "labor_leverage_trends", "software_productivity_scatter")
    }
    universe_n, software_n = (data["security"].nunique(), software["security"].nunique())
    footnote = f"Equal-security medians; available observations, untrimmed. Universe: {universe_n} securities; Software: {software_n}.\nCoverage varies by metric/year; share classes may repeat issuers. Shading marks the comparison period. Descriptive associations only."
    with plt.rc_context(STYLE):
        fig, axes = plt.subplots(1, 2, figsize=(11.4, 5.5))
        fig.suptitle("Output per employee", x=0.07, ha="left", fontsize=17, fontweight="bold")
        for ax, metric, title in zip(
            axes, ("revenue_per_employee", "gross_profit_per_employee"), ("Revenue", "Gross profit")
        ):
            shown = _line(ax, data, metric, "All securities", NAVY)
            shown = _line(ax, software, metric, "Software", TEAL) or shown
            ax.set_title(title, loc="left", pad=12)
            ax.yaxis.set_major_formatter(FuncFormatter(lambda value, position: f"${value / 1000:,.0f}k"))
            ax.set_ylabel("USD per employee")
            _style_axis(ax, data, post, shown)
        _finish(fig, result["productivity_trends"], footnote)
        fig, axes = plt.subplots(1, 2, figsize=(11.4, 5.5))
        fig.suptitle("Labor leverage", x=0.07, ha="left", fontsize=17, fontweight="bold")
        for ax, metric, title in zip(
            axes,
            ("labor_leverage", "gross_profit_labor_leverage"),
            ("Revenue growth less headcount growth", "Gross-profit growth less headcount growth"),
        ):
            shown = _line(ax, data, metric, "All securities", NAVY, 100)
            shown = _line(ax, software, metric, "Software", TEAL, 100) or shown
            ax.set_title(title, loc="left", pad=12, fontsize=11)
            ax.axhline(0, color=GRAY, linewidth=0.7, alpha=0.7)
            ax.set_ylabel("Percentage points")
            _style_axis(ax, data, post, shown)
        _finish(fig, result["labor_leverage_trends"], footnote)
        comparisons = _company_comparisons(
            software, ["revenue_per_employee", "employee_growth"], pre, post, min_period_observations
        )
        x = numeric_column(comparisons, "employee_growth_absolute_change") * 100
        y = numeric_column(comparisons, "revenue_per_employee_relative_change") * 100
        valid = x.notna() & y.notna() & np.isfinite(x) & np.isfinite(y)
        fig, ax = plt.subplots(figsize=(8.8, 6.8))
        fig.suptitle(
            "Software: productivity and headcount growth", x=0.09, ha="left", fontsize=16, fontweight="bold"
        )
        ax.set_title(
            f"Fiscal years {min(post)}–{max(post)} versus {min(pre)}–{max(pre)}",
            loc="left",
            color=GRAY,
            fontweight="normal",
            fontsize=10,
        )
        ax.axhline(0, color=GRAY, linewidth=0.8)
        ax.axvline(0, color=GRAY, linewidth=0.8)
        ax.grid(color="#e4e9ed", linewidth=0.7)
        ax.set_axisbelow(True)
        if valid.any():
            ax.scatter(
                x.loc[valid], y.loc[valid], color=TEAL, s=46, alpha=0.8, edgecolors="white", linewidth=0.7
            )
            if valid.sum() <= 18:
                for index in comparisons.index[valid]:
                    ticker = comparisons.loc[index, "ticker"]
                    label = (
                        ticker
                        if pd.notna(ticker)
                        else comparisons.loc[index, "security"].removesuffix(" Equity")
                    )
                    ax.annotate(
                        str(label),
                        (x.loc[index], y.loc[index]),
                        xytext=(5, 5),
                        textcoords="offset points",
                        fontsize=8,
                    )
            ax.margins(x=0.15, y=0.2)
        else:
            ax.text(
                0.5,
                0.5,
                "No Software securities meet both metrics' period support",
                transform=ax.transAxes,
                ha="center",
                va="center",
                color=GRAY,
            )
        ax.set_xlabel("Change in mean headcount growth (percentage points)")
        ax.set_ylabel("Change in mean revenue per employee (%)")
        _finish(
            fig,
            result["software_productivity_scatter"],
            f"Matched securities: {int(valid.sum())}; at least {min_period_observations} observations per metric in each period. Untrimmed.\nPeriod means use available years; revenue per employee requires a positive earlier-period mean.\nCurrent classifications; share classes may repeat issuers. Descriptive association, not a causal estimate.",
        )
    return result


# ========================================================================
# Command-line workflow
# ========================================================================


def run_pipeline(client, settings: Settings, *, validate_only: bool = False) -> dict:
    """Run using an already connected client (also allows offline integration tests)."""
    settings.make_dirs()
    report = validate_fields(client, settings.diagnostics_dir)
    summary = {
        "annual_verified": report.annual_verified,
        "snapshot_verified": report.snapshot_verified,
        "fiscal_period": report.fiscal_period,
    }
    if validate_only:
        return {**summary, "status": "validation_only", "annual_ready": bool(report.fiscal_period)}
    if not report.fiscal_period:
        raise ValueError(
            "Neither FUND_PER=Y nor FY passed fiscal identity checks. See output/diagnostics; inspect candidate fields and their overrides in Bloomberg FLDS before a full pull."
        )
    members = fetch_universe(client, report.classifications, test=settings.test)
    write_frame(members, settings.raw_dir / "spx_members.csv", settings.parquet)
    print_classifications(members)
    try:
        fundamentals = pull_annual_panel(client, members, report, settings)
        panel = transform_panel(fundamentals, monetary_multiplier=MONETARY_MULTIPLIER)
        software = select_software(panel, terms=SOFTWARE_TERMS)
        snapshot = pull_current_snapshot(client, members, report, settings)
    finally:
        report.save(settings.diagnostics_dir)
    write_frame(panel, settings.processed_dir / "company_fundamentals_panel.csv", settings.parquet)
    write_frame(software, settings.processed_dir / "software_panel.csv", settings.parquet)
    write_frame(snapshot, settings.processed_dir / "current_snapshot.csv", settings.parquet)
    write_frame(quality_diagnostics(panel), settings.diagnostics_dir / "data_quality.csv")
    analysis_options = dict(
        pre_years=PRE_AI_YEARS,
        post_years=POST_AI_YEARS,
        winsor_quantiles=WINSOR_QUANTILES,
        min_period_observations=MIN_PERIOD_OBSERVATIONS,
    )
    analyze(panel, settings.output_dir / "output" / "tables", **analysis_options)
    analyze(software, settings.output_dir / "output" / "tables" / "software", **analysis_options)
    os.environ.setdefault("MPLCONFIGDIR", str((settings.output_dir / "output" / ".matplotlib").resolve()))
    charts = create_charts(
        panel,
        software,
        settings.output_dir / "output" / "charts",
        pre_years=PRE_AI_YEARS,
        post_years=POST_AI_YEARS,
        min_period_observations=MIN_PERIOD_OBSERVATIONS,
    )
    LOGGER.info(
        "Completed: %s securities, %s annual rows, %s software rows; output %s",
        len(members),
        len(panel),
        len(software),
        settings.output_dir.resolve(),
    )
    return {
        **summary,
        "status": "complete",
        "securities": len(members),
        "panel_rows": len(panel),
        "software_rows": len(software),
        "charts": {k: str(v) for k, v in charts.items()},
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Bloomberg fiscal actuals and descriptive labor leverage research"
    )
    parser.add_argument(
        "--test", action="store_true", help="Only MSFT, CRM, ADBE, NOW, ORCL; default output runs/test"
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Save live field/override diagnostics without a universe pull",
    )
    parser.add_argument(
        "--output-dir", type=Path, help="Root for data/ and output/ (default: .; --test: runs/test)"
    )
    parser.add_argument("--host", default="localhost")
    parser.add_argument("--port", type=int, default=8194)
    parser.add_argument("--batch-size", type=int, default=100)
    parser.add_argument("--timeout", type=float, default=60.0, help="Per-request timeout in seconds")
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--no-parquet", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    settings = Settings(
        output_dir=(args.output_dir or (Path("runs/test") if args.test else Path("."))).resolve(),
        host=args.host,
        port=args.port,
        batch_size=args.batch_size,
        timeout_seconds=args.timeout,
        retries=args.retries,
        test=args.test,
        parquet=not args.no_parquet,
    )
    settings.make_dirs()
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(settings.diagnostics_dir / "pipeline.log", encoding="utf-8"),
        ],
    )
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "_" + uuid4().hex[:8]
    manifest_path = settings.diagnostics_dir / "run_manifest.json"
    lock_path = settings.output_dir / ".pipeline.lock"
    try:
        lock_fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        LOGGER.error(
            "Output directory is locked by another run: %s. Check the process before removing a stale lock.",
            lock_path,
        )
        return 2
    with os.fdopen(lock_fd, "w") as handle:
        handle.write(f"pid={os.getpid()} run_id={run_id}\n")
    versions = {}
    for name in ("pandas", "numpy", "matplotlib", "pyarrow", "blpapi"):
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            versions[name] = "not installed"
    manifest = {
        "run_id": run_id,
        "status": "running",
        "started_at_utc": datetime.now(timezone.utc).isoformat(),
        "version": __version__,
        "python": sys.version,
        "platform": platform.platform(),
        "dependencies": versions,
        "settings": asdict(settings),
        "start_fy": START_FY,
        "end_fy": END_FY,
        "pre_years": PRE_AI_YEARS,
        "post_years": POST_AI_YEARS,
        "software_terms": SOFTWARE_TERMS,
        "universe_basis": "five-security debug set" if settings.test else "current SPX Index INDX_MEMBERS",
        "interpretation": "Descriptive; current membership and GICS, potentially revised actuals; no causal AI estimate",
    }
    try:
        write_json(manifest, manifest_path)
        with BloombergClient(
            host=settings.host,
            port=settings.port,
            batch_size=settings.batch_size,
            timeout_seconds=settings.timeout_seconds,
            retries=settings.retries,
            raw_dir=settings.raw_dir / "requests" / run_id,
        ) as client:
            result = run_pipeline(client, settings, validate_only=args.validate_only)
        manifest.update(result)
        manifest["finished_at_utc"] = datetime.now(timezone.utc).isoformat()
        write_json(manifest, manifest_path)
        return 0
    except (BloombergError, ValueError, OSError, ImportError) as exc:
        manifest.update(
            status="failed", error=str(exc), finished_at_utc=datetime.now(timezone.utc).isoformat()
        )
        write_json(manifest, manifest_path)
        LOGGER.error("Pipeline stopped: %s", exc)
        return 1
    except BaseException as exc:
        manifest.update(
            status="interrupted" if isinstance(exc, KeyboardInterrupt) else "failed",
            error=str(exc),
            finished_at_utc=datetime.now(timezone.utc).isoformat(),
        )
        write_json(manifest, manifest_path)
        raise
    finally:
        lock_path.unlink(missing_ok=True)


if __name__ == "__main__":
    raise SystemExit(main())
