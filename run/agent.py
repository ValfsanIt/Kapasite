# -*- coding: utf-8 -*-

from contextlib import contextmanager
import threading
import time
import warnings

import pandas as pd

try:
    import pyodbc
except ImportError:
    pyodbc = None


def _get_conn():
    
    import config
    drivers = [
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 18 for SQL Server",
        "SQL Server",
    ]
    for driver in drivers:
        try:
            conn_str = (
                f"DRIVER={{{driver}}};"
                f"SERVER={config.server};"
                f"DATABASE={config.database};"
                f"UID={config.username};"
                f"PWD={config.password}"
            )
            return pyodbc.connect(conn_str)
        except pyodbc.Error:
            continue
    raise RuntimeError("SQL Server ODBC sürücüsü bulunamadı. ODBC Driver 17/18 for SQL Server yükleyin.")


_THREAD_STATE = threading.local()


@contextmanager
def use_connection(conn=None):
    
    opened_here = False
    active = conn
    if pyodbc is None:
        yield None
        return
    if active is None:
        active = _get_conn()
        opened_here = True
    prev = getattr(_THREAD_STATE, "conn", None)
    _THREAD_STATE.conn = active
    try:
        yield active
    finally:
        _THREAD_STATE.conn = prev
        if opened_here:
            try:
                active.close()
            except Exception:
                pass


def run_query(sql):
    """
    SQL sorgusunu çalıştırıp sonucu pandas DataFrame olarak döndürür.
    pyodbc yoksa veya bağlantı hatası olursa boş DataFrame döner.
    """
    if pyodbc is None:
        print("Uyarı: pyodbc yüklü değil. pip install pyodbc")
        return pd.DataFrame()
    def _read_sql_frame(c):
        # pandas warns on every pyodbc read_sql; DBAPI2 is fine in practice.
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message=".*pandas only supports SQLAlchemy connectable.*",
                category=UserWarning,
            )
            return pd.read_sql(sql, c)

    def _looks_transient_db_error(err):
        msg = str(err).lower()
        transient_tokens = (
            "08s01",
            "dbnetlib",
            "communication link failure",
            "connectionread",
            "genel ağ hatası",
            "network-related",
            "connection is busy",
        )
        return any(t in msg for t in transient_tokens)

    max_attempts = 3
    for attempt in range(1, max_attempts + 1):
        conn = getattr(_THREAD_STATE, "conn", None)
        try:
            if conn is not None:
                return _read_sql_frame(conn)
            with _get_conn() as new_conn:
                return _read_sql_frame(new_conn)
        except Exception as e:
            is_transient = _looks_transient_db_error(e)
            print(f"run_query hatası (deneme {attempt}/{max_attempts}): {e}")
            if conn is not None:
                # use_connection içindeki paylaşılan bağlantı bozulduysa, tekrar denemede taze bağlantı zorla.
                _THREAD_STATE.conn = None
            if (not is_transient) or attempt >= max_attempts:
                return pd.DataFrame()
            time.sleep(0.25 * attempt)
    return pd.DataFrame()


class Agent:
    run_query = staticmethod(run_query)


ag = Agent()
