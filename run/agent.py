# -*- coding: utf-8 -*-

from contextlib import contextmanager
import threading

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
    try:
        conn = getattr(_THREAD_STATE, "conn", None)
        if conn is not None:
            return pd.read_sql(sql, conn)
        with _get_conn() as new_conn:
            return pd.read_sql(sql, new_conn)
    except Exception as e:
        print(f"run_query hatası: {e}")
        return pd.DataFrame()


class Agent:
    run_query = staticmethod(run_query)


ag = Agent()
