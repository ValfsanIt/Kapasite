# -*- coding: utf-8 -*-
"""
Yerel SQL agent – config'deki veritabanına bağlanıp sorgu çalıştırır.
Bağımsız KAPASITE için valfapp/run bağımlılığı yerine kullanılır.
"""
import pandas as pd

try:
    import pyodbc
except ImportError:
    pyodbc = None


def _get_conn():
    """config'ten connection string ile bağlantı döndürür."""
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


def run_query(sql):
    """
    SQL sorgusunu çalıştırıp sonucu pandas DataFrame olarak döndürür.
    pyodbc yoksa veya bağlantı hatası olursa boş DataFrame döner.
    """
    if pyodbc is None:
        print("Uyarı: pyodbc yüklü değil. pip install pyodbc")
        return pd.DataFrame()
    try:
        with _get_conn() as conn:
            return pd.read_sql(sql, conn)
    except Exception as e:
        print(f"run_query hatası: {e}")
        return pd.DataFrame()


class Agent:
    run_query = staticmethod(run_query)


ag = Agent()
