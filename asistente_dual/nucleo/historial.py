"""Historial local (SQLite) de consultas y llamadas, y control del tope de gasto mensual."""
import sqlite3
import threading
from datetime import datetime
from pathlib import Path

RUTA_DB = Path(__file__).resolve().parent.parent / "datos" / "historial.db"


class TopeExcedido(Exception):
    pass


class Historial:
    def __init__(self, ruta=RUTA_DB):
        Path(ruta).parent.mkdir(parents=True, exist_ok=True)
        self.con = sqlite3.connect(ruta, check_same_thread=False)
        self.lock = threading.Lock()  # el modo Comparar registra llamadas desde dos hilos
        self.con.executescript("""
            CREATE TABLE IF NOT EXISTS llamadas (fecha TEXT, proveedor TEXT, modelo TEXT,
                tokens_entrada INTEGER, tokens_salida INTEGER, costo_usd REAL);
            CREATE TABLE IF NOT EXISTS consultas (id INTEGER PRIMARY KEY, fecha TEXT, modo TEXT,
                pregunta TEXT, resultado_md TEXT, costo_usd REAL);
        """)

    def registrar_llamada(self, r):
        with self.lock, self.con:
            self.con.execute("INSERT INTO llamadas VALUES (?,?,?,?,?,?)",
                             (datetime.now().isoformat(timespec="seconds"), r.proveedor, r.modelo,
                              r.tokens_entrada, r.tokens_salida, r.costo_usd))

    def registrar_consulta(self, modo, pregunta, resultado_md, costo_usd):
        with self.lock, self.con:
            self.con.execute("INSERT INTO consultas (fecha, modo, pregunta, resultado_md, costo_usd) VALUES (?,?,?,?,?)",
                             (datetime.now().isoformat(timespec="seconds"), modo, pregunta, resultado_md, costo_usd))

    def gasto_mes(self, mes=None):
        mes = mes or datetime.now().strftime("%Y-%m")
        with self.lock:
            fila = self.con.execute("SELECT COALESCE(SUM(costo_usd),0) FROM llamadas WHERE substr(fecha,1,7)=?", (mes,)).fetchone()
        return fila[0]

    def verificar_tope(self, tope_usd, costo_estimado=0.0):
        gastado = self.gasto_mes()
        if gastado + costo_estimado > tope_usd:
            raise TopeExcedido(f"Tope mensual alcanzado: gastado USD {gastado:.2f} + estimado USD {costo_estimado:.2f} "
                               f"> tope USD {tope_usd:.2f}. Ajustá el tope en config.toml si corresponde.")

    def consultas(self, limite=50):
        cur = self.con.execute("SELECT id, fecha, modo, pregunta, resultado_md, costo_usd FROM consultas ORDER BY id DESC LIMIT ?", (limite,))
        return [dict(zip(("id", "fecha", "modo", "pregunta", "resultado_md", "costo_usd"), f)) for f in cur.fetchall()]
