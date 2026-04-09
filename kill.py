#!/usr/bin/env python3
"""
kill_locks.py — Limpia el deadlock de tracking_wompi_vp_rollos
Corre ANTES de cualquier extractor o ALTER TABLE.
"""
import pymysql

DB_HOST     = "100.99.250.115"
DB_PORT     = 3306
DB_USER     = "root"
DB_PASSWORD = "An4l1t1c4l1n34*"
DB_NAME     = "lineacom_analitica"

# PIDs confirmados del diagnóstico — el SELECT largo que tiene el lock
# y todos los que están esperando detrás de él.
PIDS_TO_KILL = [
    179456,   # SELECT ejecutando 4898s — EL QUE TIENE EL LOCK (matar primero)
    179671,   # DROP TABLE esperando metadata lock
    180023,   # SELECT esperando metadata lock
    180080,   # SELECT esperando metadata lock
    179357,   # root 5273s sin info — probablemente conexión zombie
    179358,   # root 5258s
    179359,   # root 5111s
    179360,   # root 5272s
    179361,   # root 5272s
    179362,   # root 5272s
    179363,   # root 5272s
    179364,   # root 5272s
    179365,   # root 5272s
    179449,   # root 192s
    179450,   # ALTER TABLE esperando metadata lock
    180379,   # ALTER TABLE esperando metadata lock
    180401,   # ALTER TABLE esperando metadata lock
    180429,   # ALTER TABLE esperando metadata lock
    180470,   # ALTER TABLE esperando metadata lock
    180542,   # root 191s
]

def main():
    conn = pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER,
        password=DB_PASSWORD, database=DB_NAME,
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
        connect_timeout=15,
    )
    print("✅ Conectado. Matando procesos bloqueados...\n")

    with conn.cursor() as cur:
        # Ver estado actual antes de matar
        cur.execute("SHOW FULL PROCESSLIST")
        activos = {r["Id"]: r for r in cur.fetchall()}

        killed  = []
        skipped = []

        for pid in PIDS_TO_KILL:
            if pid in activos:
                estado = activos[pid].get("State", "")
                tiempo = activos[pid].get("Time", 0)
                try:
                    cur.execute(f"KILL {pid}")
                    print(f"  ✓ KILL {pid} | Time={tiempo}s | State='{estado}'")
                    killed.append(pid)
                except Exception as e:
                    print(f"  ⚠ No se pudo matar {pid}: {e}")
            else:
                skipped.append(pid)
                print(f"  — PID {pid} ya no existe (ya terminó)")

    conn.close()

    print(f"\n{'='*50}")
    print(f"  Matados:  {len(killed)} procesos")
    print(f"  Ya muertos: {len(skipped)} procesos")
    print(f"{'='*50}")
    print("\n✅ Listo. Espera 5 segundos y corre:")
    print("   python rollos_extractor_v2.py --no-push")
    print("\n⚠  NO corras más ALTER TABLE — el extractor v2 no los necesita.")

if __name__ == "__main__":
    main()