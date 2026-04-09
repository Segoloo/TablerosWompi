import pymysql, time

conn = pymysql.connect(
    host="100.99.250.115", port=3306,
    user="root", password="An4l1t1c4l1n34*",
    database="lineacom_analitica",
    connect_timeout=30,
    read_timeout=3600,
    write_timeout=3600,
)

indices = [
    ("tracking_wompi_vp_rollos",              "idx_proyecto", "Nombre_del_proyecto(20)"),
    ("indicador_rollos_wompi_completo_raw",   "idx_tarea",    "tarea"),
    ("data_linea_todos",                      "idx_red_tipo", "red_asociada, tipo_actividad"),
]

with conn.cursor() as cur:
    for tabla, nombre, columna in indices:
        print(f"Creando {nombre} en {tabla}...")
        t0 = time.time()
        try:
            cur.execute(f"ALTER TABLE {tabla} ADD INDEX {nombre} ({columna})")
            print(f"  ✅ listo en {time.time()-t0:.1f}s")
        except Exception as e:
            print(f"  ⚠️  {e}")
conn.close()