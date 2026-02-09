
import sqlite3
import os
from pathlib import Path

# Configuración de base de datos
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = os.environ.get("XURGICAL_DB_PATH", str(BASE_DIR / "sat.db"))

def seed_checklist():
    print(f"Conectando a {DB_PATH}...")
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # Checklist por defecto para REPARACION
    reparacion_defaults = [
        "Limpieza Ultrasónica",
        "Desmontaje y Revisión",
        "Afilado / Pulido",
        "Ajuste de Tornillería",
        "Engrase (si procede)",
        "Prueba de Corte / Funcionalidad",
        "Verificación de Aislamiento"
    ]

    # Checklist por defecto para OPTICA_RIGIDA
    optica_defaults = [
        "Limpieza de Lentes Externas",
        "Prueba de Estanqueidad",
        "Verificación de Transmisión lumínica",
        "Alineación de Fibra Óptica",
        "Revisión de Envolvente y Sellos"
    ]

    # Insertar REPARACION
    print("Insertando checks de Reparación...")
    for i, nombre in enumerate(reparacion_defaults):
        cur.execute("""
            INSERT INTO checklist_items (nombre, orden, activo, tipo_trabajo)
            VALUES (?, ?, 1, 'REPARACION')
        """, (nombre, (i + 1) * 10))

    # Insertar OPTICA_RIGIDA
    print("Insertando checks de Óptica Rígida...")
    for i, nombre in enumerate(optica_defaults):
        cur.execute("""
            INSERT INTO checklist_items (nombre, orden, activo, tipo_trabajo)
            VALUES (?, ?, 1, 'OPTICA_RIGIDA')
        """, (nombre, (i + 1) * 10))

    conn.commit()
    conn.close()
    print("¡Hecho! Checklist por defecto configurados.")

if __name__ == "__main__":
    seed_checklist()
