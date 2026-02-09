
from db import get_conn

def migrate():
    print("Conectando a la base de datos para migración...")
    conn = get_conn()
    cur = conn.cursor()
    
    # Comprobamos si las columnas ya existen
    cur.execute("PRAGMA table_info(instrumentos)")
    columns = [row[1] if isinstance(row, tuple) else row["name"] for row in cur.fetchall()]
    
    if "tecnico_reparacion" not in columns:
        print("Añadiendo columna tecnico_reparacion...")
        cur.execute("ALTER TABLE instrumentos ADD COLUMN tecnico_reparacion TEXT")
    
    if "tecnico_reparacion_en" not in columns:
        print("Añadiendo columna tecnico_reparacion_en...")
        cur.execute("ALTER TABLE instrumentos ADD COLUMN tecnico_reparacion_en TEXT")
        
    conn.commit()
    conn.close()
    print("Migración finalizada correctamente.")

if __name__ == "__main__":
    migrate()
