import multiprocessing
import uvicorn
import webbrowser
import threading
import time
import os
import sys


def open_browser():
    """
    Espera un poco a que Uvicorn arranque
    y abre el navegador automáticamente
    """
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    try:
        import multiprocessing
        multiprocessing.freeze_support()

        print(">>> Cargando aplicación...")
        # Importa la app FastAPI dentro del try para capturar errores de importación
        from app import app

        # Abrir navegador en segundo plano (no bloquea el servidor)
        print(">>> Iniciando navegador...")
        threading.Thread(target=open_browser, daemon=True).start()

        # Arrancar servidor FastAPI
        print(">>> Iniciando servidor...")
        uvicorn.run(
            app,
            host="127.0.0.1",
            port=8000,
            log_level="info"
        )
    except Exception as e:
        import traceback
        print("\n" + "="*50)
        print("CRITICAL ERROR DURING STARTUP:")
        print("="*50)
        traceback.print_exc()
        print("="*50)
        input("\nPresiona ENTER para cerrar esta ventana...")
