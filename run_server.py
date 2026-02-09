import multiprocessing
import uvicorn
import webbrowser
import threading
import time

# Importa la app FastAPI
from app import app


def open_browser():
    """
    Espera un poco a que Uvicorn arranque
    y abre el navegador automáticamente
    """
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    multiprocessing.freeze_support()

    # Abrir navegador en segundo plano (no bloquea el servidor)
    threading.Thread(target=open_browser, daemon=True).start()

    # Arrancar servidor FastAPI
    uvicorn.run(
        app,
        host="127.0.0.1",
        port=8000,
        log_level="info"
    )
