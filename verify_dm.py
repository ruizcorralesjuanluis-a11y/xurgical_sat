import sys
import os
import time
import json
import argparse
import threading

# Añadir el directorio 'reader' al path para importar CameraService y SerialService
sys.path.append(os.path.join(os.path.dirname(__file__), 'reader'))

from camera_service import CameraService, CameraSettings
from serial_service import SerialService, SerialSettings, list_serial_ports_ebuho

def verify(expected_dm, timeout=15):
    result = {
        "status": "timeout",
        "scanned": None,
        "match": False,
        "error": None
    }
    
    found_event = threading.Event()
    
    # Inicializar Serial para feedback físico (buzzer/led)
    ebuho_ports = list_serial_ports_ebuho()
    serial_srv = None
    if ebuho_ports:
        try:
            serial_srv = SerialService(port=ebuho_ports[0], settings=SerialSettings())
            serial_srv.start()
        except Exception:
            serial_srv = None

    def on_decoded(text, fmt):
        if fmt == "DATAMATRIX":
            result["scanned"] = text
            matched = (text.strip() == expected_dm.strip())
            result["match"] = matched
            result["status"] = "success"
            
            if serial_srv:
                serial_srv.send_text('O' if matched else 'E')
            
            found_event.set()

    def on_fatal(msg):
        result["status"] = "error"
        result["error"] = msg
        if serial_srv:
            serial_srv.send_text('E')
        found_event.set()

    try:
        service = CameraService(
            settings=CameraSettings(width=1280, height=720, fps=10),
            on_decoded=on_decoded,
            on_fatal=on_fatal,
            allowed_cameras=["e-buho 4K camera"]
        )
        
        service.start()
        
        # Esperar hasta que se encuentre el código o pase el timeout
        finished = found_event.wait(timeout)
        
        if not finished:
            if serial_srv:
                serial_srv.send_text('E') # Error por timeout

        service.stop()
        service.close()
        
        if serial_srv:
            serial_srv.stop()
            serial_srv.close()
        
    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)
        if serial_srv:
            serial_srv.send_text('E')
            serial_srv.stop()
    
    return result

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Verifica un código DataMatrix usando el lector eBuho.')
    parser.add_argument('expected', help='El código DataMatrix esperado')
    parser.add_argument('--timeout', type=int, default=15, help='Tiempo máximo de espera (segundos)')
    
    args = parser.parse_args()
    
    final_result = verify(args.expected, args.timeout)
    print(json.dumps(final_result))
    sys.exit(0 if final_result["match"] else 1)
