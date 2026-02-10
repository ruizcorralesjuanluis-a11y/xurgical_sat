
import os
import sys

# Simular el entorno de la app
from app import load_articulos_map, _norm_codigo

print("Cargando catálogo...")
m = load_articulos_map()
print(f"Catálogo cargado. Total artículos: {len(m)}")

test_codes = ["06-LCC11", "06-LE11", "RP-06-LCC11", "RP 06-LCC11"]
for c in test_codes:
    norm = _norm_codigo(c)
    hit = m.get(norm)
    print(f"Buscando '{c}' (norm: '{norm}'): {'ENCONTRADO' if hit else 'NO ENCONTRADO'}")
    if hit:
        print(f"  -> {hit}")
