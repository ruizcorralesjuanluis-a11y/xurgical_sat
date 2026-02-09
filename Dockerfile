
# Usar una imagen base ligera de Python
FROM python:3.11-slim

# Establecer el directorio de trabajo
WORKDIR /app

# Copiar los archivos de requisitos e instalar dependencias
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el resto del código
COPY . .

# Configurar la variable de entorno para la BD temporal (necesario en Render sin disco)
ENV XURGICAL_DB_PATH=/tmp/sat.db

# Exponer el puerto (Render lo asigna dinámicamente, pero esto es buena práctica)
EXPOSE 8000

# Comando para iniciar la aplicación
# Usamos sh -c para que pueda expandir la variable $PORT que Render inyecta
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
