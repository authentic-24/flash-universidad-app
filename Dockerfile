# Usar imagen base slim para optimizar tamaño
FROM python:3.9-slim

# Establecer el directorio de trabajo dentro del contenedor
WORKDIR /app

# Copiar los archivos de dependencias antes de copiar el código
# para aprovechar el caché de Docker en instalaciones repetidas
COPY requirements.txt .

# Instalar dependencias
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el resto de la aplicación
COPY . .

# Exponer el puerto que utilizará la aplicación
EXPOSE 8080

# Ejecutar la aplicación
CMD ["python", "app.py"]