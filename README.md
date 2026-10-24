# ATTSF · Aplicación Ecodriving

## 1. Requisitos
- Python 3.10 o posterior
- El archivo Excel mensual de la flota

## 2. Instalación local
En una terminal, dentro de esta carpeta:

```bash
python -m venv .venv
```

Activar el entorno:
- Windows: `.venv\Scripts\activate`
- macOS/Linux: `source .venv/bin/activate`

Instalar dependencias y ejecutar:

```bash
pip install -r requirements.txt
streamlit run app.py
```

La aplicación se abrirá en el navegador. La base de datos `ecodriving.db` se crea automáticamente en la carpeta desde la que ejecutes Streamlit.

## 3. Contraseña de administrador
Por defecto, la contraseña de demostración es `cambiar-esta-clave`. Cámbiala antes de compartir o publicar la aplicación.

Windows PowerShell:
```powershell
$env:ATTSF_ADMIN_PASSWORD="pon-una-clave-larga-y-unica"
streamlit run app.py
```

macOS/Linux:
```bash
export ATTSF_ADMIN_PASSWORD="pon-una-clave-larga-y-unica"
streamlit run app.py
```

También se puede crear `.streamlit/secrets.toml`:
```toml
admin_password = "pon-una-clave-larga-y-unica"
```
No subas ese archivo a un repositorio público.

## 4. Formato Excel esperado
Una fila por vehículo y mes. Encabezados aceptados (se reconocen variantes comunes):
- Número de vehículo
- Distancia recorrida (km)
- Velocidad media (km/h)
- km carretera interior
- km carretera exterior
- duración viaje
- Ralentí
- Número de paradas
- Opcional: duración planificada

Las duraciones pueden ser horas numéricas o formato `HH:MM:SS` / `HH:MM`. Si la fuente proporciona el ralentí como porcentaje y no como duración, adapta `prepare_excel()` antes de usarla.

## 5. Despliegue gratuito en Streamlit Community Cloud
1. Crea un repositorio GitHub y sube `app.py`, `requirements.txt` y `README.md`.
2. Entra en https://share.streamlit.io/ y conecta tu cuenta GitHub.
3. Selecciona el repositorio, rama y archivo `app.py`.
4. En **Advanced settings / Secrets**, añade `admin_password = "una-clave-larga-y-unica"`.
5. Despliega la aplicación.

**Persistencia importante:** la base SQLite local puede no sobrevivir a reinicios/redeploys en Community Cloud y no es una base de datos compartida multiusuario fiable. Para datos reales de ATTSF, configura una base de datos persistente externa (por ejemplo PostgreSQL administrado) y adapta la capa de almacenamiento antes de producción. No publiques datos personales de conductores sin autorización y medidas de protección.

## 6. Reglas y limitaciones de la versión inicial
- El sistema calcula una puntuación por vehículo y mes (una fila por vehículo/mes), con un máximo teórico de 100 puntos.
- Los criterios se configuran en el panel de parámetros; cambios en parámetros afectan a nuevas importaciones. Reimporta un mes para recalcularlo.
- Los puntos nunca son negativos. No se restan puntos por no alcanzar una insignia.
- La estimación de combustible/CO₂ es hipotética porque el Excel no contiene consumo medido. El factor predeterminado es 2,68 kg CO₂/litro, configurable; valida el valor y la metodología para informes oficiales.
- La duración planificada es opcional. Sin ella, se concede la puntuación configurada para duración, para no penalizar la falta de datos.
- La puntuación por distancia es proporcional hasta el máximo configurado. La velocidad eficiente predeterminada es 65–85 km/h y debe validarse según el tipo de camión, carga, carretera y normativa.
- Esta aplicación es una base funcional de partida, no una auditoría de seguridad. Antes de usarla como sistema oficial, valida los cálculos y la seguridad de acceso.
