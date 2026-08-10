# �� 📈 �� 🤖 Trading Bot para Quotex con Señales de Telegram

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Playwright](https://img.shields.io/badge/Playwright-1.40%2B-green)
![Telethon](https://img.shields.io/badge/Telethon-1.29%2B-purple)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.1%2B-orange)

## � ℹ��️ �� 📋 Descripción

Este bot automatiza operaciones de trading en la plataforma **Quotex** utilizando señales recibidas desde un grupo de **Telegram**. El bot se conecta a Telegram, escucha mensajes específicos que contienen señales de trading, y luego ejecuta automáticamente las operaciones en el navegador mediante **Playwright**. Los resultados de cada operación se registran en archivos **Excel** para análisis y seguimiento.

## �� ⚙��️ �� 🛠��️ Tecnologías Utilizadas

- **Python 3.10+**: Lenguaje principal del proyecto.
- **Playwright**: Para automatización del navegador (Firefox) y interacción con la plataforma Quotex.
- **Telethon**: Para la conexión y escucha de mensajes en Telegram.
- **OpenPyXL**: Para lectura y escritura de archivos Excel donde se almacenan los reportes de operaciones.
- **Colorama**: Para colorear la salida en la consola y hacerla más legible.
- **Asyncio**: Para manejar operaciones asíncronas de manera eficiente.

## �� 💡 �� 🧠 ¿Por qué estas tecnologías?

- **Playwright**: Es moderno, confiable y soporta múltiples navegadores. Es ideal para automatizar interacciones complejas en páginas web como Quotex.
- **Telethon**: Biblioteca robusta y bien mantenida para interactuar con la API de Telegram.
- **OpenPyXL**: Permite trabajar con archivos .xlsx sin necesidad de tener Microsoft Excel instalado.
- **Colorama**: Mejora la experiencia de desarrollo y depuración con colores en la terminal.
- **Asyncio**: Esencial para manejar múltiples tareas concurrentes (navegador, Telegram, lógica de trading) sin bloqueos.

## �� 🔧 �� ⚙��️ Cómo Funciona

1. **Conexión a Telegram**: El bot se conecta a tu cuenta de Telegram y escucha mensajes en un grupo específico (o en tus mensajes guardados) que contengan la señal de trading.
2. **Procesamiento de Señal**: Cuando detecta un mensaje con el patrón esperado (que incluye zona horaria, duración, par de moneda y dirección), extrae los datos necesarios.
3. **Automatización del Navegador**: 
    - Lanza una instancia de Firefox en modo no headless (para que puedas ver lo que sucede).
    - Navega a Quotex (versión demo o real, según configuración).
    - Inicia sesión automáticamente (maneja incluso la autenticación de dos factores si es requerida).
    - Monta la operación según los parámetros de la señal (duración, par, dirección, monto).
    - Ejecuta la operación en el momento especificado.
4. **Registro de Resultados**: Después de cada operación, el bot registra los resultados en un archivo Excel ubicado en la carpeta `excel/`:
    - Fecha y hora de la operación
    - Duración
    - Par de moneda
    - Dirección (PUT/CALL)
    - Tipo de resultado (D: Directo, G1: Gale 1, G2: Gale 2, P: Pérdida)
    - Balance inicial y final
    - Ganancia/Pérdida
5. **Repetición**: El bot permanece activo, esperando nuevas señales para repetir el proceso.

## �� 🗂��️ �� 📂 Estructura del Proyecto

```
trading_bot/
│
├── main.py                 # Script principal que inicia el bot
├── variables.py            # Configuración y variables sensibles (credenciales, parámetros)
├── README.md               # Este archivo
│
├── excel/                  # �� 💰 �� 📊 Aquí se guardan los reportes de operaciones en formato .xlsx
│   ├── operaciones.xlsx
│   ├── operaciones_5_5_5.xlsx
│   ├── operaciones_6_7_7.xlsx
│   ├── operaciones_10_5_10.xlsx
│   ├── operaciones_10_10_10.xlsx
│   ├── operaciones_anterior.xlsx
│   └── consolidado de operaciones.xlsx
│
├── data/                   # �� 🔒 �� 💾 Datos persistentes (sesiones de Telegram, logs)
│   ├── session_name.session # Archivo de sesión de Telethon
│   └── registro.log         # Registro detallado de eventos (nivel DEBUG)
│
├── logs/                   # �� 📓 �� 📝 Logs de ejecución (separado por claridad)
│   └── registro.log         # Copia o registro principal de logs (configurado en main.py)
│
├── archives/               # �� 📁 �� 📦 Archivos comprimidos o históricos (ej. quotexapi-main.zip)
│   └── quotexapi-main.zip
│
├── config/                 # �� 🔧 �� ⚙��️ Configuraciones adicionales (si las hubiera)
├── documentation/          # �� 📖 �� 📚 Documentación adicional
├── images/                 # �� 🖼��️ Imágenes utilizadas en el proyecto (si las hubiera)
�└── __pycache__/            # �� 🗑��️ Cache de Python (se genera automáticamente, se puede ignorar)
```

## �� 🛠��️ �� 🚀 Implementación y Setup

### Prerrequisitos

1. **Python 3.10 o superior** instalado en tu máquina.
2. **Cuenta de Telegram** y acceso al grupo o chat donde se envían las señales.
3. **Cuenta en Quotex** (puedes empezar con la cuenta demo).
4. **Credenciales de acceso** a Quotex (email y contraseña).

### Pasos para Instalar

1. **Clona o descarga este repositorio** en tu máquina local.
2. **Instala las dependencias** ejecutando en la terminal:
   ```bash
   pip install -r requirements.txt
   ```
   > Nota: Si no tienes un `requirements.txt`, puedes instalar manualmente:
   > ```bash
   > pip install playwright telethon openpyxl colorama
   > playwright install firefox
   > ```

3. **Configura las variables** en `variables.py`:
   - `telegram_api_id` y `telegram_api_hash`: Obtén estos valores de [my.telegram.org](https://my.telegram.org).
   - `telegram_group_username`: El nombre de usuario del grupo de Telegram (ej. `@magictradersignals`).
   - `quotex_email` y `quotex_password`: Tus credenciales de Quotex.
   - Ajusta otros parámetros según tu estrategia (montos, tiempos, etc.).

4. **Ejecuta el bot**:
   ```bash
   python main.py
   ```
   - La primera vez que se ejecute, te pedirá el código de autenticación de Telegram (envía el código que recibas en tu cuenta de Telegram guardado).
   - Para Quotex, si se activa la autenticación de dos factores, deberás ingresar el código manualmente cuando se solicite en la consola.

## �� 📊 �� 📈 Reportes en Excel

El bot genera reportes detallados en la carpeta `excel/`. Cada fila representa una operación y contiene:

| Fecha Operación | Hora Operación | Duración | Par | Dirección | Tipo de Resultado | Balance Inicial | Balance Final | Profit |
|-----------------|----------------|----------|-----|-----------|-------------------|-----------------|---------------|--------|
| 10 de agosto de 2026 | 14:30:00 | 1 | USD/TRY | PUT | D | 100.00 | 102.50 | 2.50 |
| 10 de agosto de 2026 | 14:35:00 | 5 | EUR/USD | CALL | G1 | 102.50 | 100.00 | -2.50 |

- **Tipos de Resultado**:
  - **D**: Operación directa (ganó en el primer intento).
  - **G1**: Ganó en el primer intento de Gale.
  - **G2**: Ganó en el segundo intento de Gale.
  - **P**: Perdió después de todos los intentos (incluyendo Gales).

Estos archivos te permiten analizar el rendimiento del bot, ajustar parámetros y llevar un control detallado de tu actividad de trading.

## �� 👨‍���💻 Autor

**W4k4ndA**  
Desarrollador y entusiasta de la robotica.  
Este proyecto es un primer acercamiento a una herramienta para optimizar la ejecución de señales de trading recibidas mediante Telegram.

## �� 📄 �� 📜 Licencia

Este proyecto es de uso privado y educativo. No se redistribuye ni se utiliza para fines comerciales sin autorización explícita del autor.

## �� ⚠��️ Advertencia

El trading implica riesgos. Este bot es una herramienta de automatización y no garantiza ganancias. Úsalo bajo tu propia responsabilidad y siempre empieza con una cuenta demo para probar su funcionamiento.

---

���💡 **Tip**: Mantén tus credenciales seguras y nunca compartas tu `variables.py` o archivos de sesión. Añade estas rutas a tu `.gitignore` si planeas usar control de versiones.

¡Éxitos en tus operaciones! �� 🚀 �� 📈