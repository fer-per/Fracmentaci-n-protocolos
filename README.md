# Sistema de Automatización Archivística

Herramienta profesional para la fragmentación automatizada de protocolos y documentos notariales escaneados a partir de un inventario en Excel (`.xlsx`). 

Diseñada bajo los principios de **Arquitectura Limpia (Clean Architecture)** y estructurada de forma modular, esta aplicación desacopla las reglas de negocio de la infraestructura y de la interfaz gráfica.

---

## Características Principales

*   **Interfaz de Usuario Moderna**: Estilo Notion, limpia, intuitiva y responsiva.
*   **Procesamiento Asíncrono**: Fragmentación en hilos secundarios para mantener la GUI fluida y permitir la cancelación en tiempo real.
*   **Analizadores Incorporados**:
    *   **Sucesión de Folios**: Valida que la numeración de los folios en el Excel sea continua.
    *   **Data Tópica**: Verifica el formato geográfico.
    *   **Data Crónica**: Valida fechas y asegura un orden cronológico progresivo.
    *   **Cobertura del PDF**: Calcula y compara las páginas que usará el rango contra la extensión real del PDF cargado.
*   **Configuración Flexible**:
    *   Definición de folios de inicio del protocolo y páginas de inicio reales en el PDF.
    *   Manejo de saltos en el PDF (mediante la definición de segmentos adicionales).
    *   Omisión y exclusión de páginas específicas o rangos de páginas a ignorar.
*   **Reportes Detallados**: Generación automática de reportes de mapeo de folios a páginas PDF en formato de texto plano (`.txt`).

---

## Instalación

1.  Asegúrate de tener instalado Python 3.10 o superior.
2.  Instala las dependencias necesarias:

```bash
pip install -r requirements.txt
```

---

## Configuración (`configuracion.py`)

Toda la parametrización central del sistema está unificada en `configuracion.py`. Antes de ejecutar el proyecto, puedes revisar o ajustar allí constantes como:

*   Nombres exactos de las columnas del Excel (ej. `N° DE FOLIOS`, `INTERESADO 1`).
*   Líneas a omitir al leer el inventario.
*   Directorios por defecto para las salidas y bitácoras.

---

## Uso

### Ejecutar la aplicación gráfica

```bash
python gui.py
```

1.  **Paso 1**: Selecciona tu archivo de inventario en Excel (`.xlsx`) y el archivo PDF escaneado del protocolo.
2.  **Paso 2**: Revisa la vista previa de datos. Define el rango de filas a procesar (por defecto procesa todo).
3.  **Paso 2b**: Configura el folio de inicio, página inicial del PDF, y añade segmentos o páginas a ignorar si existen desfases o portadas.
4.  **Paso 2c**: Utiliza los botones de análisis para validar la consistencia de folios, fechas y cobertura antes de procesar.
5.  **Paso 3 & 4**: Elige la carpeta de destino y presiona **PROCESAR**.

---

## Estructura del Proyecto (Clean Architecture)

El proyecto está organizado en capas independientes para facilitar su mantenimiento y escalabilidad:

```text
automa_fragmentacionPDF/
├── gui.py                      # Lanzador principal de la aplicación
├── configuracion.py            # Parámetros globales y mapeo de columnas
├── requirements.txt            # Dependencias del proyecto (pandas, openpyxl, pypdf, reportlab)
├── README.md                   # Documentación del proyecto
└── modules/
    ├── contenedor.py           # Contenedor de Inyección de Dependencias (Wiring)
    ├── lector_excel.py         # Utilidad de bajo nivel para lectura de DataFrames
    │
    ├── dominio/                # CAPA DOMINIO: Modelos y servicios lógicos independientes
    │   ├── modelos.py          # Definiciones de Protocolo, Registro, Segmento, etc.
    │   └── servicios/          # Reglas de negocio puras
    │       ├── servicio_analisis.py
    │       ├── servicio_folios.py
    │       └── servicio_validacion.py
    │
    ├── aplicacion/             # CAPA APLICACIÓN: Casos de uso y puertos
    │   ├── puertos.py          # Interfaces de repositorios y servicios
    │   └── casos_uso.py        # Casos de uso del sistema (Fragmentar, Analizar, Cobertura)
    │
    ├── infraestructura/        # CAPA INFRAESTRUCTURA: Adaptadores de persistencia y PDF
    │   └── adaptadores.py      # Lector de Excel y escritor/fragmentador de PDF físicos
    │
    └── gui/                    # CAPA PRESENTACIÓN: Componentes y lógica de la GUI
        ├── aplicacion.py       # Clase principal y bucle de la aplicación
        ├── constructor_ui.py   # Construcción de widgets y vistas
        ├── estilos.py          # Paleta de colores, tipografías y estilos (Notion theme)
        ├── analizadores.py     # Controladores que invocan los casos de uso de análisis
        ├── procesador.py       # Controlador asíncrono del flujo de fragmentación
        └── manejadores_eventos.py # Eventos e interacciones de usuario (archivos, ignorados, segmentos)
```

---

## Notación de Folios Admitida

El formateador y lector de folios de la capa de dominio entiende el sistema de folios recto/vuelta (`r` / `v`):

| Formato | Ejemplo | Páginas PDF Resultantes |
| :--- | :--- | :--- |
| **Rango completo** | `1r-1v` | Página 1 y Página 2 |
| **Multi-hoja** | `4r-6v` | Páginas 7, 8, 9, 10, 11 y 12 |
| **Solo cara recto** | `7r` | Página 13 |
| **Cruzado** | `7v-8r` | Página 14 y Página 15 |
| **Rango vuelta** | `8v-12r` | Páginas 16 a 23 |
