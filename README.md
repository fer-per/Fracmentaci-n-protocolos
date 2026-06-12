# Sistema de Automatización Archivística

Automatiza la fracmentación de documentos notariales a partir de un archivo Excel (`.xlsx`).

---

## Instalación

```bash
pip install -r requirements.txt
```

---

## Configuración (`config.py`)

Antes de ejecutar, edita `config.py` y ajusta las rutas:

```python
EXCEL_PATH = BASE_DIR / "inventario.xlsx"   # ← tu archivo Excel real
PDF_PATH   = BASE_DIR / "documento.pdf"     # ← tu PDF escaneado real
```

---

## Uso Rápido

### Ejecutar la aplicación gráfica

```bash
python gui.py
```

Esto iniciará la interfaz interactiva donde podrás:
1. Seleccionar el inventario Excel y el PDF original.
2. Visualizar y configurar el rango de filas.
3. Configurar folios de inicio, páginas de inicio, saltos en PDF y exclusión de páginas a ignorar.
4. Ejecutar validaciones (sucesión de folios, tópica, crónica y cobertura).
5. Iniciar la fragmentación y guardar los documentos ordenadamente.

---

## Estructura de Salida

```
output/
└── PORTUGAL, Cesar/
    └── 1567 - Protocolo N° 1/
        └── Obligacion/
            └── Julio/
                └── Diego de Aramburu/
                    └── Antonio de Oviedo.pdf
```

---

## Logs

| Archivo | Contenido |
|---|---|
| `logs/process.log` | Log completo de cada operación |
| `logs/pendientes.csv` | Registros no procesados con motivo |

---

## Notación de Folios

| Formato | Ejemplo | Páginas PDF |
|---|---|---|
| Rango completo | `1r-1v` | 1, 2 |
| Multi-hoja | `4r-6v` | 7, 8, 9, 10, 11, 12 |
| Solo recto | `7r` | 13 |
| Cruzado | `7v-8r` | 14, 15 |
| Rango vuelta | `8v-12r` | 16–23 |

---

## Estructura del Proyecto

```
fracmen_auto/
├── gui.py                   # Lanzador principal de la interfaz gráfica
├── config.py                # Configuración centralizada
├── requirements.txt
├── modules/
│   ├── excel_reader.py      # Lectura del Excel .xlsx
│   ├── folio_parser.py      # Algoritmo r/v → páginas PDF
│   ├── pdf_extractor.py     # Extracción y escritura de PDFs
│   ├── folder_builder.py    # Construcción de estructura de carpetas
│   ├── validator.py         # Validación de registros
│   └── gui/                 # Interfaz de usuario (estilos, controladores, hilos, componentes)
├── output/                  # PDFs generados (se crea automáticamente)
└── logs/                    # Logs y pendientes (se crea automáticamente)
```
