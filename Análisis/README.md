# Análisis

Carpeta destinada a generar y almacenar los archivos de resultados del proyecto **Legal PP Project LC**.

Aquí se guardan los productos del análisis (resúmenes, revisiones, informes, hallazgos, etc.) generados a partir de los documentos de la carpeta `Contract`.

## Formatos de salida

Cada análisis se genera automáticamente en **tres formatos**:

- 📄 **PDF** (`.pdf`) — informe legible y listo para compartir
- 📝 **Word** (`.docx`) — documento editable
- 📊 **Excel** (`.xlsx`) — datos tabulados por secciones (una hoja por sección)

## Generador automático

El generador vive en la subcarpeta [`generador/`](generador/):

| Archivo | Descripción |
|---|---|
| `generar_analisis.py` | Script que produce PDF, Word y Excel a partir de un JSON |
| `ejemplo_contrato.json` | Ejemplo del esquema de datos esperado |
| `requirements.txt` | Dependencias de Python |
| `ejemplos_salida/` | Documentos de muestra ya generados (PDF, Word, Excel) |

### Cómo usarlo

```bash
# 1. (una sola vez) instalar dependencias
pip install -r Análisis/generador/requirements.txt

# 2. generar los tres formatos a partir de un análisis en JSON
python Análisis/generador/generar_analisis.py mi_contrato.json

# Opciones:
#   - Carpeta de salida personalizada:
python Análisis/generador/generar_analisis.py mi_contrato.json carpeta_salida
#   - Solo algunos formatos:
python Análisis/generador/generar_analisis.py mi_contrato.json . --formatos pdf,docx
```

Si no se indica carpeta de salida, los archivos se crean en esta carpeta `Análisis/`.

## Flujo de trabajo

1. Subes un contrato a la carpeta `Contract/`.
2. Se analiza el documento y se vuelca el resultado en un JSON con el esquema de `ejemplo_contrato.json`.
3. El generador produce automáticamente el informe en **PDF, Word y Excel**.
