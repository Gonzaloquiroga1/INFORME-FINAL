# Generador de Informe General de Turno

## Cómo ejecutar
1. Coloca estos archivos en la misma carpeta:
   - `app_generador_informe_turno.py`
   - `F-GE-1483_template.xlsx`
   - `requirements_generador_informe_turno.txt`

2. Instala dependencias:
   ```bash
   pip install -r requirements_generador_informe_turno.txt
   ```

3. Ejecuta la app:
   ```bash
   streamlit run app_generador_informe_turno.py
   ```

## Qué hace
- Carga el formato base del informe.
- Completa el libro con los datos diligenciados.
- Exporta un archivo Excel listo para descargar.

