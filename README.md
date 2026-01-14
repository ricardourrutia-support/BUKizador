# BUKizador# 🤖 BUKizador

Una herramienta minimalista para transformar, limpiar e inyectar turnos de formato "Supervisor" directamente a plantillas de carga masiva de **BUK**.

## ✨ Características

* **Algoritmo de Limpieza Vectorizada:** Procesa miles de turnos en milisegundos.
* **Búsqueda Difusa (Fuzzy Matching):** Detecta colaboradores aunque el supervisor escriba mal el nombre (ej: "Anahis" vs "Anais").
* **Inyección de Plantilla:** Respeta al 100% los metadatos y encabezados de tu archivo original de BUK.
* **Interfaz Minimalista:** Sin distracciones, solo Input -> Proceso -> Output.

## 🚀 Instalación y Uso

1.  **Clonar repositorio:**
    ```bash
    git clone <tu-repo-url>
    cd bukizador
    ```

2.  **Instalar dependencias:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Ejecutar la aplicación:**
    ```bash
    streamlit run app.py
    ```

## 📂 Archivos Requeridos

1.  **Input de Turnos (Excel):** Debe contener 3 hojas:
    * `Turnos Formato Supervisor`: Matriz visual de turnos.
    * `Base de Colaboradores`: Maestro con RUT, Nombre, Área, Supervisor.
    * `Codificación de Turnos`: Diccionario de horarios a siglas.
2.  **Plantilla BUK (XLS/CSV):** El archivo vacío descargado desde BUK donde quieres inyectar los datos.

---
*Hecho para simplificar la vida de RRHH.*
