# Control de Nóminas

Una aplicación de escritorio desarrollada en Python para gestionar y controlar datos de nóminas de empleados de forma sencilla e intuitiva.

## 📋 Tecnologías Utilizadas

- **Python 3.x** - Lenguaje de programación principal
- **Tkinter** - Biblioteca para crear la interfaz gráfica de usuario (GUI)
- **OpenPyXL** - Librería para leer y escribir archivos Excel (.xlsx)
- **Datetime & Calendar** - Módulos de Python para manejo de fechas y calendarios

## 🚀 Cómo Ejecutar

### Requisitos Previos

Asegúrate de tener Python 3 instalado en tu sistema. Luego, instala las dependencias necesarias:

```bash
pip install openpyxl
```

> **Nota:** Tkinter generalmente viene incluido con Python, pero si no está disponible, debes instalarlo según tu sistema operativo.

### Ejecución

Desde la terminal, navega a la carpeta del proyecto y ejecuta:

```bash
python control_nomina.py
```

También puedes hacer doble clic en el archivo `control_nomina.py` si tu sistema está configurado para ejecutar scripts de Python.

## 📝 Descripción del Funcionamiento

### ¿Qué hace esta aplicación?

El sistema **Control de Nóminas** es una herramienta para registrar y organizar información sobre empleados y sus datos salariales. La aplicación proporciona:

#### Características Principales:

1. **Interfaz Gráfica Intuitiva**
   - Formulario con campos para ingresar datos de empleados
   - Tema oscuro personalizado para una mejor experiencia visual
   - Navegación sencilla y accesible

2. **Selector de Fechas Integrado**
   - Calendarios emergentes para seleccionar fechas
   - Evita errores manuales en la entrada de datos
   - Navegación entre meses para mayor flexibilidad

3. **Gestión de Datos de Empleados**
   - Código de empleado
   - Nombre y apellidos
   - Fecha de ingreso
   - Cargo
   - Horas trabajadas y no trabajadas en el año
   - Precio por hora
   - Retenciones (IRPF)
   - Cargas familiares (10 categorías diferentes)

4. **Exportación a Excel**
   - Almacenamiento de datos en archivo Excel (`nominas.xlsx`)
   - Estructura ordenada con encabezados claros
   - Fácil acceso para análisis posterior

5. **Opciones de Cargas Familiares**
   - Sin cargas familiares
   - Hijos/as (1, 2, 3 o más)
   - Cónyuge a cargo
   - Familia numerosa
   - Familia monoparental
   - Ascendientes a cargo
   - Dependientes con discapacidad

### Flujo de Trabajo

1. El usuario abre la aplicación
2. Completa el formulario con los datos del empleado
3. Utiliza los calendarios emergentes para seleccionar fechas
4. Selecciona la categoría de cargas familiares
5. Guarda los datos, que se registran automáticamente en el archivo Excel
6. Puede continuar agregando más empleados

## 📁 Archivos del Proyecto

- `control_nomina.py` - Archivo principal con toda la lógica de la aplicación
- `nominas.xlsx` - Archivo Excel donde se guardan los registros (se crea automáticamente)
- `README.md` - Este archivo con la documentación

## 💡 Notas Técnicas

- La aplicación busca automáticamente los archivos en la misma carpeta donde se encuentra el script
- El archivo Excel se crea automáticamente en la primera ejecución
- La interfaz incluye un icono personalizado (nota: compatibilidad varía según el sistema operativo)
- Los datos se organizan de forma clara y estructurada para facilitar reportes posteriores

## ⚙️ Requisitos del Sistema

- Python 3.6 o superior
- 50 MB de espacio en disco disponible
- Cualquier sistema operativo con soporte para Python (Windows, macOS, Linux)

---

**Creado por:** Naiara Mega Garay  
**Última actualización:** 2026
<img width="656" height="706" alt="Control_nominas" src="https://github.com/user-attachments/assets/9572c00f-11db-4ade-811c-f2bfbe661954" />
