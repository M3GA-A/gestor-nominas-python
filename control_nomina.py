import calendar
import os
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk
from tkinter import PhotoImage #para que aparezca el icono personalizado en la ventana, aunque no es compatible con todos los sistemas operativos.
import openpyxl


# Ruta donde esta guardado este archivo .py.
# Se usa para que el Excel y el icono se busquen siempre en la misma carpeta.
RUTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(RUTA_SCRIPT, "nominas.xlsx")

# Opciones que apareceran en el desplegable de cargas familiares.
CARGAS_FAMILIARES = [
    "Sin cargas familiares",
    "1 hijo/a",
    "2 hijos/as",
    "3 o más hijos/as",
    "Cónyuge a cargo",
    "Familia numerosa",
    "Familia monoparental",
    "Ascendiente a cargo",
    "Hijo/a con discapacidad",
    "Dependiente con discapacidad",
]

# Nombres de los meses para mostrar el calendario.
MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

# Encabezados que se escriben en la primera fila del archivo Excel.
CAMPOS_EXCEL = [
    "Cod empleado",
    "Nombre",
    "Apellidos",
    "Fecha de ingreso",
    "Cargo",
    "Horas trabajadas año",
    "Horas no trabajadas año",
    "Precio hora",
    "IRPF",
    "Cargas familiares",
]

# Colores principales del formulario.
COLOR_FONDO = "#1f1f24"
COLOR_PANEL = "#2b2b31"
COLOR_ENTRADA = "#32323a"
COLOR_TEXTO = "#f5f5f7"
COLOR_TEXTO_SECUNDARIO = "#c7c7cc"
COLOR_ACENTO = "#5f7f73"
COLOR_ACENTO_HOVER = "#6f9284"
COLOR_BOTON = "#3a3a42"

#clase para mostrar un calendario y seleccionar una fecha, evitando que el usuario tenga que escribirla manualmente y pueda cometer errores.
class CalendarioPopup:
    """Ventana secundaria para seleccionar una fecha desde un calendario."""

    def __init__(self, ventana_padre, entrada_fecha):
        # Guardamos el Entry donde se escribira la fecha seleccionada.
        self.entrada_fecha = entrada_fecha
        self.fecha_actual = datetime.today()

        # Toplevel crea una ventana nueva encima del formulario principal.
        self.ventana = tk.Toplevel(ventana_padre)
        self.ventana.title("Seleccionar fecha")
        self.ventana.configure(bg=COLOR_FONDO)
        self.ventana.resizable(False, False)
        self.ventana.transient(ventana_padre)
        self.ventana.grab_set()

        self.contenedor = ttk.Frame(self.ventana, padding=12)
        self.contenedor.grid(row=0, column=0)
        self.dibujar_calendario()

    def dibujar_calendario(self):
        # Cada vez que cambiamos de mes se borra el calendario anterior.
        for widget in self.contenedor.winfo_children():
            widget.destroy()

        encabezado = ttk.Frame(self.contenedor)
        encabezado.grid(row=0, column=0, columnspan=7, pady=(0, 8))

        # Boton para ir al mes anterior.
        ttk.Button(encabezado, text="<", width=3, command=self.mes_anterior).grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Label(
            encabezado,
            text=f"{MESES[self.fecha_actual.month - 1]} {self.fecha_actual.year}",
            width=20,
            anchor="center",
        ).grid(row=0, column=1)
        # Boton para ir al mes siguiente.
        ttk.Button(encabezado, text=">", width=3, command=self.mes_siguiente).grid(
            row=0,
            column=2,
            padx=(8, 0),
        )

        dias_semana = ["L", "M", "X", "J", "V", "S", "D"]
        for columna, dia in enumerate(dias_semana):
            ttk.Label(self.contenedor, text=dia, anchor="center", width=4).grid(
                row=1,
                column=columna,
                pady=(0, 4),
            )

        calendario_mes = calendar.monthcalendar(self.fecha_actual.year, self.fecha_actual.month)
        for fila, semana in enumerate(calendario_mes, start=2):
            for columna, dia in enumerate(semana):
                if dia == 0:
                    # Los ceros son huecos antes o despues del mes.
                    ttk.Label(self.contenedor, text="", width=4).grid(row=fila, column=columna)
                else:
                    ttk.Button(
                        self.contenedor,
                        text=str(dia),
                        width=4,
                        command=lambda d=dia: self.seleccionar_fecha(d),
                    ).grid(row=fila, column=columna, padx=1, pady=1)

    def mes_anterior(self): 
        mes = self.fecha_actual.month - 1
        anio = self.fecha_actual.year
        if mes == 0:
            mes = 12
            anio -= 1
        self.fecha_actual = self.fecha_actual.replace(year=anio, month=mes, day=1)
        self.dibujar_calendario()

    def mes_siguiente(self):
        mes = self.fecha_actual.month + 1
        anio = self.fecha_actual.year
        if mes == 13:
            mes = 1
            anio += 1
        self.fecha_actual = self.fecha_actual.replace(year=anio, month=mes, day=1)
        self.dibujar_calendario()

    def seleccionar_fecha(self, dia):
        # Al pulsar un dia, se escribe la fecha en formato DD/MM/AAAA.
        fecha = self.fecha_actual.replace(day=dia)
        self.entrada_fecha.delete(0, tk.END)
        self.entrada_fecha.insert(0, fecha.strftime("%d/%m/%Y"))
        self.ventana.destroy()

# Crea el archivo Excel con sus encabezados si todavia no existe.
def inicializar_excel():
    """Crea el archivo Excel con sus encabezados si todavia no existe."""
    if not os.path.exists(ARCHIVO_EXCEL):
        libro = openpyxl.Workbook()
        hoja = libro.active
        hoja.title = "Nominas"
        hoja.append(CAMPOS_EXCEL)
        libro.save(ARCHIVO_EXCEL)

# Valida que la fecha tenga el formato correcto y sea una fecha real(ej: 21/02/2030 no es valida). No se permiten fechas futuras.
def validar_fecha(fecha):
    """Comprueba que la fecha tenga formato DD/MM/AAAA."""
    try:
        datetime.strptime(fecha, "%d/%m/%Y")
        return True
    except ValueError:
        return False

# Valida y convierte un texto numerico a float, permitiendo punto o coma decimal y simbolo de euro.
def validar_numero(valor):
    """Devuelve el float valido o None si no es un numero valido."""
    try:
        return float(valor.replace("€", "").replace(",", ".").strip())
    except (ValueError, AttributeError):
        return None


# Valida que un campo de texto solo tenga letras y espacios, para evitar caracteres invalidos en el Excel.
def validar_texto(valor):
    """Valida campos que deben tener solo letras y espacios."""
    return all(letra.isalpha() or letra.isspace() for letra in valor)

# Lee los campos del formulario, valida que sean correctos y devuelve una lista de valores ordenada para escribir en Excel.
def obtener_valores():
    """Lee los campos del formulario y aplica todas las validaciones."""
    valores = {
        "Cod empleado": entry_cod_empleado.get().strip(),
        "Nombre": entry_nombre.get().strip(),
        "Apellidos": entry_apellidos.get().strip(),
        "Fecha de ingreso": entry_fecha_ingreso.get().strip(),
        "Cargo": entry_cargo.get().strip(),
        "Horas trabajadas año": entry_horas_trabajadas.get().strip(),
        "Horas no trabajadas año": entry_horas_no_trabajadas.get().strip(),
        "Precio hora": entry_precio_hora.get().strip(),
        "IRPF": entry_irpf.get().strip(),
        "Cargas familiares": combo_cargas.get().strip(),
    }

    if not all(valores.values()):
        messagebox.showerror("Error", "Todos los campos son obligatorios.")
        return None

    # El codigo de empleado puede ser un numero interno, DNI o NIE.
    if not valores["Cod empleado"].isalnum():
        messagebox.showerror("Error", "Cod empleado solo puede tener letras y numeros.")
        return None

    if not validar_texto(valores["Nombre"]):
        messagebox.showerror("Error", "Nombre solo puede tener letras y espacios.")
        return None

    if not validar_texto(valores["Apellidos"]):
        messagebox.showerror("Error", "Apellidos solo puede tener letras y espacios.")
        return None

    if not validar_texto(valores["Cargo"]):
        messagebox.showerror("Error", "Cargo solo puede tener letras y espacios.")
        return None

    if not validar_fecha(valores["Fecha de ingreso"]):
        messagebox.showerror("Error", "Fecha de ingreso debe tener formato DD/MM/AAAA.")
        return None

    fecha_ingreso = datetime.strptime(valores["Fecha de ingreso"], "%d/%m/%Y")
    if fecha_ingreso > datetime.today():
        messagebox.showerror("Error", "Fecha de ingreso no puede ser una fecha futura.")
        return None

    campos_numericos = [
        "Horas trabajadas año",
        "Horas no trabajadas año",
        "Precio hora",
        "IRPF",
    ]
    for campo in campos_numericos:
        if validar_numero(valores[campo]) is None:
            messagebox.showerror("Error", f"{campo} debe ser un numero valido.")
            return None

    # Convertimos a numero los campos que necesitan rangos concretos usando validar_numero.
    horas_trabajadas = validar_numero(valores["Horas trabajadas año"])
    precio_hora = validar_numero(valores["Precio hora"])
    irpf = validar_numero(valores["IRPF"])

    if horas_trabajadas < 0 or horas_trabajadas > 2500:
        messagebox.showerror("Error", "Horas trabajadas año debe estar entre 0 y 2500.")
        return None

    if precio_hora <= 0:
        messagebox.showerror("Error", "Precio hora debe ser mayor que 0.")
        return None

    if irpf < 0 or irpf > 100:
        messagebox.showerror("Error", "IRPF debe estar entre 0 y 100.")
        return None

    valores["Precio hora"] = str(precio_hora)

    if valores["Cargas familiares"] not in CARGAS_FAMILIARES:
        CARGAS_FAMILIARES.append(valores["Cargas familiares"])
        combo_cargas["values"] = CARGAS_FAMILIARES

    return [valores[campo] for campo in CAMPOS_EXCEL]

# Guarda los datos en una nueva fila de Excel si todas las validaciones son correctas.
def guardar_nomina(event=None):
    """Guarda una nueva fila en Excel si todos los datos son validos."""
    valores = obtener_valores()
    if valores is None:
        return

    try:
        # Se abre el Excel existente y se agrega una fila al final.
        libro = openpyxl.load_workbook(ARCHIVO_EXCEL)
        hoja = libro.active
        hoja.append(valores)
        libro.save(ARCHIVO_EXCEL)
    except PermissionError:
        messagebox.showerror("Error", "Cierra el archivo Excel antes de guardar.")
        return

    messagebox.showinfo("Guardado", "Nomina guardada correctamente.")
    limpiar_formulario()

# Limpia el formulario para una nueva entrada.
def limpiar_formulario():
    """Borra el formulario y devuelve el desplegable a su primera opcion."""
    for entrada in entradas_texto:
        entrada.delete(0, tk.END)
    combo_cargas.set(CARGAS_FAMILIARES[0])
    entry_cod_empleado.focus()

# Permite al usuario agregar nuevas opciones de cargas familiares que no esten en la lista predefinida.
def agregar_carga_familiar():
    """Anade al desplegable una nueva opcion escrita por el usuario."""
    nueva_opcion = combo_cargas.get().strip()

    if not nueva_opcion:
        messagebox.showerror("Error", "Escribe una opcion para poder anadirla.")
        return

    if nueva_opcion not in CARGAS_FAMILIARES:
        CARGAS_FAMILIARES.append(nueva_opcion)
        combo_cargas["values"] = CARGAS_FAMILIARES
        messagebox.showinfo("Añadido", "La opcion se añadio al desplegable.")

    combo_cargas.set(nueva_opcion)

#funcion auxiliar para crear las filas del formulario de forma mas ordenada y evitar repetir codigo.
def crear_fila(ventana, fila, texto):
    """Crea una etiqueta y una caja de texto en una fila del grid."""
    ttk.Label(ventana, text=texto).grid(row=fila, column=0, sticky="w", pady=7, padx=(0, 14))
    entrada = ttk.Entry(ventana, width=34)
    entrada.grid(row=fila, column=1, sticky="w", pady=7)
    return entrada


root = tk.Tk()
root.tk.call("tk", "scaling", 1.0) # Ayuda a que el tamano sea mas parecido en macOS Retina.
root.title("Control de Nóminas")
root.geometry("590x620")
root.configure(bg=COLOR_FONDO, padx=28, pady=24)
root.bind("<Return>", guardar_nomina) # Permite guardar pulsando Enter.

# Carga del icono de la ventana.
ruta_icono = os.path.join(RUTA_SCRIPT, "icono.icon")
try:
    if os.path.exists(ruta_icono):
        try:
            root.iconbitmap(ruta_icono)
        except Exception:
            icono = PhotoImage(file=ruta_icono)
            root.iconphoto(True, icono)
    else:
        raise FileNotFoundError(f"No existe el icono: {ruta_icono}")
except Exception as e:
    print(f"Error: No se pudo cargar el icono en {ruta_icono}")
    print(e)


inicializar_excel()

# Configuracion visual: tema oscuro, fuente fija y colores de los widgets.
style = ttk.Style()
style.theme_use("clam") #tema clam que se adapta mejor a los colores personalizados.
root.option_add("*TCombobox*Listbox.font", ("Arial", 13)) #forzamos que el desplegable tenga la misma fuente que el resto del formulario.
style.configure("TFrame", background=COLOR_FONDO)
style.configure(
    "TLabel",
    background=COLOR_FONDO,
    foreground=COLOR_TEXTO_SECUNDARIO,
    font=("Arial", 13),)
style.configure(
    "Titulo.TLabel",
    background=COLOR_FONDO,
    foreground=COLOR_TEXTO,
    font=("Arial", 18, "bold"),
)
style.configure(
    "TEntry",
    fieldbackground=COLOR_ENTRADA,
    background=COLOR_ENTRADA,
    foreground=COLOR_TEXTO,
    insertcolor=COLOR_TEXTO,
    bordercolor=COLOR_PANEL,
    lightcolor=COLOR_PANEL,
    darkcolor=COLOR_PANEL,
    padding=6,
)
style.configure(
    "TCombobox",
    fieldbackground=COLOR_ENTRADA,
    background=COLOR_ENTRADA,
    foreground=COLOR_TEXTO,
    arrowcolor=COLOR_TEXTO,
    bordercolor=COLOR_PANEL,
    lightcolor=COLOR_PANEL,
    darkcolor=COLOR_PANEL,
    padding=4,
    font=("Arial", 13),
)
style.configure(
    "TButton",
    background=COLOR_BOTON,
    foreground=COLOR_TEXTO,
    bordercolor=COLOR_BOTON,
    lightcolor=COLOR_BOTON,
    darkcolor=COLOR_BOTON,
    padding=(14, 8),
    font=("Arial", 13, "bold"),
)
style.configure(
    "Accent.TButton",
    background=COLOR_ACENTO,
    foreground=COLOR_TEXTO,
    bordercolor=COLOR_ACENTO,
    lightcolor=COLOR_ACENTO,
    darkcolor=COLOR_ACENTO,
    focusthickness=0,
    padding=(16, 8),
    font=("Arial", 13, "bold"),
)
style.map(
    "TButton",
    background=[("active", COLOR_PANEL)],
    foreground=[("active", COLOR_TEXTO)],
)
style.map(
    "Accent.TButton",
    background=[("active", COLOR_ACENTO_HOVER)],
    foreground=[("active", COLOR_TEXTO)],
)
style.map(
    "TCombobox",
    fieldbackground=[("readonly", COLOR_ENTRADA)],
    foreground=[("readonly", COLOR_TEXTO)],
)

root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)

# Titulo centrado del formulario.
ttk.Label(
    root,
    text="Control de Nóminas",
    style="Titulo.TLabel",
    anchor="center",).grid(row=0,column=0,columnspan=2,sticky="ew",pady=(0, 18),)

# Campos del formulario, cada uno con su etiqueta y caja de texto. La funcion crear_fila evita repetir codigo.
entry_cod_empleado = crear_fila(root, 1, "1- Cod empleado:")
entry_nombre = crear_fila(root, 2, "2- Nombre:")
entry_apellidos = crear_fila(root, 3, "3- Apellidos:")

ttk.Label(root, text="4- Fecha de ingreso:").grid(row=4, column=0, sticky="w", pady=7, padx=(0, 14))
frame_fecha = ttk.Frame(root)
frame_fecha.grid(row=4, column=1, sticky="w", pady=7)
entry_fecha_ingreso = ttk.Entry(frame_fecha, width=22)
entry_fecha_ingreso.grid(row=0, column=0, sticky="w")

ttk.Button(
    frame_fecha,
    text="Calendario",
    width=12,
    command=lambda: CalendarioPopup(root, entry_fecha_ingreso),
).grid(row=0, column=1, padx=(8, 0))

entry_cargo = crear_fila(root, 5, "5- Cargo:")
entry_horas_trabajadas = crear_fila(root, 6, "6- Horas trabajadas año:")
entry_horas_no_trabajadas = crear_fila(root, 7, "7- Horas no trabajadas año:")
entry_precio_hora = crear_fila(root, 8, "8- Precio hora:")
entry_irpf = crear_fila(root, 9, "9- IRPF:")

# Combobox para elegir una opcion existente o escribir una nueva.
ttk.Label(root, text="10- Cargas familiares:").grid(row=10, column=0, sticky="w", pady=7, padx=(0, 14))
frame_cargas = ttk.Frame(root)
frame_cargas.grid(row=10, column=1, sticky="w", pady=7)
combo_cargas = ttk.Combobox(
    frame_cargas,
    values=CARGAS_FAMILIARES,
    width=24,
    font=("Arial", 13),
)
combo_cargas.grid(row=0, column=0, sticky="w")
combo_cargas.set(CARGAS_FAMILIARES[0])
ttk.Button(
    frame_cargas,
    text="Añadir",
    width=10,
    command=agregar_carga_familiar,
).grid(row=0, column=1, padx=(8, 0))

# Guardamos las cajas de texto en una lista para poder limpiarlas facilmente despues de guardar o al pulsar el boton Limpiar.
entradas_texto = [
    entry_cod_empleado,
    entry_nombre,
    entry_apellidos,
    entry_fecha_ingreso,
    entry_cargo,
    entry_horas_trabajadas,
    entry_horas_no_trabajadas,
    entry_precio_hora,
    entry_irpf,
]

# Botones principales del formulario.
frame_botones = ttk.Frame(root)
frame_botones.grid(row=11, column=0, columnspan=2, pady=(24, 0))


ttk.Button(
    frame_botones,
    text="Guardar",
    width=14,
    command=guardar_nomina,
    style="Accent.TButton",
).grid(row=0, column=0, padx=10)

ttk.Button(
    frame_botones,
    text="Limpiar",
    width=14,
    command=limpiar_formulario,
).grid(row=0, column=1, padx=10)

entry_cod_empleado.focus()
root.mainloop()
