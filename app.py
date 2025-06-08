import pandas as pd
from datetime import timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Variables globales
estado_label = None

def cargar_archivo():
    global estado_label
    ruta_archivo = filedialog.askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if not ruta_archivo:
        return

    try:
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
        df['Contrato Fin'] = pd.to_datetime(df['Contrato Fin'], format='%d/%m/%Y', errors='coerce')
        app.dataframe_original = df
        estado_label.config(text=f"Estado: Archivo cargado ({len(df)} registros)", fg="green")
        messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
    except Exception as e:
        estado_label.config(text=f"Estado: Error al cargar archivo", fg="red")
        messagebox.showerror("Error", f"No se pudo leer el archivo.\n{e}")

def procesar_archivo():

    global estado_label

    if not hasattr(app, "dataframe_original"):
        estado_label.config(text="Estado: No hay archivo cargado", fg="red")
        messagebox.showerror("Error", "Primero debes cargar un archivo.")
        return

    try:
        df = app.dataframe_original.copy()
        hoy = pd.Timestamp.today().normalize()
        modo = filtro_var.get()

        if modo not in ["dias", "mes"]:
            messagebox.showerror("Error", "Selecciona un tipo de filtrado válido.")
            return

        if modo == "dias":
            dias_horizonte = int(slider_dias.get())
            fecha_horizonte = hoy + timedelta(days=dias_horizonte)
            resultado = df[(df['Contrato Fin'] >= hoy) & (df['Contrato Fin'] <= fecha_horizonte)].copy()
            resultado['Dias_faltantes'] = (resultado['Contrato Fin'] - hoy).dt.days
            resultado['Finaliza el'] = resultado['Contrato Fin'].dt.strftime('%d/%m/%Y')

        elif modo == "mes":

            if hoy.month == 12:
                mes_siguiente = 1
                sigYear = hoy.year + 1
            else:
                mes_siguiente = hoy.month + 1
                sigYear = hoy.year

            primer_dia_mes = pd.Timestamp(sigYear, mes_siguiente, 1)
            primer_dia_mes_despues = primer_dia_mes + pd.offsets.MonthBegin(1)
            ultimo_dia_mes = primer_dia_mes_despues - timedelta(days=1)

            resultado = df[(df['Contrato Fin'] >= primer_dia_mes) & (df['Contrato Fin'] <= ultimo_dia_mes)].copy()
            resultado['Dias_faltantes'] = (resultado['Contrato Fin'] - hoy).dt.days
            resultado['Finaliza el'] = resultado['Contrato Fin'].dt.strftime('%d/%m/%Y')

        resultado = resultado.sort_values(by='Contrato Fin')
        resultado['Contrato Fin'] = resultado['Contrato Fin'].dt.strftime('%d/%m/%Y')
        app.resultado = resultado

        # Mostrar todos los registros
        vista.delete(*vista.get_children())
        if resultado.empty:
            messagebox.showinfo("Información", "No se encontraron registros que cumplan con los criterios.")
            return
            
        # Mostrar todos los registros en la tabla
        for _, row in resultado.iterrows():
            valores = [row.get(col, "") for col in ["Numero documento", "Contrato Fin", "Dias_faltantes"]]
            vista.insert('', 'end', values=valores)

        estado_label.config(text=f"Estado: Se encontraron {len(resultado)} registros", fg="green")
        messagebox.showinfo("Éxito", f"Datos procesados correctamente. Se encontraron {len(resultado)} registros.")
    except Exception as e:
        messagebox.showerror("Error durante el procesamiento", str(e))

def guardar_resultado():

    if not hasattr(app, "resultado") or app.resultado.empty:
        messagebox.showerror("Error", "No hay datos procesados para guardar.")
        return

    ruta_guardado = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivo Excel", "*.xlsx")],
        title="Guardar archivo como"
    )
    if ruta_guardado:
        try:
            app.resultado.to_excel(ruta_guardado, index=False)
            messagebox.showinfo("Guardado", f"Archivo guardado en:\n{ruta_guardado}")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))


# Crear ventana principal
app = tk.Tk()
app.title("Filtrado de Contratos")
app.geometry("850x550")

# Frame para el botón de carga y el estado
frame_carga = tk.Frame(app)
frame_carga.pack(pady=10)

# Botón para cargar archivo y label de estado
tk.Button(frame_carga, text="Cargar Excel", command=cargar_archivo).pack(side="left", padx=(0, 10))
estado_label = tk.Label(frame_carga, text="Estado: No hay archivo cargado", fg="red")
estado_label.pack(side="left")

# Selector de filtro
filtro_var = tk.StringVar(value="dias")
label_tipo_filtrado = tk.Label(app, text="Selecciona tipo de filtrado:")
label_tipo_filtrado.pack()

# Create the slider frame but don't pack it yet
frame_slider = tk.Frame(app)
tk.Label(frame_slider, text="Selecciona días de horizonte:").pack(side="left")
slider_dias = tk.Scale(frame_slider, from_=1, to=180, orient="horizontal")
slider_dias.set(7)
slider_dias.pack(side="left")

# Function to show/hide slider based on filter type
def actualizar_slider(*args):
    if filtro_var.get() == "dias":
        # Ensure it's always packed in the same position
        frame_slider.pack_forget()  # First remove it if it exists
        frame_slider.pack(after=label_tipo_filtrado, pady=10)  # Pack it after a specific widget
    else:
        frame_slider.pack_forget()

# Connect the function to the StringVar
filtro_var.trace_add("write", actualizar_slider)

# Create radio buttons
tk.Radiobutton(app, text="Próximos N días", variable=filtro_var, value="dias").pack()
tk.Radiobutton(app, text="Mes siguiente completo", variable=filtro_var, value="mes").pack()

# Initial call to set correct visibility
actualizar_slider()

# Procesar
tk.Button(app, text="Procesar Datos", command=procesar_archivo).pack(pady=10)

# Frame para contener la tabla y las barras de desplazamiento
frame_tabla = tk.Frame(app)
frame_tabla.pack(pady=10, fill="both", expand=True)

# Tabla de visualización completa
cols = ["Número de documento", "Fecha de fin del contrato", "Días faltantes"]
vista = ttk.Treeview(frame_tabla, columns=cols, show='headings', height=15)

# Configurar columnas con ancho adecuado
for col in cols:
    vista.heading(col, text=col)
    vista.column(col, width=200)

# Añadir scrollbars
scrollbar_y = ttk.Scrollbar(frame_tabla, orient="vertical", command=vista.yview)
scrollbar_x = ttk.Scrollbar(frame_tabla, orient="horizontal", command=vista.xview)
vista.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

# Posicionar los elementos con grid para mejor control
vista.grid(row=0, column=0, sticky="nsew")
scrollbar_y.grid(row=0, column=1, sticky="ns")
scrollbar_x.grid(row=1, column=0, sticky="ew")

# Configurar la expansión del grid
frame_tabla.grid_rowconfigure(0, weight=1)
frame_tabla.grid_columnconfigure(0, weight=1)

# Guardar archivo
tk.Button(app, text="Guardar Resultado", command=guardar_resultado).pack(pady=10)

# Ejecutar
app.mainloop()
