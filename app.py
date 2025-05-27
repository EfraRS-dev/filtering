import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def cargar_archivo():
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
        messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo.\n{e}")

def procesar_archivo():
    if not hasattr(app, "dataframe_original"):
        messagebox.showerror("Error", "Primero debes cargar un archivo.")
        return

    try:
        df = app.dataframe_original.copy()
        hoy = pd.Timestamp.today().normalize()
        modo = filtro_var.get()

        if modo == "dias":
            dias_horizonte = int(slider_dias.get())
            fecha_horizonte = hoy + timedelta(days=dias_horizonte)
            resultado = df[(df['Contrato Fin'] >= hoy) & (df['Contrato Fin'] <= fecha_horizonte)].copy()
            resultado['Dias_faltantes'] = (resultado['Contrato Fin'] - hoy).dt.days

        elif modo == "mes":
            if hoy.month == 12:
                mes_siguiente = 1
                año_siguiente = hoy.year + 1
            else:
                mes_siguiente = hoy.month + 1
                año_siguiente = hoy.year

            primer_dia_mes = pd.Timestamp(año_siguiente, mes_siguiente, 1)
            primer_dia_mes_despues = primer_dia_mes + pd.offsets.MonthBegin(1)
            ultimo_dia_mes = primer_dia_mes_despues - timedelta(days=1)

            resultado = df[(df['Contrato Fin'] >= primer_dia_mes) & (df['Contrato Fin'] <= ultimo_dia_mes)].copy()
            resultado['Dias_faltantes'] = (resultado['Contrato Fin'] - hoy).dt.days

        resultado = resultado.sort_values(by='Contrato Fin')
        resultado['Contrato Fin'] = resultado['Contrato Fin'].dt.strftime('%d/%m/%Y')
        app.resultado = resultado

        # Mostrar vista previa
        vista.delete(*vista.get_children())
        for _, row in resultado.head(10).iterrows():
            valores = [row.get(col, "") for col in ["Numero documento", "Contrato Fin", "Dias_faltantes"]]
            vista.insert('', 'end', values=valores)

        messagebox.showinfo("Éxito", "Datos procesados correctamente.")
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

# Botón para cargar archivo
tk.Button(app, text="Cargar Excel", command=cargar_archivo).pack(pady=10)


# Selector de filtro
filtro_var = tk.StringVar(value="dias")
tk.Label(app, text="Selecciona tipo de filtrado:").pack()
tk.Radiobutton(app, text="Próximos N días", variable=filtro_var, value="dias").pack()
tk.Radiobutton(app, text="Mes siguiente completo", variable=filtro_var, value="mes").pack()

# Slider de días
frame_slider = tk.Frame(app)
tk.Label(frame_slider, text="Selecciona días de horizonte:").pack(side="left")
slider_dias = tk.Scale(frame_slider, from_=1, to=180, orient="horizontal")
slider_dias.set(7)
slider_dias.pack(side="left")
frame_slider.pack(pady=10)

# Procesar
tk.Button(app, text="Procesar Datos", command=procesar_archivo).pack(pady=10)

# Tabla de vista previa
cols = ["Número de documento", "Fecha de fin del contrato", "Días faltantes"]
vista = ttk.Treeview(app, columns=cols, show='headings', height=10)
for col in cols:
    vista.heading(col, text=col)
    vista.column(col, width=200)
vista.pack(pady=10, fill="x")

# Guardar archivo
tk.Button(app, text="Guardar Resultado", command=guardar_resultado).pack(pady=10)

# Ejecutar
app.mainloop()
