import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def calcular_porcentajes(input_file, output_file, col_item, col_bodega, col_cantidad):
    try:
        # Cargar el archivo Excel
        df = pd.read_excel(input_file)
        
        # Propagar los valores de la columna "Item" hacia abajo para completar los NaN
        df[col_item] = df[col_item].ffill()
        
        # Convertir la columna de cantidad a numérico, reemplazando NaN, caracteres no numéricos y negativos por 0
        df[col_cantidad] = pd.to_numeric(df[col_cantidad], errors='coerce').fillna(0)
        df[col_cantidad] = df[col_cantidad].apply(lambda x: max(x, 0))  # Reemplazar negativos con 0
        
        # Calcular el total de inventario por ítem
        total_por_item = df.groupby(col_item)[col_cantidad].transform('sum')
        
        # Convertir total a numérico y evitar división por cero
        total_por_item = pd.to_numeric(total_por_item, errors='coerce').fillna(0)
        total_por_item.replace(0, 1, inplace=True)  # Evitar división por cero, usando 1 como valor mínimo
        
        # Calcular el porcentaje de participación por bodega
        df['% PARTICIPACION'] = (df[col_cantidad] / total_por_item) * 100
        
        # Redondear a un decimal y agregar el símbolo de porcentaje
        df['% PARTICIPACION'] = df['% PARTICIPACION'].round(1).astype(str) + '%'
        
        # Guardar el DataFrame con los porcentajes en un nuevo archivo Excel
        df.to_excel(output_file, index=False)
        messagebox.showinfo("Éxito", f"Archivo guardado en: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

def seleccionar_columnas(input_file, archivo_salida):
    df = pd.read_excel(input_file)
    columnas = df.columns.tolist()
    
    def confirmar_seleccion():
        col_item = combo_item.get()
        col_bodega = combo_bodega.get()
        col_cantidad = combo_cantidad.get()
        root_columns.destroy()
        calcular_porcentajes(input_file, archivo_salida, col_item, col_bodega, col_cantidad)
    
    root_columns = tk.Toplevel()
    root_columns.title("Seleccionar Columnas")
    
    tk.Label(root_columns, text="Seleccione la columna para 'Item'").pack()
    combo_item = ttk.Combobox(root_columns, values=columnas)
    combo_item.pack()
    
    tk.Label(root_columns, text="Seleccione la columna para 'Bodega'").pack()
    combo_bodega = ttk.Combobox(root_columns, values=columnas)
    combo_bodega.pack()
    
    tk.Label(root_columns, text="Seleccione la columna para 'Cuenta de Cantidad Inv.'").pack()
    combo_cantidad = ttk.Combobox(root_columns, values=columnas)
    combo_cantidad.pack()
    
    btn_confirmar = tk.Button(root_columns, text="Confirmar", command=confirmar_seleccion)
    btn_confirmar.pack()

def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    
    archivo_entrada = filedialog.askopenfilename(title="Seleccionar archivo de inventario", filetypes=[("Archivos Excel", "*.xlsx")])
    if not archivo_entrada:
        messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
        return
    
    archivo_salida = filedialog.asksaveasfilename(title="Guardar archivo con porcentajes", defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
    if not archivo_salida:
        messagebox.showwarning("Advertencia", "No se seleccionó un nombre para el archivo de salida.")
        return
    
    seleccionar_columnas(archivo_entrada, archivo_salida)

def iniciar_interfaz():
    root = tk.Tk()
    root.title("Calculadora de Porcentajes de Inventario")
    root.geometry("400x200")
    
    label = tk.Label(root, text="Seleccione un archivo Excel de inventario", font=("Arial", 12))
    label.pack(pady=20)
    
    boton_seleccionar = tk.Button(root, text="Seleccionar Archivo", command=seleccionar_archivo, font=("Arial", 12), bg="lightblue")
    boton_seleccionar.pack()
    
    root.mainloop()
    
# Iniciar la interfaz gráfica
iniciar_interfaz()