import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def calcular_porcentajes(input_file, output_file):
    try:
        # Cargar el archivo Excel
        df = pd.read_excel(input_file)
        
        # Propagar los valores de la columna "Item" hacia abajo para completar los NaN
        df['Item'] = df['Item'].ffill()
        
        # Calcular el total de inventario por ítem
        total_por_item = df.groupby('Item')['Cuenta de Cantidad Inv.'].transform('sum')
        
        # Calcular el porcentaje de participación por bodega
        df['% PARTICIPACION'] = (df['Cuenta de Cantidad Inv.'] / total_por_item) * 100
        
        # Redondear a un decimal y agregar el símbolo de porcentaje
        df['% PARTICIPACION'] = df['% PARTICIPACION'].round(1).astype(str) + '%'
        
        # Guardar el DataFrame con los porcentajes en un nuevo archivo Excel
        df.to_excel(output_file, index=False)
        messagebox.showinfo("Éxito", f"Archivo guardado en: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

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
    
    calcular_porcentajes(archivo_entrada, archivo_salida)

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
