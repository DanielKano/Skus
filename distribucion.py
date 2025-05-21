import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import math

def calcular_distribucion(input_file, output_file, col_item, col_bodega, col_ventas, col_inventario):
    try:
        # Leer datos de la primera hoja (ventas)
        df = pd.read_excel(input_file, sheet_name=0)
        df[col_item] = df[col_item].ffill()
        df[col_ventas] = pd.to_numeric(df[col_ventas], errors='coerce').fillna(0)
        df[col_ventas] = df[col_ventas].apply(lambda x: max(x, 0))

        # Calcular total por ítem
        total_por_item = df.groupby(col_item)[col_ventas].transform('sum')
        total_por_item = total_por_item.replace(0, 1)

        # Calcular % participación
        df['% PARTICIPACION'] = (df[col_ventas] / total_por_item * 100).round(1).astype(str) + '%'
        df['participacion_decimal'] = df[col_ventas] / total_por_item

        # Leer inventario desde la segunda hoja
        inventario_df = pd.read_excel(input_file, sheet_name=1)
        inventario_df[col_inventario] = pd.to_numeric(inventario_df[col_inventario], errors='coerce').fillna(0)

        # Crear nueva columna para la distribución
        df['DISTRIBUCION'] = 0

        # Procesar distribución por cada ítem
        for item in df[col_item].unique():
            df_item = df[df[col_item] == item].copy()
            inv_disponible = inventario_df[inventario_df[col_item] == item][col_inventario].sum()

            if inv_disponible == 0:
                df.loc[df[col_item] == item, 'DISTRIBUCION'] = 0
                continue

            participaciones = df_item['participacion_decimal'].tolist()

            if all(p == 0 for p in participaciones):
                # Reparto equitativo si nadie participó
                tiendas = df_item.index.tolist()
                cantidad = min(inv_disponible, len(tiendas))
                for i in range(cantidad):
                    df.at[tiendas[i], 'DISTRIBUCION'] = 1
            else:
                # Reparto proporcional
                cantidades = [math.floor(p * inv_disponible) for p in participaciones]
                sobrante = int(inv_disponible - sum(cantidades))

                # Asignar las cantidades iniciales
                for idx, val in zip(df_item.index, cantidades):
                    df.at[idx, 'DISTRIBUCION'] = val

                # Distribuir sobrante a los de mayor participación
                df_item['residuo'] = df_item['participacion_decimal'] * inv_disponible - df['DISTRIBUCION']
                df_item = df_item.sort_values(by='participacion_decimal', ascending=False)
                for idx in df_item.index:
                    if sobrante <= 0:
                        break
                    df.at[idx, 'DISTRIBUCION'] += 1
                    sobrante -= 1

        # Guardar archivo con dos hojas
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.drop(columns=['participacion_decimal'], errors='ignore').to_excel(writer, sheet_name='Participacion', index=False)
            df[[col_item, col_bodega, 'DISTRIBUCION']].to_excel(writer, sheet_name='Distribucion', index=False)

        messagebox.showinfo("Éxito", f"Archivo guardado correctamente en:\n{output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

def seleccionar_columnas(input_file, archivo_salida):
    df = pd.read_excel(input_file, sheet_name=0)
    columnas = df.columns.tolist()

    def confirmar():
        col_item = combo_item.get()
        col_bodega = combo_bodega.get()
        col_ventas = combo_ventas.get()
        col_inventario = combo_inventario.get()
        top.destroy()
        calcular_distribucion(input_file, archivo_salida, col_item, col_bodega, col_ventas, col_inventario)

    top = tk.Toplevel()
    top.title("Seleccionar Columnas")

    tk.Label(top, text="Columna de 'Item'").pack()
    combo_item = ttk.Combobox(top, values=columnas)
    combo_item.pack()

    tk.Label(top, text="Columna de 'Bodega'").pack()
    combo_bodega = ttk.Combobox(top, values=columnas)
    combo_bodega.pack()

    tk.Label(top, text="Columna de 'Ventas / Cantidad'").pack()
    combo_ventas = ttk.Combobox(top, values=columnas)
    combo_ventas.pack()

    tk.Label(top, text="Columna de 'Inventario' (hoja 2)").pack()
    combo_inventario = ttk.Combobox(top, values=columnas)
    combo_inventario.pack()

    tk.Button(top, text="Confirmar", command=confirmar).pack(pady=10)

def seleccionar_archivo():
    root.withdraw()
    archivo_entrada = filedialog.askopenfilename(title="Seleccionar archivo Excel con ventas e inventario", filetypes=[("Archivos Excel", "*.xlsx")])
    if not archivo_entrada:
        messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
        return

    archivo_salida = filedialog.asksaveasfilename(title="Guardar archivo de salida", defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
    if not archivo_salida:
        messagebox.showwarning("Advertencia", "No se seleccionó un archivo de salida.")
        return

    seleccionar_columnas(archivo_entrada, archivo_salida)

def iniciar_app():
    global root
    root = tk.Tk()
    root.title("Distribuidor de Inventario por Participación")
    root.geometry("500x200")

    tk.Label(root, text="Seleccione el archivo Excel con ventas (hoja 1) e inventario (hoja 2)", wraplength=400, font=("Arial", 12)).pack(pady=20)

    tk.Button(root, text="Seleccionar archivo", command=seleccionar_archivo, font=("Arial", 12), bg="lightblue").pack()

    root.mainloop()

# Iniciar aplicación
iniciar_app()
