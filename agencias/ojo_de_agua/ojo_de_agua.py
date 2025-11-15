
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from fpdf import FPDF
import os
import sys
import unicodedata
from pathlib import Path
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph

# --------------------
# RUTAS Y CONST
# --------------------
CARPETA_DATOS = Path.home() / "Desktop" / "Archivos"
CARPETA_EXCEL = CARPETA_DATOS / "Excel"
CARPETA_EXPORT = CARPETA_DATOS / "Export"
LOGO_DIR = CARPETA_DATOS / "LOGO"

# Crear carpetas si no existen
os.makedirs(CARPETA_DATOS, exist_ok=True)
os.makedirs(CARPETA_EXCEL, exist_ok=True)
os.makedirs(CARPETA_EXPORT, exist_ok=True)
os.makedirs(LOGO_DIR, exist_ok=True)

ARCHIVO_INVENTARIO = CARPETA_EXCEL / "inventario.xlsx"
ARCHIVO_PEDIDOS = CARPETA_EXCEL / "pedidos.xlsx"
ARCHIVO_VENTAS = CARPETA_EXCEL / "ventas.xlsx"
ARCHIVO_TALLER = CARPETA_EXCEL / "taller.xlsx"
ARCHIVO_COTIZACIONES = CARPETA_EXCEL / "cotizaciones.xlsx"

# --------------------
# UTILIDADES
# --------------------
def quitar_acentos(texto):
    if not isinstance(texto, str):
        return texto
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def _create_empty_excel(path: Path, columns):
    df = pd.DataFrame(columns=columns)
    df.to_excel(path, index=False, engine="openpyxl")

def load_file(path: Path, columns):
    if path.exists():
        try:
            return pd.read_excel(path, engine="openpyxl", dtype=str).fillna("")
        except Exception:
            return pd.DataFrame(columns=columns)
    else:
        _create_empty_excel(path, columns)
        return pd.DataFrame(columns=columns)

def load_inventario_file():
    return load_file(ARCHIVO_INVENTARIO, ["codigo", "descripcion", "ubicacion", "stock", "precio"])

def load_ventas_file():
    return load_file(ARCHIVO_VENTAS, ["fecha", "codigo", "descripcion", "cantidad", "precio", "total"])

def load_taller_file():
    return load_file(ARCHIVO_TALLER, ["fecha", "codigo", "descripcion", "cantidad", "estado"])

def save_df(path: Path, df: pd.DataFrame):
    df.to_excel(path, index=False, engine="openpyxl")

def obtener_estado_codigo(codigo, cantidad_total):
    try:
        return int(cantidad_total), 0
    except Exception:
        return 0, 0

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def habilitar_copia_treeview(tree):
    """
    Permite copiar filas del Treeview al portapapeles con Ctrl+C
    """
    def copiar(event):
        seleccion = tree.selection()
        if not seleccion: return
        texto = ""
        for item in seleccion:
            texto += "\t".join([str(tree.set(item, col)) for col in tree["columns"]]) + "\n"
        tree.clipboard_clear()
        tree.clipboard_append(texto)
    tree.bind("<Control-c>", copiar)


# --------------------
# TREEVIEW ESTILO PERSONALIZADO
# --------------------
def estilo_treeview(root, fondo="#1ec2df"):
    style = ttk.Style(root)
    style.theme_use('clam')

    if fondo.lower() in ["#dd1111",'white']:
        fg_color = 'black'
        sel_bg = "#abe01a"
    else:
        fg_color = 'white'
        sel_bg = '#016630'

    style.configure("Treeview",
                    background=fondo,
                    foreground=fg_color,
                    fieldbackground=fondo,
                    rowheight=24,
                    font=("Segoe UI", 10))

    style.configure("Treeview.Heading",
                    background=fondo,
                    foreground=fg_color,
                    font=("Segoe UI", 10, "bold"))

    style.map("Treeview",
              background=[('selected', sel_bg)],
              foreground=[('selected', fg_color)])


# --------------------
# STYLES GENERALES
# --------------------
def aplicar_estilos(root):
    estilo_treeview(root, fondo="#11d8be")  # Fondo blanco por defecto
    style = ttk.Style(root)
    root.configure(bg="#0753c4")  # Fondo general blanco




class Stock(ttk.Frame):
    def __init__(self, parent, controller=None):
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text='STOCK', font=('Segoe UI', 12, 'bold')).pack(anchor='w', padx=6, pady=6)

        # -------------------------------------------------
        # BUSCAR POR CÓDIGO
        # -------------------------------------------------
        frame = ttk.Frame(self)
        frame.pack(fill='x', padx=6, pady=(5,0))

        ttk.Label(frame, text='Código:').grid(row=0, column=0, sticky='w')
        self.entry_codigo = ttk.Entry(frame, width=25)
        self.entry_codigo.grid(row=0, column=1, padx=4)

        ttk.Button(frame, text='Buscar', command=self.buscar_codigo).grid(row=0, column=2, padx=4)

        # -------------------------------------------------
        # BUSCAR POR DESCRIPCIÓN
        # -------------------------------------------------
        frame2 = ttk.Frame(self)
        frame2.pack(fill='x', padx=6, pady=(6,0))

        ttk.Label(frame2, text='Descripción:').grid(row=0, column=0, sticky='w')
        self.entry_desc = ttk.Entry(frame2, width=40)
        self.entry_desc.grid(row=0, column=1, padx=4)

        ttk.Button(frame2, text='Buscar descripción', command=self.buscar_descripcion).grid(row=0, column=2, padx=4)

        # -------------------------------------------------
        # BOTONES ARRIBA A LA DERECHA
        # -------------------------------------------------
        frame_top_btns = ttk.Frame(self)
        frame_top_btns.pack(fill='x', padx=6)

        ttk.Button(frame_top_btns, text='Importar', command=self.importar_inventario).pack(side='right', padx=4)
        ttk.Button(frame_top_btns, text='Exportar', command=self.exportar_inventario).pack(side='right', padx=4)

        # ------------------
        # AGREGAR / DESCONTAR
        # ------------------
        frame_desc = ttk.LabelFrame(self, text="AGREGAR / DESCONTAR")
        frame_desc.pack(fill='x', padx=6, pady=6)

        ttk.Label(frame_desc, text="Código:").grid(row=0, column=0, padx=4, pady=4, sticky='w')
        self.desc_codigo = ttk.Entry(frame_desc, width=15)
        self.desc_codigo.grid(row=0, column=1, padx=4, pady=4)

        ttk.Label(frame_desc, text="Cantidad:").grid(row=0, column=2, padx=4, pady=4, sticky='w')
        self.desc_cantidad = ttk.Entry(frame_desc, width=10)
        self.desc_cantidad.grid(row=0, column=3, padx=4, pady=4)

        ttk.Button(frame_desc, text="Agregar", command=self.agregar_refaccion).grid(row=0, column=4, padx=4, pady=4)
        ttk.Button(frame_desc, text="Descontar", command=self.descontar_refaccion).grid(row=0, column=5, padx=4, pady=4)

        # ------------------
        # AGREGAR / BORRAR ARTÍCULO
        # ------------------
        frame_art = ttk.LabelFrame(self, text="Agregar / Borrar Artículo")
        frame_art.pack(fill='x', padx=6, pady=6)

        ttk.Label(frame_art, text="Código:").grid(row=0, column=0, padx=4, pady=2)
        self.art_codigo = ttk.Entry(frame_art, width=12)
        self.art_codigo.grid(row=0, column=1, padx=4, pady=2)

        ttk.Label(frame_art, text="Descripción:").grid(row=0, column=2, padx=4, pady=2)
        self.art_desc = ttk.Entry(frame_art, width=20)
        self.art_desc.grid(row=0, column=3, padx=4, pady=2)

        ttk.Label(frame_art, text="Ubicación:").grid(row=0, column=4, padx=4, pady=2)
        self.art_ubi = ttk.Entry(frame_art, width=12)
        self.art_ubi.grid(row=0, column=5, padx=4, pady=2)

        ttk.Label(frame_art, text="Stock:").grid(row=1, column=0, padx=4, pady=2)
        self.art_stock = ttk.Entry(frame_art, width=12)
        self.art_stock.grid(row=1, column=1, padx=4, pady=2)

        ttk.Label(frame_art, text="Precio:").grid(row=1, column=2, padx=4, pady=2)
        self.art_precio = ttk.Entry(frame_art, width=12)
        self.art_precio.grid(row=1, column=3, padx=4, pady=2)

        ttk.Button(frame_art, text="Agregar/Actualizar", command=self.agregar_articulo).grid(row=1, column=4, padx=4, pady=2)
        ttk.Button(frame_art, text="Borrar Seleccionado", command=self.borrar_seleccionado).grid(row=1, column=5, padx=4, pady=2)

        # ------------------
        # TREEVIEW PRINCIPAL
        # ------------------
        cols = ["codigo", "descripcion", "ubicacion", "stock", "precio", "libres", "en_taller", "nuevas_entradas"]
        self.tree = ttk.Treeview(self, columns=cols, show='headings', height=14)
        for c in cols:
            self.tree.heading(c, text=c.capitalize())
            self.tree.column(c, width=120, anchor='center')
        self.tree.pack(fill='both', expand=True, padx=6, pady=6)

        self.cargar_datos()

    # --------------------------------------------------------
    # CARGAR DATOS
    # --------------------------------------------------------
    def cargar_datos(self):
        df = load_inventario_file()

        for col in ['libres', 'en_taller', 'nuevas_entradas']:
            df[col] = df.get(col, 0)

        if 'codigo' in df.columns:
            df['codigo_norm'] = df['codigo'].astype(str).apply(lambda x: quitar_acentos(x).strip().upper())
            df.drop_duplicates(subset=['codigo_norm'], keep='last', inplace=True)
            df.drop(columns=['codigo_norm'], inplace=True)

        self.tree.delete(*self.tree.get_children())

        for _, r in df.iterrows():
            stock = int(r.get("stock", 0))
            libres, en_taller = obtener_estado_codigo(r.get("codigo", ""), stock)

            vals = (
                r.get('codigo',''),
                r.get('descripcion',''),
                r.get('ubicacion',''),
                stock,
                r.get('precio',0),
                libres,
                en_taller,
                r.get('nuevas_entradas',0)
            )
            self.tree.insert('', 'end', values=vals)

    # --------------------------------------------------------
    # BUSCAR POR CÓDIGO (AUTORRELLENO)
    # --------------------------------------------------------
    def buscar_codigo(self):
        codigo = quitar_acentos(self.entry_codigo.get().strip()).upper()
        if not codigo:
            messagebox.showinfo('Atención','Escribe un código para buscar')
            return

        df = load_inventario_file()
        if 'codigo' not in df.columns: return

        df['codigo_clean'] = df['codigo'].astype(str).apply(
            lambda x: quitar_acentos(x).upper()
        )
        r = df[df['codigo_clean']==codigo]

        if r.empty:
            messagebox.showinfo('Atención', f'Código {codigo} no encontrado')
            return

        row = r.iloc[0]
        stock = int(row.get('stock') or 0)
        libres, en_taller = obtener_estado_codigo(row.get('codigo',''), stock)

        # mostrar solo 1 resultado
        self.tree.delete(*self.tree.get_children())
        self.tree.insert('', 'end', values=(
            row.get('codigo',''),
            row.get('descripcion',''),
            row.get('ubicacion',''),
            row.get('stock',0),
            row.get('precio',0),
            libres,
            en_taller,
            row.get('nuevas_entradas',0)
        ))

        # rellenar formularios
        self.art_codigo.delete(0,'end'); self.art_codigo.insert(0,row.get('codigo',''))
        self.art_desc.delete(0,'end'); self.art_desc.insert(0,row.get('descripcion',''))
        self.art_ubi.delete(0,'end'); self.art_ubi.insert(0,row.get('ubicacion',''))
        self.art_stock.delete(0,'end'); self.art_stock.insert(0,str(row.get('stock','')))
        self.art_precio.delete(0,'end'); self.art_precio.insert(0,str(row.get('precio','')))

    # --------------------------------------------------------
    # BUSCAR POR DESCRIPCIÓN
    # --------------------------------------------------------
    def buscar_descripcion(self):
        desc = quitar_acentos(self.entry_desc.get().strip()).upper()
        if not desc:
            messagebox.showinfo('Atención','Escribe descripción')
            return

        df = load_inventario_file()
        if 'descripcion' not in df.columns: return

        df['desc_clean'] = df['descripcion'].astype(str).apply(
            lambda x: quitar_acentos(x).upper()
        )
        r = df[df['desc_clean'].str.contains(desc, na=False)]

        self.tree.delete(*self.tree.get_children())

        for _, row in r.iterrows():
            stock = int(row.get("stock") or 0)
            libres, en_taller = obtener_estado_codigo(row.get('codigo',''), stock)

            self.tree.insert('', 'end', values=(
                row.get('codigo',''),
                row.get('descripcion',''),
                row.get('ubicacion',''),
                stock,
                row.get('precio',0),
                libres,
                en_taller,
                row.get('nuevas_entradas',0)
            ))

    # --------------------------------------------------------
    # AGREGAR / DESCONTAR STOCK
    # --------------------------------------------------------
    def agregar_refaccion(self):
        codigo = self.desc_codigo.get().strip().upper()
        try: cantidad = int(self.desc_cantidad.get())
        except: cantidad = 0

        if not codigo or cantidad <= 0:
            messagebox.showwarning("Atención", "Ingrese código y cantidad válida")
            return

        df = load_inventario_file()
        mask = df['codigo'].astype(str).str.upper() == codigo

        if mask.any():
            idx = df[mask].index[0]
            df.at[idx, 'stock'] = int(df.at[idx, 'stock']) + cantidad
            save_df(ARCHIVO_INVENTARIO, df)
            self.cargar_datos()
            messagebox.showinfo("OK", f"Agregado {cantidad} a {codigo}")
        else:
            messagebox.showwarning("No encontrado", "Código no encontrado")

    def descontar_refaccion(self):
        codigo = self.desc_codigo.get().strip().upper()
        try: cantidad = int(self.desc_cantidad.get())
        except: cantidad = 0

        if not codigo or cantidad <= 0:
            messagebox.showwarning("Atención", "Ingrese código y cantidad válida")
            return

        df = load_inventario_file()
        mask = df['codigo'].astype(str).str.upper() == codigo

        if mask.any():
            idx = df[mask].index[0]
            df.at[idx, 'stock'] = max(int(df.at[idx, 'stock']) - cantidad, 0)
            save_df(ARCHIVO_INVENTARIO, df)
            self.cargar_datos()
            messagebox.showinfo("OK", f"Descontado {cantidad} de {codigo}")
        else:
            messagebox.showwarning("No encontrado", "Código no encontrado")

    # --------------------------------------------------------
    # AGREGAR / BORRAR ARTÍCULO
    # --------------------------------------------------------
    def agregar_articulo(self):
        codigo = self.art_codigo.get().strip()
        desc = self.art_desc.get().strip()
        ubi = self.art_ubi.get().strip()

        try: stock = int(self.art_stock.get())
        except: stock = 0

        try: precio = float(self.art_precio.get())
        except: precio = 0.0

        if not codigo:
            messagebox.showwarning("Atención", "Código requerido")
            return

        df = load_inventario_file()
        mask = df['codigo'].astype(str).str.upper() == codigo.upper()

        if mask.any():
            idx = df[mask].index[0]
            df.at[idx,'descripcion'] = desc
            df.at[idx,'ubicacion'] = ubi
            df.at[idx,'stock'] = stock
            df.at[idx,'precio'] = precio
        else:
            df.loc[len(df)] = [codigo, desc, ubi, stock, precio]

        save_df(ARCHIVO_INVENTARIO, df)
        self.cargar_datos()
        messagebox.showinfo("OK","Artículo agregado/actualizado")

    def borrar_seleccionado(self):
        sel = self.tree.selection()
        if not sel: return

        df = load_inventario_file()

        for s in sel:
            codigo = self.tree.item(s)['values'][0]
            df = df[df['codigo'].astype(str).str.upper() != codigo.upper()]

        save_df(ARCHIVO_INVENTARIO, df)
        self.cargar_datos()
        messagebox.showinfo("OK","Artículo(s) borrado(s)")

    # --------------------------------------------------------
    # IMPORTAR / EXPORTAR
    # --------------------------------------------------------
    def importar_inventario(self):
        archivo = filedialog.askopenfilename(
            title='Seleccionar Excel',
            filetypes=[('Excel', '*.xlsx;*.xls')]
        )
        if not archivo: return

        try:
            df_new = pd.read_excel(archivo, engine='openpyxl', dtype=str).fillna('')
            df_old = load_inventario_file()
            df_comb = pd.concat([df_old, df_new], ignore_index=True)

            if 'codigo' in df_comb.columns:
                df_comb.drop_duplicates(subset=['codigo'], keep='last', inplace=True)

            save_df(ARCHIVO_INVENTARIO, df_comb)
            self.cargar_datos()
            messagebox.showinfo('Éxito','Inventario importado')

        except Exception as e:
            messagebox.showerror('Error', str(e))

    def exportar_inventario(self):
        df = load_inventario_file()
        ruta = os.path.join(CARPETA_EXPORT, 'inventario_exportado.xlsx')
        df.to_excel(ruta, index=False, engine='openpyxl')
        messagebox.showinfo('Éxito', f'Exportado en:\n{ruta}')


class Ventas(ttk.Frame):
    def __init__(self, parent, controller=None):
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text='VENTAS', font=('Segoe UI', 12, 'bold')).pack(anchor='w', padx=6, pady=6)

        # ------------------
        # Panel de ingreso de venta
        # ------------------
        frame_venta = ttk.LabelFrame(self, text="Agregar a la Venta")
        frame_venta.pack(fill='x', padx=6, pady=6)

        ttk.Label(frame_venta, text="Código:").grid(row=0, column=0, padx=4, pady=4)
        self.cod_entry = ttk.Entry(frame_venta, width=15)
        self.cod_entry.grid(row=0, column=1, padx=4, pady=4)
        self.cod_entry.bind("<FocusOut>", self.completar_datos)  # Al salir del entry, completa descripción y precio

        ttk.Label(frame_venta, text="Descripción:").grid(row=0, column=2, padx=4, pady=4)
        self.desc_entry = ttk.Entry(frame_venta, width=30, state='readonly')
        self.desc_entry.grid(row=0, column=3, padx=4, pady=4)

        ttk.Label(frame_venta, text="Precio:").grid(row=0, column=4, padx=4, pady=4)
        self.precio_entry = ttk.Entry(frame_venta, width=10, state='readonly')
        self.precio_entry.grid(row=0, column=5, padx=4, pady=4)

        ttk.Label(frame_venta, text="Cantidad:").grid(row=0, column=6, padx=4, pady=4)
        self.cant_entry = ttk.Entry(frame_venta, width=8)
        self.cant_entry.grid(row=0, column=7, padx=4, pady=4)

        ttk.Button(frame_venta, text="Agregar a Venta", command=self.agregar_a_venta).grid(row=0, column=8, padx=4, pady=4)

        # -------------------------
        # Forma de pago
        # -------------------------
        frame_pago = ttk.LabelFrame(self, text="Forma de pago", padding=10)
        frame_pago.pack(fill="x", padx=10, pady=10)

        self.forma_pago = tk.StringVar()
        self.forma_pago.set("Efectivo")

        ttk.Radiobutton(frame_pago, text="Efectivo", value="Efectivo",          variable=self.forma_pago).pack(side="left", padx=5)
        ttk.Radiobutton(frame_pago, text="Tarjeta", value="Tarjeta", variable=self.forma_pago).pack(side="left", padx=5)
        ttk.Radiobutton(frame_pago, text="Transferencia", value="Transferencia", variable=self.forma_pago).pack(side="left", padx=5)

# -------------------------
# Botón guardar Excel por forma de pago
# -------------------------
        ttk.Button(self, text="Guardar Excel por Pago", command=self.guardar_excel_pago).pack(pady=10)


        # ------------------
        # Tabla de venta
        # ------------------
        cols = ["codigo", "descripcion", "precio", "cantidad", "total"]
        self.tree = ttk.Treeview(self, columns=cols, show='headings', height=14)
        for c in cols:
            self.tree.heading(c, text=c.capitalize())
            self.tree.column(c, width=120, anchor='center')
        self.tree.pack(fill='both', expand=True, padx=6, pady=6)

    # ------------------
    # Funciones internas
    # ------------------
    def completar_datos(self, event=None):
        codigo = self.cod_entry.get().strip().upper()
        if not codigo:
            return
        df = load_inventario_file()
        mask = df['codigo'].astype(str).str.upper() == codigo
        if mask.any():
            fila = df[mask].iloc[0]
            self.desc_entry.config(state='normal')
            self.precio_entry.config(state='normal')
            self.desc_entry.delete(0, tk.END)
            self.desc_entry.insert(0, fila['descripcion'])
            self.precio_entry.delete(0, tk.END)
            self.precio_entry.insert(0, fila['precio'])
            self.desc_entry.config(state='readonly')
            self.precio_entry.config(state='readonly')
        else:
            self.desc_entry.config(state='normal')
            self.precio_entry.config(state='normal')
            self.desc_entry.delete(0, tk.END)
            self.precio_entry.delete(0, tk.END)
            self.desc_entry.config(state='readonly')
            self.precio_entry.config(state='readonly')

    def agregar_a_venta(self):
        codigo = self.cod_entry.get().strip().upper()
        try:
            cantidad = int(self.cant_entry.get())
        except:
            cantidad = 0
        try:
            precio = float(self.precio_entry.get())
        except:
            precio = 0.0
        descripcion = self.desc_entry.get()
        if not codigo or cantidad <= 0 or precio <= 0:
            messagebox.showwarning("Atención", "Código, cantidad y precio deben ser válidos")
            return
        total = cantidad * precio
        self.tree.insert('', 'end', values=(codigo, descripcion, precio, cantidad, total))
        # Limpiar campos
        self.cod_entry.delete(0, tk.END)
        self.desc_entry.config(state='normal'); self.desc_entry.delete(0, tk.END); self.desc_entry.config(state='readonly')
        self.precio_entry.config(state='normal'); self.precio_entry.delete(0, tk.END); self.precio_entry.config(state='readonly')
        self.cant_entry.delete(0, tk.END)

    def guardar_excel_pago(self):
        try:
            productos = [self.tree.item(i)["values"] for i in self.tree.get_children()]
            if not productos:
                messagebox.showwarning("Atención", "No hay productos en la venta.")
                return
        
            df = pd.DataFrame(productos, columns=["Código", "Descripción", "Precio", "Cantidad", "Total"])
            forma = self.forma_pago.get()

            archivo = f"venta_{forma}.xlsx"
            df.to_excel(archivo, index=False)

            messagebox.showinfo("Éxito", f"Venta guardada como:\n{archivo}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar en Excel:\n{e}")


# ===============================
# COTIZACION
# ===============================
# ===============================
# COTIZACION
# ===============================
class Cotizacion(ttk.Frame):
    def __init__(self, parent, controller=None, inventario_df=None):
        super().__init__(parent)
        self.controller = controller
        self.inventario_df = inventario_df

        # ------------------------
        # Variables
        # ------------------------
        self.total_parcial_var = tk.StringVar(value="0.00")
        self.total_general_var = tk.StringVar(value="0.00")

        # ------------------------
        # Título
        # ------------------------
        ttk.Label(self, text='COTIZACIÓN', font=('Segoe UI', 12, 'bold')).pack(anchor='w', padx=6, pady=6)

        # ------------------------
        # Frame de búsqueda / agregar
        # ------------------------
        frame_buscar = ttk.LabelFrame(self, text="Agregar Producto a Cotización", padding=10)
        frame_buscar.pack(fill="x", padx=10, pady=10)

        # ---- Código ----
        ttk.Label(frame_buscar, text="Código:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_codigo = ttk.Entry(frame_buscar, width=25)
        self.entry_codigo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.entry_codigo.bind("<KeyRelease>", self.autocompletar_producto)

        # ---- Descripción ----
        ttk.Label(frame_buscar, text="Descripción:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_desc = ttk.Entry(frame_buscar, width=40)
        self.entry_desc.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # ---- Precio ----
        ttk.Label(frame_buscar, text="Precio:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.entry_precio = ttk.Entry(frame_buscar, width=20)
        self.entry_precio.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.entry_precio.bind("<KeyRelease>", self.actualizar_total_parcial)

        # ---- Stock ----
        ttk.Label(frame_buscar, text="Cantidad disponible:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.entry_stock = ttk.Entry(frame_buscar, width=20)
        self.entry_stock.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # ---- Cantidad ----
        ttk.Label(frame_buscar, text="Cantidad a cotizar:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.entry_cantidad = ttk.Entry(frame_buscar, width=10)
        self.entry_cantidad.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.entry_cantidad.bind("<KeyRelease>", self.actualizar_total_parcial)

        # ---- Disponibilidad ----
        ttk.Label(frame_buscar, text="Disponibilidad:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.combo_disp = ttk.Combobox(frame_buscar, values=["Disponible", "No disponible"], width=18)
        self.combo_disp.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        self.combo_disp.set("Disponible")

        # ---- Total parcial dinámico ----
        ttk.Label(frame_buscar, text="TOTAL: ", font=("Arial", 10, "bold")).grid(row=6, column=0, pady=5, sticky="e")
        ttk.Label(frame_buscar, textvariable=self.total_parcial_var, font=("Arial", 12, "bold"), foreground="green").grid(row=6, column=1, pady=5, sticky="w")

        # ---- Botón agregar ----
        ttk.Button(frame_buscar, text="Agregar a Cotización", command=self.agregar_producto).grid(
            row=7, column=0, columnspan=2, pady=10
        )

        # ------------------------
        # Treeview
        # ------------------------
        columnas = ("Código", "Descripción", "Precio", "Cantidad", "Total", "Disponibilidad")
        self.tree = ttk.Treeview(self, columns=columnas, show="headings", height=12)
        for col in columnas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        habilitar_copia_treeview(self.tree)

        # ------------------------
        # Total general
        # ------------------------
        frame_tg = ttk.Frame(self)
        frame_tg.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame_tg, text="Total general:", font=("Arial", 10, "bold")).pack(side="left")
        ttk.Label(frame_tg, textvariable=self.total_general_var, font=("Arial", 10, "bold")).pack(side="left", padx=8)

        # ------------------------
        # Botones finales
        # ------------------------
        frame_acciones_final = ttk.Frame(self)
        frame_acciones_final.pack(fill="x", padx=10, pady=10)

        ttk.Button(frame_acciones_final, text="Eliminar Seleccionado", command=self.eliminar_producto).pack(side="left", padx=5)
        ttk.Button(frame_acciones_final, text="Guardar en Excel", command=self.guardar_excel).pack(side="left", padx=5)
        ttk.Button(frame_acciones_final, text="Crear Ticket PDF", command=self.crear_ticket_pdf).pack(side="left", padx=5)

    # =====================================================
    # AUTOCOMPLETAR PRODUCTO
    # =====================================================
    def autocompletar_producto(self, event=None):
        codigo = quitar_acentos(self.entry_codigo.get().strip()).upper()
        if len(codigo) < 1:
            return

        try:
            df = load_inventario_file()
            df.columns = df.columns.str.strip().str.lower()
            df.fillna("", inplace=True)

            if "codigo" not in df.columns:
                df["codigo"] = df.iloc[:, 0]

            df["codigo_clean"] = df["codigo"].apply(lambda x: quitar_acentos(str(x)).upper())
            resultado = df[df["codigo_clean"] == codigo]

            if resultado.empty:
                resultado = df[df["codigo_clean"].str.contains(codigo, na=False)]
                if resultado.empty:
                    for e in [self.entry_desc, self.entry_precio, self.entry_stock]:
                        e.delete(0, tk.END)
                    self.total_parcial_var.set("0.00")
                    return

            fila = resultado.iloc[0]

            self.entry_desc.delete(0, tk.END)
            self.entry_desc.insert(0, fila.get("descripcion", ""))

            self.entry_precio.delete(0, tk.END)
            self.entry_precio.insert(0, str(fila.get("precio", "")))

            self.entry_stock.delete(0, tk.END)
            self.entry_stock.insert(0, str(fila.get("stock", "")))

            self.actualizar_total_parcial()

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar inventario: {e}")

    # =====================================================
    # TOTAL PARCIAL AUTOMÁTICO
    # =====================================================
    def actualizar_total_parcial(self, event=None):
        try:
            precio = float(self.entry_precio.get())
        except:
            precio = 0
        try:
            cantidad = float(self.entry_cantidad.get())
        except:
            cantidad = 0

        total = precio * cantidad
        self.total_parcial_var.set(f"{total:,.2f}")

    # =====================================================
    # AGREGAR PRODUCTO A COTIZACIÓN
    # =====================================================
    def agregar_producto(self):
        try:
            codigo = self.entry_codigo.get().strip()
            desc = self.entry_desc.get().strip()
            precio = float(self.entry_precio.get())
            cantidad = int(self.entry_cantidad.get())
            total = precio * cantidad
            disp = self.combo_disp.get().strip()

            self.tree.insert("", tk.END, values=(codigo, desc, f"{precio:.2f}", cantidad, f"{total:.2f}", disp))

            for e in [self.entry_codigo, self.entry_desc, self.entry_precio, self.entry_stock, self.entry_cantidad]:
                e.delete(0, tk.END)

            self.combo_disp.set("Disponible")
            self.total_parcial_var.set("0.00")

            self.recalcular_total_general()

        except Exception as e:
            messagebox.showwarning("Atención", f"Error al agregar producto: {e}")

    # =====================================================
    # ELIMINAR PRODUCTO
    # =====================================================
    def eliminar_producto(self):
        sel = self.tree.selection()
        for item in sel:
            self.tree.delete(item)
        self.recalcular_total_general()

    # =====================================================
    # TOTAL GENERAL
    # =====================================================
    def recalcular_total_general(self):
        total = 0.0
        for item in self.tree.get_children():
            try:
                total += float(self.tree.item(item)["values"][4])
            except:
                pass
        self.total_general_var.set(f"{total:,.2f}")

    # =====================================================
    # GUARDAR EXCEL
    # =====================================================
    def guardar_excel(self):
        try:
            productos = [self.tree.item(i)["values"] for i in self.tree.get_children()]
            if not productos:
                messagebox.showwarning("Atención", "No hay productos para guardar.")
                return

            df = pd.DataFrame(productos, columns=["Código", "Descripción", "Precio", "Cantidad", "Total", "Disponibilidad"])
            archivo = "cotizacion.xlsx"
            df.to_excel(archivo, index=False)

            messagebox.showinfo("Éxito", f"Cotización guardada en {archivo}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar en Excel: {e}")

    # =====================================================
    # CREAR TICKET PDF
    # =====================================================
    def crear_ticket_pdf(self):
        try:
            productos = [self.tree.item(i)["values"] for i in self.tree.get_children()]
            if not productos:
                messagebox.showwarning("Atención", "No hay productos para exportar.")
                return

            from reportlab.pdfgen import canvas

            archivo = "ticket_cotizacion.pdf"
            c = canvas.Canvas(archivo)
            y = 800

            c.setFont("Helvetica-Bold", 14)
            c.drawString(50, y, "COTIZACIÓN")
            y -= 30

            c.setFont("Helvetica", 10)
            for p in productos:
                linea = f"{p[0]} | {p[1]} | {p[3]} x ${p[2]} = ${p[4]}"
                c.drawString(40, y, linea)
                y -= 20

            c.setFont("Helvetica-Bold", 12)
            c.drawString(40, y-10, f"TOTAL GENERAL: ${self.total_general_var.get()}")

            c.save()
            messagebox.showinfo("Éxito", f"PDF creado: {archivo}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear el PDF: {e}")



class Taller(ttk.Frame):
    def __init__(self, parent, controller=None, inventario_df=None):
        super().__init__(parent)
        self.controller = controller
        self.inventario_df = inventario_df

        ttk.Label(self, text='REGISTRO DEL TALLER', font=('Segoe UI', 12, 'bold')).pack(anchor='w', padx=6, pady=6)

        # ------------------------
        # Frame agregar insumo
        # ------------------------
        frame_add = ttk.LabelFrame(self, text="Agregar insumo utilizado", padding=10)
        frame_add.pack(fill="x", padx=10, pady=10)

        # Código
        ttk.Label(frame_add, text="Código:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        codigo_var = tk.StringVar()
        entry_codigo = ttk.Entry(frame_add, width=20, textvariable=codigo_var)
        entry_codigo.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Descripción
        ttk.Label(frame_add, text="Descripción:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        descripcion_var = tk.StringVar()
        entry_desc = ttk.Entry(frame_add, width=40, textvariable=descripcion_var)
        entry_desc.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Precio
        ttk.Label(frame_add, text="Precio:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        precio_var = tk.StringVar()
        entry_precio = ttk.Entry(frame_add, width=15, textvariable=precio_var)
        entry_precio.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # Cantidad
        ttk.Label(frame_add, text="Cantidad:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        cantidad_var = tk.StringVar(value="1")
        entry_cant = ttk.Entry(frame_add, width=10, textvariable=cantidad_var)
        entry_cant.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Total
        ttk.Label(frame_add, text="Total:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        total_var = tk.StringVar(value="0.00")
        entry_total = ttk.Entry(frame_add, width=15, textvariable=total_var, state="readonly")
        entry_total.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        # ------------------------
        # FUNCIONES INTERNAS
        # ------------------------
        def actualizar_total(*args):
            try:
                p = float(precio_var.get())
                c = float(cantidad_var.get())
                total_var.set(f"{p * c:.2f}")
            except:
                total_var.set("0.00")

        cantidad_var.trace("w", actualizar_total)
        precio_var.trace("w", actualizar_total)

        # ------------------------
        # AUTOCOMPLETAR POR CÓDIGO
        # ------------------------
        codigo_var.trace("w", lambda *args: (
            self.cargar_descripcion_precio(codigo_var, descripcion_var, precio_var),
            actualizar_total()
        ))

        # ------------------------
        # BOTÓN AGREGAR
        # ------------------------
        ttk.Button(frame_add, text="Agregar insumo", command=lambda: self.agregar_insumo(
            codigo_var, descripcion_var, precio_var, cantidad_var, total_var
        )).grid(row=5, column=0, columnspan=2, pady=10)

        # ------------------------
        # TREEVIEW
        # ------------------------
        cols = ("Código", "Descripción", "Precio", "Cantidad", "Total")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, anchor="center")

        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        # ------------------------
        # TOTAL GENERAL
        # ------------------------
        self.total_taller_var = tk.StringVar(value="0.00")
        frame_total = ttk.Frame(self)
        frame_total.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_total, text="Total del Taller:", font=("Arial", 10, "bold")).pack(side="left")
        ttk.Label(frame_total, textvariable=self.total_taller_var, font=("Arial", 10, "bold")).pack(side="left", padx=10)

        # ------------------------
        # ACCIONES
        # ------------------------
        frame_btns = ttk.Frame(self)
        frame_btns.pack(fill="x", padx=10, pady=10)

        ttk.Button(frame_btns, text="Eliminar seleccionado",
                   command=self.eliminar_insumo).pack(side="left", padx=5)

        ttk.Button(frame_btns, text="Guardar Taller en Excel",
                   command=self.guardar_excel).pack(side="left", padx=5)

        ttk.Button(frame_btns, text="Crear Ticket PDF",
                   command=self.crear_ticket_pdf).pack(side="left", padx=5)

    # =====================================================================
    # ----------- FUNCIONES PRINCIPALES (IGUAL QUE VENTAS) ---------------
    # =====================================================================

    def cargar_descripcion_precio(self, codigo_var, desc_var, precio_var):
        try:
            codigo = codigo_var.get().strip().upper()
            if not codigo or self.inventario_df is None or self.inventario_df.empty:
                return

            df = self.inventario_df.copy()
            df.columns = df.columns.str.lower().str.strip()
            df["codigo"] = df["codigo"].astype(str).str.upper()

            fila = df[df["codigo"] == codigo]

            if not fila.empty:
                fila = fila.iloc[0]
                desc_var.set(fila.get("descripcion", ""))
                precio_var.set(str(fila.get("precio", "")))
            else:
                desc_var.set("")
                precio_var.set("")

        except Exception as e:
            print("Error autocompletar Taller:", e)

    def agregar_insumo(self, cod, desc, precio, cant, total):
        try:
            datos = (
                cod.get(),
                desc.get(),
                f"{float(precio.get()):.2f}",
                cant.get(),
                f"{float(total.get()):.2f}"
            )

            self.tree.insert("", tk.END, values=datos)

            cod.set("")
            desc.set("")
            precio.set("")
            cant.set("1")
            total.set("0.00")

            self.recalcular_total_taller()

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo agregar insumo:\n{e}")

    def eliminar_insumo(self):
        for sel in self.tree.selection():
            self.tree.delete(sel)
        self.recalcular_total_taller()

    def recalcular_total_taller(self):
        total = 0
        for item in self.tree.get_children():
            total += float(self.tree.item(item)["values"][4])
        self.total_taller_var.set(f"{total:.2f}")

    def guardar_excel(self):
        try:
            insumos = [self.tree.item(i)["values"] for i in self.tree.get_children()]
            if not insumos:
                messagebox.showwarning("Atención", "No hay insumos para guardar.")
                return

            df = pd.DataFrame(insumos, columns=["Código", "Descripción", "Precio", "Cantidad", "Total"])
            df.to_excel("taller.xlsx", index=False)
            messagebox.showinfo("Éxito", "Guardado como taller.xlsx")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{e}")

    def crear_ticket_pdf(self):
        try:
            items = [self.tree.item(i)["values"] for i in self.tree.get_children()]
            if not items:
                messagebox.showwarning("Atención", "No hay insumos para el ticket.")
                return

            archivo = "taller_ticket.pdf"
            c = canvas.Canvas(archivo, pagesize=letter)
            y = 750

            c.setFont("Helvetica-Bold", 14)
            c.drawString(40, y, "TICKET DEL TALLER")
            y -= 30
            total_general = 0

            c.setFont("Helvetica", 10)
            for codigo, desc, precio, cant, total in items:
                total_general += float(total)

                c.drawString(40, y, f"{codigo} - {desc}")
                y -= 15
                c.drawString(60, y, f"Precio: {precio}  Cant: {cant}  Total: {total}")
                y -= 25

                if y < 60:
                    c.showPage()
                    y = 750

            c.setFont("Helvetica-Bold", 12)
            c.drawString(40, y, f"Total General: ${total_general:.2f}")
            c.save()

            messagebox.showinfo("Éxito", f"Ticket creado:\n{archivo}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar PDF:\n{e}")



# --------------------
# MAIN APP
# --------------------
class AppUnificada(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("JQ MOTORS SISTEM")
        self.geometry("1300x800")
        self.configure(bg="white")

        # DataFrame de inventario (inicializar vacío o cargar desde Excel/CSV)
        self.inventario_df = pd.DataFrame()  # reemplaza con tu carga real si tienes archivo

        # Definir la variable del total
        self.total_var = tk.StringVar()
        self.total_var.set("0.00")

        # HEADER
        header = ttk.Frame(self)
        header.pack(fill='x', pady=(5,0))
        ttk.Label(header, text='🚀 JQ MOTORS OJO DE AGUA', font=('Segoe UI',14,'bold')).pack(side='left', padx=10)

        # Notebook principal
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True)

        # Crear instancias de cada pestaña
        self.stock_tab = Stock(self.notebook)
        self.ventas_tab = Ventas(self.notebook)
        self.cot_tab = Cotizacion(self.notebook, controller=self, inventario_df=self.inventario_df)
        self.taller_tab = Taller(self.notebook)

        # Añadir pestañas al notebook
        self.notebook.add(self.stock_tab, text='Stock')
        self.notebook.add(self.ventas_tab, text='Ventas')
        self.notebook.add(self.cot_tab, text='Cotizacion')
        self.notebook.add(self.taller_tab, text='Taller')

        # Frame principal para entradas
        self.frame = tk.Frame(self, bg="white")
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Cambiar estilo de Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="white", foreground="#003366", fieldbackground="white", rowheight=25)
        style.configure("Treeview.Heading", background="#003366", foreground="white", font=("Arial", 10, "bold"))

        # Lista interna de registros
        self.registros = []

    # Funciones de agregar, borrar y limpiar entradas
    def agregar_registro(self):
        codigo = self.codigo_entry.get()
        descripcion = self.descripcion_entry.get()
        ubicacion = self.ubicacion_entry.get()
        stock = self.stock_entry.get()
        precio = self.precio_entry.get()

        if codigo and descripcion and ubicacion and stock and precio:
            self.tree.insert("", "end", values=(codigo, descripcion, ubicacion, stock, precio))
            self.registros.append((codigo, descripcion, ubicacion, stock, precio))
            self.limpiar_entradas()
        else:
            messagebox.showwarning("Atención", "Todos los campos son obligatorios")

    def borrar_registro(self):
        selected_item = self.tree.selection()
        if selected_item:
            self.tree.delete(selected_item)
        else:
            messagebox.showwarning("Atención", "Seleccione un registro para borrar")

    def limpiar_entradas(self):
        self.codigo_entry.delete(0, tk.END)
        self.descripcion_entry.delete(0, tk.END)
        self.ubicacion_entry.delete(0, tk.END)
        self.stock_entry.delete(0, tk.END)
        self.precio_entry.delete(0, tk.END)
        # Puedes aplicar aquí tu función de estilos si la tienes:
        # aplicar_estilos(self)

import json
import requests
import time
import threading
import sys
import os

# -------- CONFIGURACIÓN --------
NOMBRE_AGENCIA = "OJO DE AGUA"   # <<< CAMBIAR EN cada archivo
URL_JEFE = "http://localhost:5002/api/inventario_agencia"  # Puerto 5002


# -------- FUNCIÓN PARA OBTENER INVENTARIO REAL --------
def obtener_inventario():
    """
    Aquí se pone el código que obtiene el inventario real.
    Debe devolver una lista de diccionarios, por ejemplo:
    [{"codigo": "A001", "descripcion": "Producto 1", "stock": 10}, ...]
    """
    inventario = [
        {"codigo": "A001", "descripcion": "Producto 1", "stock": 10},
        {"codigo": "A002", "descripcion": "Producto 2", "stock": 5},
    ]
    return inventario


# -------- ENVÍO DEL INVENTARIO --------
def enviar_inventario():
    try:
        inventario = obtener_inventario()

        # Agregar nombre de agencia a cada ítem
        for item in inventario:
            item["agencia"] = NOMBRE_AGENCIA

        payload = json.dumps(inventario)

        # Enviar al servidor jefe en puerto 5002
        requests.post(
            URL_JEFE,
            data=payload,
            headers={"Content-Type": "application/json"},
            timeout=10
        )

    except:
        pass  # No mostrar errores al usuario

    # Volver a enviar en 60 segundos
    threading.Timer(60, enviar_inventario).start()


# -------- OCULTAR CONSOLA (PARA EXE) --------
def ocultar_consola():
    if os.name == "nt":  # Windows
        try:
            import ctypes
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd != 0:
                ctypes.windll.user32.ShowWindow(hwnd, 0)
        except:
            pass


# -------- INICIO DEL PROGRAMA --------
if __name__ == "__main__":
    ocultar_consola()
    enviar_inventario()

    # Mantener el programa corriendo
    while True:
        time.sleep(1)


# --------------------
# MAIN
# --------------------
if __name__=="__main__":
    app = AppUnificada()
    app.mainloop()
