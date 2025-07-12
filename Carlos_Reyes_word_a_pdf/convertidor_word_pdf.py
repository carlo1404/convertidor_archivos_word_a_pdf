import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx2pdf import convert
import os
import threading
# elaborado por Carlos Reyes
# fecha: 11/07/2025
# version: 1.0
# descripcion: este programa convierte un archivo word a pdf
# uso: python convertidor_word_pdf.py
# requisitos: python 3.10
# requisitos: docx2pdf
# Colores y fuentes para el diseño
COLOR_FONDO = "#f0f4f8"
COLOR_BOTON = "#1976d2"
COLOR_BOTON_TEXTO = "#ffffff"
COLOR_LABEL = "#2B2B2B"
FUENTE_TITULO = ("Arial", 16, "bold")
FUENTE_NORMAL = ("Arial", 11)
FUENTE_BOTON = ("Arial", 11, "bold")

# hACEMOS UNA CLASE PARA LA INTERFAZ GRAFIFCA Y UN CONSTRUCTOR PARA INICIALIZAR LA VENTANA  CON LOS ATRIBUTOS ROOT 
class ConvertidorWordPDF:
    def __init__(self, root):

        # Inicializa la ventana principal y los widgets de la interfaz gráfica.
        # :param root: Ventana principal de Tkinter
        self.root = root
        self.root.title("Convertidor Word a PDF")
        self.root.geometry("1080x720")
        self.root.configure(bg=COLOR_FONDO)
        self.archivo_word = None  # Ruta del archivo seleccionado
        self.ventana_carga = None
        self.progressbar = None

        # Marco para centrar el contenido
        self.frame = tk.Frame(root, bg=COLOR_FONDO)
        self.frame.pack(expand=True)

        # Etiqueta de título
        self.label_titulo = tk.Label(self.frame, text="Convertidor Word a PDF", font=FUENTE_TITULO, fg=COLOR_LABEL, bg=COLOR_FONDO)
        self.label_titulo.pack(pady=(10, 5))

        # Etiqueta de instrucciones
        self.label = tk.Label(self.frame, text="Selecciona un archivo .docx para convertir a PDF:", font=FUENTE_NORMAL, fg=COLOR_LABEL, bg=COLOR_FONDO)
        self.label.pack(pady=5)

        # Botón para seleccionar archivo
        self.boton_seleccionar = tk.Button(self.frame, text="Seleccionar archivo", font=FUENTE_BOTON, bg=COLOR_BOTON, fg=COLOR_BOTON_TEXTO, activebackground="#1565c0", activeforeground=COLOR_BOTON_TEXTO, command=self.seleccionar_archivo, cursor="hand2")
        self.boton_seleccionar.pack(pady=5, ipadx=10, ipady=2)

        # Etiqueta para mostrar el archivo seleccionado
        self.label_archivo = tk.Label(self.frame, text="Ningún archivo seleccionado", font=FUENTE_NORMAL, fg="#3B3A3A", bg=COLOR_FONDO)
        self.label_archivo.pack(pady=5)

        # Botón para convertir (deshabilitado hasta que se seleccione un archivo)
        self.boton_convertir = tk.Button(self.frame, text="Convertir a PDF", font=FUENTE_BOTON, bg="#43a047", fg=COLOR_BOTON_TEXTO, activebackground="#388e3c", activeforeground=COLOR_BOTON_TEXTO, command=self.convertir_a_pdf, state=tk.DISABLED, cursor="hand2")
        self.boton_convertir.pack(pady=12, ipadx=10, ipady=2)

    def seleccionar_archivo(self):

        # Abre un cuadro de diálogo para seleccionar un archivo .docx.
        # Habilita el botón de conversión si se selecciona un archivo válido.
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Word",
            filetypes=[("Archivos Word", "*.docx")]
        )
        if archivo:
            self.archivo_word = archivo
            self.label_archivo.config(text=os.path.basename(archivo), fg=COLOR_LABEL)
            self.boton_convertir.config(state=tk.NORMAL)
        else:
            self.archivo_word = None
            self.label_archivo.config(text="Ningún archivo seleccionado", fg="#606060")
            self.boton_convertir.config(state=tk.DISABLED)

    def mostrar_animacion_carga(self):
        """
        Muestra una ventana emergente con una barra de progreso indeterminada.
        """
        self.ventana_carga = tk.Toplevel(self.root)
        self.ventana_carga.title("Convirtiendo...")
        self.ventana_carga.geometry("300x100")
        self.ventana_carga.resizable(False, False)
        self.ventana_carga.configure(bg=COLOR_FONDO)
        self.ventana_carga.grab_set()
        label = tk.Label(self.ventana_carga, text="Convirtiendo a PDF...", font=FUENTE_NORMAL, fg=COLOR_LABEL, bg=COLOR_FONDO)
        label.pack(pady=(20, 10))
        self.progressbar = ttk.Progressbar(self.ventana_carga, mode='indeterminate', length=220)
        self.progressbar.pack(pady=5)
        self.progressbar.start(10)
        # Centrar la ventana de carga
        self.ventana_carga.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (300 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (100 // 2)
        self.ventana_carga.geometry(f"300x100+{x}+{y}")

    def cerrar_animacion_carga(self):
        # Cerramos la animacion de carga con el stop y el destroy 
        if self.progressbar:
            self.progressbar.stop()
        if self.ventana_carga:
            self.ventana_carga.destroy()
            self.ventana_carga = None

    def convertir_a_pdf(self):
        # Inicia la conversión en un hilo para no congelar la interfaz y muestra la animación de carga.
        
        if not self.archivo_word:
            messagebox.showerror("Error", "No se ha seleccionado ningún archivo.")
            return
        carpeta_salida = filedialog.askdirectory(title="Selecciona la carpeta de destino para el PDF")
        if not carpeta_salida:
            return  # El usuario canceló la selección
        # Mostrar animación de carga y lanzar hilo
        self.mostrar_animacion_carga()
        hilo = threading.Thread(target=self._convertir_en_hilo, args=(self.archivo_word, carpeta_salida))
        hilo.start()

    def _convertir_en_hilo(self, archivo_word, carpeta_salida):
        # Realiza la conversión en un hilo separado y gestiona los mensajes y la animación.

        try:
            convert(archivo_word, carpeta_salida)
            nombre_pdf = os.path.splitext(os.path.basename(archivo_word))[0] + ".pdf"
            ruta_pdf = os.path.join(carpeta_salida, nombre_pdf)
            self.root.after(0, self.cerrar_animacion_carga)
            if os.path.exists(ruta_pdf):
                self.root.after(0, lambda: messagebox.showinfo("Éxito", f"Archivo convertido exitosamente:\n{ruta_pdf}"))
            else:
                self.root.after(0, lambda: messagebox.showwarning("Advertencia", "La conversión terminó, pero no se encontró el PDF."))
        except Exception as e:
            self.root.after(0, self.cerrar_animacion_carga)
            self.root.after(0, lambda: messagebox.showerror("Error", f"Ocurrió un error al convertir el archivo:\n{e}"))

def main():
    """
    Función principal que inicia la aplicación.
    """
    root = tk.Tk()
    app = ConvertidorWordPDF(root)
    root.mainloop()

if __name__ == "__main__":
    main() 