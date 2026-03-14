import customtkinter as ctk
from tkinter import filedialog
import fitz  

ctk.set_appearance_mode("Dark")  
ctk.set_default_color_theme("blue")  

class LimpiadorEditorialApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Limpiador Editorial Pro - UNAM")
        self.geometry("900x700")
        
        # Aquí guardaremos la información de cada bloque para armar el HTML al final
        self.datos_bloques = []

        self.titulo_label = ctk.CTkLabel(
            self, 
            text="Extractor de PDF a HTML Semántico", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.titulo_label.pack(pady=10)

        # Marco superior para poner los botones juntos
        self.frame_botones = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_botones.pack(pady=10)

        self.boton_cargar = ctk.CTkButton(
            self.frame_botones, 
            text="Cargar PDF", 
            command=self.evento_cargar_archivo
        )
        self.boton_cargar.pack(side="left", padx=10)

        # ¡NUEVO BOTÓN VERDE PARA EXPORTAR!
        self.boton_exportar = ctk.CTkButton(
            self.frame_botones, 
            text="Generar HTML Limpio", 
            command=self.evento_exportar_html,
            fg_color="#28a745", hover_color="#218838"
        )
        self.boton_exportar.pack(side="left", padx=10)

        self.frame_scroll = ctk.CTkScrollableFrame(self, width=800, height=500)
        self.frame_scroll.pack(pady=10)

    def evento_cargar_archivo(self):
        ruta_archivo = filedialog.askopenfilename(
            title="Selecciona el PDF de la revista",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        
        if ruta_archivo:
            # Limpiamos todo al cargar un archivo nuevo
            for widget in self.frame_scroll.winfo_children():
                widget.destroy()
            self.datos_bloques.clear() # Vaciamos la memoria
                
            try:
                pdf_doc = fitz.open(ruta_archivo)
                bloques_utiles = []
                
                for num_pagina in range(min(3, len(pdf_doc))):
                    pagina = pdf_doc.load_page(num_pagina)
                    bloques = pagina.get_text("blocks")
                    
                    for b in bloques:
                        tipo = b[6] 
                        if tipo == 0: 
                            texto = b[4].replace('\n', ' ').strip()
                            if len(texto) > 10: 
                                bloques_utiles.append(("texto", texto))
                        elif tipo == 1:
                            bloques_utiles.append(("imagen", "[IMAGEN DETECTADA EN EL PDF]"))
                
                for item in bloques_utiles: 
                    self.crear_bloque_ui(item)
                    
            except Exception as e:
                print(f"Error al procesar: {e}")

    def crear_bloque_ui(self, item):
        tipo, contenido = item
        
        if tipo == "texto":
            texto_mostrar = contenido[:150] + "..." if len(contenido) > 150 else contenido
            color_fondo = "#2b2b2b"
        else:
            texto_mostrar = f"🖼️ {contenido}"
            color_fondo = "#1f538d"

        bloque_frame = ctk.CTkFrame(self.frame_scroll, fg_color=color_fondo)
        bloque_frame.pack(fill="x", padx=10, pady=5)
                
        label_texto = ctk.CTkLabel(bloque_frame, text=texto_mostrar, wraplength=550, justify="left")
        label_texto.pack(side="left", padx=10, pady=10)
        
        opciones = ["Normal", "Pie de Figura", "Título", "Ignorar"]
        menu_clasificacion = ctk.CTkOptionMenu(bloque_frame, values=opciones)
        
        if tipo == "imagen":
            menu_clasificacion.set("Imagen")
            menu_clasificacion.configure(state="disabled")
            
        menu_clasificacion.pack(side="right", padx=10, pady=10)

        # ¡CLAVE! Guardamos la referencia del menú y el texto en nuestra memoria
        self.datos_bloques.append({
            "contenido": contenido,
            "menu": menu_clasificacion
        })

    # ¡LA FUNCIÓN MÁGICA QUE CREA EL HTML!
    def evento_exportar_html(self):
        if not self.datos_bloques:
            print("Carga un PDF primero antes de exportar.")
            return

        # Te pregunta dónde quieres guardar el nuevo archivo HTML
        ruta_guardar = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("Archivo HTML", "*.html")],
            title="Guardar HTML Semántico"
        )
        
        if not ruta_guardar: # Si cancelas la ventana, no hace nada
            return

        # Empezamos a armar el esqueleto del HTML
        html_final = "<!DOCTYPE html>\n<html lang='es'>\n<head>\n<meta charset='UTF-8'>\n<title>Artículo UNAM</title>\n</head>\n<body>\n"

        # Leemos qué elegiste en cada menú
        for bloque in self.datos_bloques:
            eleccion = bloque["menu"].get()
            texto = bloque["contenido"]

            if eleccion == "Ignorar":
                continue
            elif eleccion == "Título":
                html_final += f"  <h2>{texto}</h2>\n"
            elif eleccion == "Normal":
                html_final += f"  <p>{texto}</p>\n"
            elif eleccion == "Pie de Figura":
                html_final += f"  <figcaption>{texto}</figcaption>\n"
            elif eleccion == "Imagen":
                html_final += f"  <figure>\n    <img src='ruta_imagen_pendiente.jpg' alt='Imagen'>\n  </figure>\n"

        html_final += "</body>\n</html>"

        # Guardamos el archivo en tu compu
        with open(ruta_guardar, "w", encoding="utf-8") as archivo:
            archivo.write(html_final)
            
        print("¡BOOM! HTML guardado con éxito.")

if __name__ == "__main__":
    app = LimpiadorEditorialApp()
    app.mainloop()