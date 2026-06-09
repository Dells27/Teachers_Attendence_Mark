"""
Sistema de Marcas con Reconocimiento Facial
Escuela - Tkinter Desktop App
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import cv2
import face_recognition
import numpy as np
import pandas as pd
import os
import json
import pickle
from datetime import datetime
from pathlib import Path

# ──────────────────────────────────────────
# CONFIGURACIÓN DE CARPETAS
# ──────────────────────────────────────────
BASE_DIR     = Path(__file__).parent
FOTOS_DIR    = BASE_DIR / "profesores_fotos"
ENCODINGS_FILE = BASE_DIR / "encodings.pkl"
MARCAS_FILE  = BASE_DIR / "marcas.xlsx"

FOTOS_DIR.mkdir(exist_ok=True)


# ──────────────────────────────────────────
# UTILIDADES DE DATOS
# ──────────────────────────────────────────
def cargar_encodings():
    """Carga los encodings faciales guardados."""
    if ENCODINGS_FILE.exists():
        with open(ENCODINGS_FILE, "rb") as f:
            return pickle.load(f)
    return {}  # {nombre: [encoding1, encoding2, ...]}


def guardar_encodings(encodings):
    with open(ENCODINGS_FILE, "wb") as f:
        pickle.dump(encodings, f)


def registrar_marca(nombre):
    """Agrega una fila al Excel con entrada o salida."""
    hoy = datetime.now().strftime("%Y-%m-%d")
    hora = datetime.now().strftime("%H:%M:%S")

    # Cargar marcas existentes
    if MARCAS_FILE.exists():
        df = pd.read_excel(MARCAS_FILE)
    else:
        df = pd.DataFrame(columns=["Nombre", "Fecha", "Hora", "Tipo"])

    # Determinar si es entrada o salida
    marcas_hoy = df[(df["Nombre"] == nombre) & (df["Fecha"] == hoy)]
    if len(marcas_hoy) % 2 == 0:
        tipo = "Entrada"
    else:
        tipo = "Salida"

    # Agregar fila
    nueva = pd.DataFrame([{"Nombre": nombre, "Fecha": hoy, "Hora": hora, "Tipo": tipo}])
    df = pd.concat([df, nueva], ignore_index=True)

    try:
        df.to_excel(MARCAS_FILE, index=False)
    except PermissionError:
        raise PermissionError(
            "No se pudo guardar la marca porque el archivo 'marcas.xlsx' "
            "está abierto en Excel.\n\nCiérralo e intenta de nuevo."
        )

    return tipo, hora


# ──────────────────────────────────────────
# NOTIFICACIÓN DE MARCA
# ──────────────────────────────────────────
class NotificacionMarca(tk.Toplevel):
    """
    Popup grande que aparece al registrar una marca.
    Se cierra solo después de 4 segundos o al hacer clic.
    """
    DURACION_MS = 4000

    def __init__(self, parent, nombre, tipo, hora, foto_path=None):
        super().__init__(parent)
        self.overrideredirect(True)   # sin barra de título
        self.attributes("-topmost", True)
        self.configure(bg="#ffffff")

        es_entrada = tipo == "Entrada"
        color_bg   = "#0F6E56" if es_entrada else "#3C3489"
        icono      = "⬆  ENTRADA" if es_entrada else "⬇  SALIDA"

        # ── Barra de color superior ──
        barra = tk.Frame(self, bg=color_bg, height=8)
        barra.pack(fill="x")

        # ── Cuerpo ──
        cuerpo = tk.Frame(self, bg="#ffffff", padx=30, pady=20)
        cuerpo.pack(fill="both", expand=True)

        # Foto del profesor (si existe)
        if foto_path and Path(foto_path).exists():
            try:
                from PIL import Image, ImageTk, ImageDraw
                img = Image.open(foto_path).convert("RGB")
                img.thumbnail((72, 72))
                # Recorte circular
                mask = Image.new("L", img.size, 0)
                draw = ImageDraw.Draw(mask)
                draw.ellipse((0, 0) + img.size, fill=255)
                img.putalpha(mask)
                self._foto_tk = ImageTk.PhotoImage(img)
                tk.Label(cuerpo, image=self._foto_tk, bg="#ffffff").pack(pady=(0, 8))
            except Exception:
                pass

        # Tipo (ENTRADA / SALIDA)
        tk.Label(cuerpo, text=icono,
                 font=("Helvetica", 13, "bold"),
                 fg=color_bg, bg="#ffffff").pack()

        # Nombre del profesor
        tk.Label(cuerpo, text=nombre,
                 font=("Helvetica", 20, "bold"),
                 fg="#1a1a1a", bg="#ffffff").pack(pady=(4, 2))

        # Hora
        tk.Label(cuerpo, text=hora,
                 font=("Helvetica", 14),
                 fg="#666666", bg="#ffffff").pack()

        # Barra de progreso de cuenta regresiva
        self._barra_prog = ttk.Progressbar(
            cuerpo, length=220, maximum=self.DURACION_MS, value=self.DURACION_MS
        )
        self._barra_prog.pack(pady=(14, 2))
        tk.Label(cuerpo, text="Toca para cerrar",
                 font=("Helvetica", 9), fg="#aaaaaa", bg="#ffffff").pack()

        # ── Barra de color inferior ──
        tk.Frame(self, bg=color_bg, height=8).pack(fill="x")

        # Centrar en la pantalla
        self.update_idletasks()
        ancho  = self.winfo_reqwidth()
        alto   = self.winfo_reqheight()
        x = (self.winfo_screenwidth()  // 2) - (ancho // 2)
        y = (self.winfo_screenheight() // 2) - (alto // 2)
        self.geometry(f"+{x}+{y}")

        # Cerrar al hacer clic en cualquier parte
        self.bind("<Button-1>", lambda e: self.destroy())
        for widget in self.winfo_children():
            self._bind_clic(widget)

        # Cuenta regresiva y cierre automático
        self._tiempo_restante = self.DURACION_MS
        self._tick()

    def _bind_clic(self, widget):
        widget.bind("<Button-1>", lambda e: self.destroy())
        for hijo in widget.winfo_children():
            self._bind_clic(hijo)

    def _tick(self):
        if not self.winfo_exists():
            return
        self._tiempo_restante -= 50
        self._barra_prog["value"] = max(0, self._tiempo_restante)
        if self._tiempo_restante <= 0:
            self.destroy()
        else:
            self.after(50, self._tick)


# ──────────────────────────────────────────
# VENTANA PRINCIPAL (MENÚ)
# ──────────────────────────────────────────
class AppPrincipal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Marcas — Escuela")
        self.geometry("480x360")
        self.resizable(False, False)
        self.configure(bg="#f5f5f0")
        self._build_ui()

    def _build_ui(self):
        # Encabezado
        header = tk.Frame(self, bg="#3C3489", height=70)
        header.pack(fill="x")
        tk.Label(
            header,
            text="📋  Sistema de Marcas Docentes",
            font=("Helvetica", 15, "bold"),
            fg="white", bg="#3C3489"
        ).pack(pady=20)

        # Contenedor de botones
        frame = tk.Frame(self, bg="#f5f5f0")
        frame.pack(expand=True, fill="both", padx=40, pady=20)

        botones = [
            ("👤  Registrar Profesor",  "#3C3489", self.abrir_registro),
            ("✅  Iniciar Marcas",       "#0F6E56", self.abrir_marcas),
            ("📊  Ver Reportes",         "#854F0B", self.abrir_reportes),
        ]

        for texto, color, cmd in botones:
            btn = tk.Button(
                frame,
                text=texto,
                font=("Helvetica", 12),
                bg=color, fg="white",
                activebackground=color,
                activeforeground="white",
                relief="flat",
                cursor="hand2",
                command=cmd,
                pady=10
            )
            btn.pack(fill="x", pady=6, ipady=2)

        tk.Label(
            self,
            text="v1.0  •  Uso interno escolar",
            font=("Helvetica", 9),
            fg="#888", bg="#f5f5f0"
        ).pack(pady=(0, 8))

    def abrir_registro(self):
        if hasattr(self, "_win_registro") and self._win_registro.winfo_exists():
            self._win_registro.lift()
            return
        self._win_registro = VentanaRegistro(self)

    def abrir_marcas(self):
        if hasattr(self, "_win_marcas") and self._win_marcas.winfo_exists():
            self._win_marcas.lift()
            return
        self._win_marcas = VentanaMarcas(self)

    def abrir_reportes(self):
        if hasattr(self, "_win_reportes") and self._win_reportes.winfo_exists():
            self._win_reportes.lift()
            return
        self._win_reportes = VentanaReportes(self)


# ──────────────────────────────────────────
# VENTANA: REGISTRAR PROFESOR
# ──────────────────────────────────────────
class VentanaRegistro(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Registrar Profesor")
        self.geometry("500x420")
        self.resizable(False, False)
        self.configure(bg="#f5f5f0")
        self.fotos_capturadas = []
        self.cap = None
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="Registrar Profesor", font=("Helvetica", 14, "bold"),
                 bg="#f5f5f0").pack(pady=(16, 4))

        form = tk.Frame(self, bg="#f5f5f0")
        form.pack(padx=30, fill="x")

        tk.Label(form, text="Nombre completo:", bg="#f5f5f0",
                 font=("Helvetica", 11)).grid(row=0, column=0, sticky="w", pady=6)
        self.nombre_var = tk.StringVar()
        tk.Entry(form, textvariable=self.nombre_var, font=("Helvetica", 11),
                 width=28).grid(row=0, column=1, padx=8)

        tk.Label(form, text="Materia (opcional):", bg="#f5f5f0",
                 font=("Helvetica", 11)).grid(row=1, column=0, sticky="w", pady=6)
        self.materia_var = tk.StringVar()
        tk.Entry(form, textvariable=self.materia_var, font=("Helvetica", 11),
                 width=28).grid(row=1, column=1, padx=8)

        # Estado de fotos
        self.status_fotos = tk.Label(self, text="Fotos capturadas: 0 / 5",
                                     font=("Helvetica", 11), bg="#f5f5f0", fg="#555")
        self.status_fotos.pack(pady=(14, 2))

        self.barra = ttk.Progressbar(self, length=300, maximum=5)
        self.barra.pack(pady=4)

        # Botones
        btn_frame = tk.Frame(self, bg="#f5f5f0")
        btn_frame.pack(pady=14)

        tk.Button(btn_frame, text="📷  Capturar fotos con cámara",
                  font=("Helvetica", 11), bg="#3C3489", fg="white",
                  relief="flat", cursor="hand2", padx=10, pady=6,
                  command=self.capturar_fotos).pack(side="left", padx=6)

        tk.Button(btn_frame, text="💾  Guardar Profesor",
                  font=("Helvetica", 11), bg="#0F6E56", fg="white",
                  relief="flat", cursor="hand2", padx=10, pady=6,
                  command=self.guardar_profesor).pack(side="left", padx=6)

        self.info_label = tk.Label(self, text="", font=("Helvetica", 10),
                                   bg="#f5f5f0", fg="#3C3489", wraplength=420)
        self.info_label.pack(pady=6)

    def capturar_fotos(self):
        nombre = self.nombre_var.get().strip()
        if not nombre:
            messagebox.showwarning("Falta nombre", "Ingresa el nombre del profesor primero.")
            return

        self.info_label.config(text="Abriendo cámara... presiona ESPACIO para capturar, ESC para terminar.")
        self.update()

        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            messagebox.showerror(
                "Cámara no disponible",
                "No se pudo abrir la cámara.\n"
                "Verificá que esté conectada y no en uso por otra aplicación."
            )
            return

        fotos = []
        avisos_sin_cara = 0  # para no spamear el aviso

        while True:
            ret, frame = cap.read()
            if not ret:
                break

            display = frame.copy()
            cv2.putText(display,
                        f"ESPACIO=capturar ({len(fotos)}/5)  ESC=terminar",
                        (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (100, 220, 100), 2)

            # Detectar cara en tiempo real
            small = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
            rgb_small = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
            locs = face_recognition.face_locations(rgb_small)

            if locs:
                avisos_sin_cara = 0
                for (top, right, bottom, left) in locs:
                    cv2.rectangle(display,
                                  (left*4, top*4), (right*4, bottom*4),
                                  (0, 200, 0), 2)
                cv2.putText(display, "Cara detectada ✓", (10, 60),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 200, 0), 2)
            else:
                avisos_sin_cara += 1
                cv2.putText(display, "No se detecta cara — ajusta posicion",
                            (10, 60), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 80, 255), 2)

            cv2.imshow("Captura de fotos — " + nombre, display)
            key = cv2.waitKey(1)

            if key == 27:   # ESC
                break
            if key == 32:   # ESPACIO
                if not locs:
                    # Mostrar aviso brevemente en cámara, no interrumpir con popup
                    cv2.putText(display, "⚠ Ninguna cara visible — no se capturó",
                                (10, 90), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
                    cv2.imshow("Captura de fotos — " + nombre, display)
                    cv2.waitKey(800)
                    continue

                if len(locs) > 1:
                    cv2.putText(display, "⚠ Más de una cara — acercate solo tú",
                                (10, 90), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 140, 255), 2)
                    cv2.imshow("Captura de fotos — " + nombre, display)
                    cv2.waitKey(800)
                    continue

                fotos.append(frame.copy())
                self.fotos_capturadas = fotos
                self.status_fotos.config(text=f"Fotos capturadas: {len(fotos)} / 5")
                self.barra["value"] = len(fotos)
                self.update()
                if len(fotos) >= 5:
                    break

        cap.release()
        cv2.destroyAllWindows()

        if len(fotos) == 0:
            self.info_label.config(
                text="⚠ No se capturó ninguna foto. Intentá de nuevo con mejor iluminación."
            )
        else:
            self.info_label.config(
                text=f"✅ {len(fotos)} foto(s) capturadas. Pulsa 'Guardar Profesor' para continuar."
            )

    def guardar_profesor(self):
        nombre = self.nombre_var.get().strip()
        if not nombre:
            messagebox.showwarning("Falta nombre", "Ingresa el nombre del profesor.")
            return
        if len(self.fotos_capturadas) == 0:
            messagebox.showwarning(
                "Sin fotos",
                "Capturá al menos una foto antes de guardar.\n\n"
                "Consejos para mejores resultados:\n"
                "• Buena iluminación frontal\n"
                "• Mirá directo a la cámara\n"
                "• Evitá gorras o lentes de sol"
            )
            return

        # Guardar fotos en carpeta
        carpeta = FOTOS_DIR / nombre
        carpeta.mkdir(exist_ok=True)
        for i, foto in enumerate(self.fotos_capturadas):
            cv2.imwrite(str(carpeta / f"{i+1}.jpg"), foto)

        # Generar encodings con reporte de cuántas fallaron
        encodings = cargar_encodings()
        lista_enc = []
        fotos_sin_cara = 0

        for foto in self.fotos_capturadas:
            rgb = cv2.cvtColor(foto, cv2.COLOR_BGR2RGB)
            locs = face_recognition.face_locations(rgb)
            if locs:
                enc = face_recognition.face_encodings(rgb, locs)[0]
                lista_enc.append(enc)
            else:
                fotos_sin_cara += 1

        if not lista_enc:
            messagebox.showerror(
                "No se pudo registrar",
                "Ninguna de las fotos tuvo una cara reconocible.\n\n"
                "Intentá de nuevo con:\n"
                "• Mejor iluminación\n"
                "• Cara más centrada y de frente\n"
                "• Sin obstáculos (manos, objetos)"
            )
            return

        encodings[nombre] = lista_enc
        guardar_encodings(encodings)

        msg = f"Profesor '{nombre}' registrado con {len(lista_enc)} foto(s) válida(s)."
        if fotos_sin_cara > 0:
            msg += f"\n({fotos_sin_cara} foto(s) descartadas por no tener cara clara)"

        messagebox.showinfo("✅ Listo", msg)
        self.destroy()


# ──────────────────────────────────────────
# VENTANA: INICIAR MARCAS
# ──────────────────────────────────────────
class VentanaMarcas(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Marcas — Reconocimiento Facial")
        self.geometry("520x320")
        self.resizable(False, False)
        self.configure(bg="#f5f5f0")
        self._corriendo = False   # flag para detener el hilo de cámara
        self._hilo = None
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._al_cerrar)

    def _build_ui(self):
        tk.Label(self, text="Iniciar Marcas", font=("Helvetica", 14, "bold"),
                 bg="#f5f5f0").pack(pady=(16, 4))

        tk.Label(self,
                 text="La cámara se abrirá y reconocerá a los profesores automáticamente.\n"
                      "Presiona ESC en la cámara o el botón Detener para parar.",
                 font=("Helvetica", 11), bg="#f5f5f0", fg="#555", justify="center"
                 ).pack(pady=8)

        self.log = tk.Text(self, height=7, width=56, font=("Courier", 10),
                           state="disabled", bg="#1e1e1e", fg="#00ff88",
                           relief="flat", padx=8, pady=6)
        self.log.pack(padx=20, pady=6)

        btn_frame = tk.Frame(self, bg="#f5f5f0")
        btn_frame.pack(pady=8)

        self.btn_iniciar = tk.Button(
            btn_frame, text="▶  Iniciar cámara",
            font=("Helvetica", 12), bg="#0F6E56", fg="white",
            relief="flat", cursor="hand2", padx=12, pady=6,
            command=self.iniciar)
        self.btn_iniciar.pack(side="left", padx=6)

        self.btn_detener = tk.Button(
            btn_frame, text="⏹  Detener",
            font=("Helvetica", 12), bg="#888", fg="white",
            relief="flat", cursor="hand2", padx=12, pady=6,
            state="disabled", command=self.detener)
        self.btn_detener.pack(side="left", padx=6)

    def log_msg(self, msg):
        """Seguro para llamar desde cualquier hilo via after()."""
        def _escribir():
            self.log.config(state="normal")
            self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.config(state="disabled")
        self.after(0, _escribir)

    def iniciar(self):
        # Guard: no iniciar si ya hay un hilo corriendo
        if self._corriendo:
            messagebox.showwarning("Cámara activa", "La cámara ya está en uso. Presiná Detener primero.")
            return

        encodings = cargar_encodings()
        if not encodings:
            messagebox.showwarning("Sin profesores",
                                   "No hay profesores registrados. Regístralos primero.")
            return

        todos_enc = [enc for lista in encodings.values() for enc in lista]
        etiquetas = [nombre for nombre, lista in encodings.items() for _ in lista]

        # Intentar abrir la cámara con reintentos
        cap = None
        for intento in range(3):
            cap = cv2.VideoCapture(0)
            if cap.isOpened():
                break
            cap.release()
            cap = None

        if cap is None:
            messagebox.showerror(
                "Cámara no disponible",
                "No se pudo conectar a la cámara después de 3 intentos.\n\n"
                "Verificá que:\n"
                "• La cámara esté conectada al equipo\n"
                "• No esté siendo usada por otra aplicación\n"
                "• Los drivers estén instalados correctamente"
            )
            return

        self._corriendo = True
        self.btn_iniciar.config(state="disabled")
        self.btn_detener.config(state="normal")

        import threading
        self._hilo = threading.Thread(
            target=self._loop_camara,
            args=(cap, todos_enc, etiquetas),
            daemon=True
        )
        self._hilo.start()

    def detener(self):
        self._corriendo = False
        self.btn_iniciar.config(state="normal")
        self.btn_detener.config(state="disabled")

    def _al_cerrar(self):
        self._corriendo = False
        self.destroy()

    def _loop_camara(self, cap, todos_enc, etiquetas):
        """Corre en un hilo secundario — NUNCA toca widgets Tk directamente."""
        self.after(0, lambda: self.log_msg("▶ Cámara iniciada. Esperando profesores..."))
        ultimo_reconocido = {}

        fallos_consecutivos = 0
        MAX_FALLOS = 30  # ~1 segundo a 30fps

        while self._corriendo:
            ret, frame = cap.read()
            if not ret:
                fallos_consecutivos += 1
                if fallos_consecutivos >= MAX_FALLOS:
                    self.log_msg("⚠  Cámara desconectada o sin señal.")
                    self.after(0, lambda: messagebox.showerror(
                        "Cámara perdida",
                        "La cámara se desconectó durante la sesión.\n"
                        "Reconectá el dispositivo y volvé a iniciar."
                    ))
                    break
                continue
            fallos_consecutivos = 0

            small = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
            rgb_small = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)

            locs    = face_recognition.face_locations(rgb_small)
            encodes = face_recognition.face_encodings(rgb_small, locs)

            for encode_actual, loc in zip(encodes, locs):
                distancias = face_recognition.face_distance(todos_enc, encode_actual)
                idx = np.argmin(distancias)

                if distancias[idx] < 0.5:
                    nombre = etiquetas[idx]

                    ahora = datetime.now()
                    ultima = ultimo_reconocido.get(nombre)
                    if ultima and (ahora - ultima).seconds < 10:
                        continue

                    try:
                        tipo, hora = registrar_marca(nombre)
                    except PermissionError as e:
                        self.log_msg(f"⚠  {str(e)}")
                        self.after(0, lambda msg=str(e):
                                   messagebox.showwarning("Archivo bloqueado", msg))
                        continue

                    ultimo_reconocido[nombre] = ahora
                    self.log_msg(f"[{hora}]  {nombre}  →  {tipo}")

                    # Foto del profesor
                    foto_path = None
                    carpeta_fotos = FOTOS_DIR / nombre
                    if carpeta_fotos.exists():
                        fotos = list(carpeta_fotos.glob("*.jpg"))
                        if fotos:
                            foto_path = str(fotos[0])

                    # Notificación — siempre via after() para ejecutar en hilo principal
                    self.after(0, lambda n=nombre, t=tipo, h=hora, f=foto_path:
                               NotificacionMarca(self.master, n, t, h, f))

                    top, right, bottom, left = [v*4 for v in loc]
                    color = (0, 200, 0) if tipo == "Entrada" else (0, 100, 255)
                    cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
                    cv2.putText(frame, f"{nombre} ({tipo})",
                                (left, top - 10),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)
                else:
                    top, right, bottom, left = [v*4 for v in loc]
                    cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 200), 2)
                    cv2.putText(frame, "Desconocido",
                                (left, top - 10),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 200), 2)

            cv2.putText(frame, "ESC para detener", (10, 30),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (200, 200, 200), 2)
            cv2.imshow("Marcas en curso", frame)

            if cv2.waitKey(1) == 27:
                self._corriendo = False
                break

        cap.release()
        cv2.destroyAllWindows()
        self.log_msg("⏹ Cámara detenida.")
        self.after(0, lambda: self.btn_iniciar.config(state="normal"))
        self.after(0, lambda: self.btn_detener.config(state="disabled"))


# ──────────────────────────────────────────
# VENTANA: VER REPORTES
# ──────────────────────────────────────────
class VentanaReportes(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Reportes de Asistencia")
        self.geometry("680x460")
        self.configure(bg="#f5f5f0")
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="Reportes de Asistencia", font=("Helvetica", 14, "bold"),
                 bg="#f5f5f0").pack(pady=(14, 4))

        # Filtro por fecha
        filtro_frame = tk.Frame(self, bg="#f5f5f0")
        filtro_frame.pack(pady=4)
        tk.Label(filtro_frame, text="Filtrar por fecha (YYYY-MM-DD):",
                 bg="#f5f5f0", font=("Helvetica", 10)).pack(side="left", padx=4)
        self.fecha_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        tk.Entry(filtro_frame, textvariable=self.fecha_var,
                 font=("Helvetica", 10), width=14).pack(side="left", padx=4)
        tk.Button(filtro_frame, text="🔍 Filtrar", bg="#3C3489", fg="white",
                  relief="flat", cursor="hand2", font=("Helvetica", 10),
                  command=self.cargar_datos).pack(side="left", padx=4)
        tk.Button(filtro_frame, text="Ver todo", bg="#888", fg="white",
                  relief="flat", cursor="hand2", font=("Helvetica", 10),
                  command=self.ver_todo).pack(side="left", padx=4)

        # Tabla
        cols = ("Nombre", "Fecha", "Hora", "Tipo")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")
        self.tree.pack(padx=20, pady=8, fill="both", expand=True)

        # Scrollbar
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)

        # Botón exportar
        tk.Button(self, text="📥  Abrir archivo Excel", bg="#854F0B", fg="white",
                  font=("Helvetica", 10), relief="flat", cursor="hand2",
                  command=self.abrir_excel).pack(pady=6)

        self.cargar_datos()

    def cargar_datos(self):
        self._limpiar_tabla()
        if not MARCAS_FILE.exists():
            return
        df = pd.read_excel(MARCAS_FILE)
        fecha = self.fecha_var.get().strip()
        if fecha:
            df = df[df["Fecha"] == fecha]
        for _, row in df.iterrows():
            tag = "entrada" if row["Tipo"] == "Entrada" else "salida"
            self.tree.insert("", "end",
                             values=(row["Nombre"], row["Fecha"], row["Hora"], row["Tipo"]),
                             tags=(tag,))
        self.tree.tag_configure("entrada", foreground="#0F6E56")
        self.tree.tag_configure("salida",  foreground="#854F0B")

    def ver_todo(self):
        self.fecha_var.set("")
        self.cargar_datos()

    def _limpiar_tabla(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def abrir_excel(self):
        if not MARCAS_FILE.exists():
            messagebox.showinfo("Sin datos", "Aún no hay marcas registradas.")
            return
        import subprocess, sys
        if sys.platform == "win32":
            os.startfile(str(MARCAS_FILE))
        elif sys.platform == "darwin":
            subprocess.call(["open", str(MARCAS_FILE)])
        else:
            subprocess.call(["xdg-open", str(MARCAS_FILE)])


# ──────────────────────────────────────────
# ENTRADA
# ──────────────────────────────────────────
if __name__ == "__main__":
    app = AppPrincipal()
    app.mainloop()
