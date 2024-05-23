import sys
import os
import subprocess
import win32com.client as win32
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QComboBox, QPushButton, QLabel, QMessageBox, QLineEdit, QHBoxLayout, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import openpyxl
import shutil
import re
from datetime import datetime
import logging

class MiVentana(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Memo Automático - Inchcape")
        self.setGeometry(100, 100, 600, 400)
        self.setWindowIcon(QIcon('path_to_icon.png'))  # Asegúrate de que el ícono exista
        logging.basicConfig(filename='app.log', level=logging.ERROR)  # Configura el archivo de registro y el nivel de registro

        titulo_style = """
            QLabel {
                font-family: 'Consolas', 'Arial';
                font-size: 28px;
                color: #C0C0C0;
                padding: 20px;
            }
        """

        boton_style = """
            QPushButton {
                font: 15px;
                color: #FFFFFF;
                background-color: #00A2E8;
                border-radius: 10px;
                padding: 10px;
                margin: 5px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #008CBA;
            }
            QPushButton:pressed {
                background-color: #007399;
            }
        """

        self.setStyleSheet("""
            QMainWindow {
                background-color: #2c3e50;
            }
            QPushButton {
                font: 15px;
                color: #ecf0f1;
                background-color: #3498db;
                border: none;
                padding: 10px;
                margin: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QComboBox {
                font: 15px;
                color: #5DADE2;
                margin: 5px;
                padding: 5px;
            }
        """ + boton_style)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        titulo = QLabel("Memo Automático - Inchcape", self)
        titulo.setAlignment(Qt.AlignCenter)
        titulo.setStyleSheet(titulo_style)

        self.listado_dt = QComboBox()
        self.listado_proveedor = QComboBox()

        self.archivo_excel = self.cargar_archivo_excel()
        self.cargar_valores_proveedor()
        self.listado_proveedor.currentTextChanged.connect(self.actualizar_dt_por_proveedor)

        layout.addWidget(titulo)
        layout.addWidget(self.listado_proveedor)
        layout.addWidget(self.listado_dt)

        self.boton_correo = QPushButton("Enviar Correo", self)
        self.boton_correo.clicked.connect(self.enviar_correo)
        layout.addWidget(self.boton_correo)
        
        self.boton_copiar_archivos = QPushButton("Copiar Hojas de Seguridad", self)
        self.boton_copiar_archivos.clicked.connect(self.copiar_archivos_material_oc)
        layout.addWidget(self.boton_copiar_archivos)

    def copiar_archivos_material_oc(self):
        dt_seleccionado = self.listado_dt.currentText()
        datos_dt = self.obtener_datos_dt(dt_seleccionado)
        materiales_oc = set(datos_dt["Material OC"])  # Convertir la lista a un conjunto para evitar duplicados
        nombre_usuario = os.getlogin()

        ruta_origen = fr"C:/Users/{nombre_usuario}/OneDrive - Inchcape/00 - HDS VIGENTES"
        carpeta_base = fr"C:/Users/{nombre_usuario}/OneDrive - Inchcape/AFM(Recuperado ok)"
        patron_busqueda = f"DT {dt_seleccionado}"
        carpeta_destino_encontrada = None

        # Buscar la carpeta destino correcta dentro de las subcarpetas
        for dirpath, dirnames, filenames in os.walk(carpeta_base):
            if patron_busqueda in dirpath:
                carpeta_destino_encontrada = dirpath
                break

        if materiales_oc and carpeta_destino_encontrada:
            try:
                archivos_encontrados = set()  # Usar un conjunto para rastrear los archivos ya copiados
                for material_oc in materiales_oc:
                    print(f"Buscando archivo para Material OC: {material_oc}")
                    for dirpath, dirnames, filenames in os.walk(ruta_origen):
                        for nombre_archivo in filenames:
                            if re.search(re.escape(material_oc), nombre_archivo, re.IGNORECASE) and material_oc not in archivos_encontrados:
                                ruta_archivo_origen = os.path.join(dirpath, nombre_archivo)
                                archivos_encontrados.add(material_oc)
                                shutil.copy(ruta_archivo_origen, carpeta_destino_encontrada)
                                break  # Romper el bucle interno una vez que se encuentra un archivo para el material OC

                if archivos_encontrados:
                    print(f"Archivos copiados para los siguientes Materiales OC: {archivos_encontrados}")
                    QMessageBox.information(self, "Éxito", f"Archivos relacionados a los Materiales OC copiados a DT {dt_seleccionado}.")
                else:
                    print(f"No se encontraron archivos para los Materiales OC.")
                    QMessageBox.warning(self, "No se encontraron archivos", "No se encontraron archivos para los Materiales OC.")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Ocurrió un error al copiar archivos: {e}")
        else:
            QMessageBox.warning(self, "Error", "No se encontró la carpeta destino o no hay Materiales OC.")



    def cargar_archivo_excel(self):
        nombre_usuario = os.getlogin()
        ruta_excel = fr"C:/Users/{nombre_usuario}/OneDrive - Inchcape/Macro Memo/df_app.xlsx"
        try:
            archivo_excel = openpyxl.load_workbook(ruta_excel)
            return archivo_excel["Sheet1"]
        except FileNotFoundError:
            print(f"No se pudo encontrar el archivo: {ruta_excel}")
            return None

    def cargar_valores_proveedor(self):
        if self.archivo_excel:
            valores_proveedor = {}
            for celda in self.archivo_excel["D"]:
                if celda.value:
                    texto_procesado = self.procesar_texto_excel(celda.value)
                    valores_proveedor[texto_procesado] = True
            self.listado_proveedor.addItems(sorted(valores_proveedor.keys()))

    def cargar_valores_dt(self, proveedor_seleccionado):
        if self.archivo_excel:
            valores_dt = {}
            for celda in self.archivo_excel["A"]:
                if celda.value and self.archivo_excel[f"D{celda.row}"].value == proveedor_seleccionado:
                    valores_dt[celda.value] = True
            self.listado_dt.clear()
            self.listado_dt.addItems(sorted(valores_dt.keys()))

    def actualizar_dt_por_proveedor(self):
        proveedor_seleccionado = self.listado_proveedor.currentText()
        proveedor_seleccionado = self.procesar_texto_excel(proveedor_seleccionado)  # Procesar texto
        self.cargar_valores_dt(proveedor_seleccionado)

    def obtener_datos_dt(self, dt_seleccionado):
        try:
            if self.archivo_excel:
                datos_dt = {
                    "Nro DT": set(),
                    "Referencia": set(),
                    "FE.ATA": set(),
                    "CBE": set(),
                    "Entrega entrante": set(),
                    "Vía (Texto)": set(),
                    "Documento de embarque": set(),
                    "Proveedor": set(),
                    "Contenedor": set(),
                    "Valor": set(),
                    "Requiere CDA": set(),
                    "Material OC": set(),
                    "Nave/Aerolínea": set(),
                    "INCOTERM": set(),
                    "MONEDA": set()
                }

                for fila in self.archivo_excel.iter_rows(min_row=2):
                    if fila[0].value == dt_seleccionado:
                        datos_dt["Nro DT"].add(fila[0].value)
                        
                        # Aquí agregamos el manejo de codificación para la Referencia
                        referencia = fila[16].value
                        if referencia:
                            referencia = referencia.encode('latin-1').decode('utf-8', 'ignore')
                        datos_dt["Referencia"].add(referencia)
                        
                        # Verificar si el valor en la columna es un objeto datetime para aplicar strftime
                        if isinstance(fila[10].value, datetime):
                            datos_dt["FE.ATA"].add(fila[10].value.strftime('%d-%m-%Y'))
                        else:
                            datos_dt["FE.ATA"].add('')  # Dejar vacío si no es una fecha
                        
                        # Los demás campos se pueden manejar de manera similar
                        datos_dt["CBE"].add(fila[22].value)
                        datos_dt["Entrega entrante"].add(fila[2].value)
                        datos_dt["Vía (Texto)"].add(fila[6].value)
                        datos_dt["Documento de embarque"].add(fila[9].value)
                        datos_dt["Proveedor"].add(fila[3].value)
                        datos_dt["Contenedor"].add(fila[11].value)
                        datos_dt["Valor"].add(fila[17].value)
                        datos_dt["Requiere CDA"].add(fila[18].value)
                        datos_dt["Material OC"].add(fila[12].value)
                        datos_dt["Nave/Aerolínea"].add(fila[8].value)
                        datos_dt["INCOTERM"].add(fila[15].value)
                        datos_dt["MONEDA"].add(fila[1].value)
                        
                return datos_dt
        except Exception as e:
            print(f"Error al obtener datos del archivo Excel: {e}")
            return {}
        
    def procesar_texto_excel(self, texto):
        try:
            # Intenta decodificar con la codificación esperada si no es None o ya está en str
            if texto and not isinstance(texto, str):
                texto = texto.encode('latin-1').decode('utf-8', 'ignore')
        except AttributeError:
            # Si el texto no tiene el método encode (por ejemplo, números), lo deja tal cual
            pass
        return texto
    
    def crear_cuerpo_correo(self, datos_dt, dt_seleccionado):
        # Estilo CSS para el correo
        estilo_css = """
        <style>
            body {
                font-family: Arial, sans-serif;
                background-color: #f4f4f4;
                color: #333;
                margin: 0;
                padding: 0;
            }
            .container {
                max-width: 600px;
                margin: auto;
                background: white;
                padding: 20px;
            }
            table {
                width: 100%;
                border-collapse: collapse;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 8px;
                text-align: left;
            }
            th {
                background-color: #00008B;
                color: white;
            }
            tr:nth-child(even) {
                background-color: #f2f2f2;
            }
            .observaciones th {
                background-color: #00008B;
                color: white;
                padding: 8px;
            }
            .observaciones td {
                background-color: #f4f4f4;
                color: black;
                padding: 8px;
            }
        </style>
        """
        via_texto = self.procesar_texto_excel(next(iter(datos_dt["Vía (Texto)"]), ''))
        proveedor = self.procesar_texto_excel(next(iter(datos_dt["Proveedor"]), ''))
        cuerpo_correo = "<html><head>" + estilo_css + "</head><body>"
        cuerpo_correo += f"<div class='container'><p>Estimad@,</p>"
        cuerpo_correo += f"<p>Envío memo {via_texto} correspondiente a {proveedor}.</p>"
        cuerpo_correo += "<table>"

        # Configurar las columnas según el valor de "Vía (Texto)"
        if via_texto.lower() in ["maritimo", "terrestre", "camión"]:
            encabezados = ["Nro DT", "Referencia", "FE.ATA", "CBE", "Documento de embarque", "Contenedor", "Valor", "MONEDA", "NAVE", "INCOTERM"]
        else:
            encabezados = ["Nro DT", "FE.ATA", "CBE", "Entrega entrante", "Documento de embarque", "Factura"]
            datos_dt["Factura"] = set(["N/A" for _ in range(len(datos_dt["Nro DT"]))])

        # Crear encabezados de la tabla
        cuerpo_correo += "<tr>"
        for encabezado in encabezados:
            cuerpo_correo += f"<th>{encabezado}</th>"
        cuerpo_correo += "</tr>"

        # Crear filas de la tabla
        num_rows = max(len(datos_dt[encabezado]) for encabezado in encabezados if encabezado in datos_dt)
        for i in range(num_rows):
            cuerpo_correo += "<tr>"
            for encabezado in encabezados:
                valores = list(datos_dt[encabezado]) if encabezado in datos_dt else []
                valor = self.procesar_texto_excel(valores[i] if i < len(valores) else '')
                cuerpo_correo += f"<td>{valor}</td>"
            cuerpo_correo += "</tr>"
        cuerpo_correo += "</table>"

        # Agregar observaciones para marítimo
        if via_texto.lower() in ["maritimo", "terrestre", "camión"]:
            cuerpo_correo += """
            <br><br>
            <table class='observaciones'>
                <tr>
                    <th colspan='2'>OBSERVACIONES</th>
                </tr>
                <tr>
                    <td>CARGA</td>
                    <td>NORMAL</td>
                </tr>
                <tr>
                    <td>BL/CR/AWB</td>
                    <td>ADJUNTO</td>
                </tr>
                <tr>
                    <td>C.O.</td>
                    <td>ADJUNTO</td>
                </tr>
                <tr>
                    <td>OTRAS</td>
                    <td></td>
                </tr>
            </table>
            """

        # Añadir la firma
        firma = self.obtener_firma()
        if firma:
            cuerpo_correo += "<br>" + firma

        cuerpo_correo += "</body></html>"
        return cuerpo_correo

    # Al leer el archivo de la firma:
    def obtener_firma(self):
        ruta_firmas = os.path.join(os.environ['APPDATA'], 'Microsoft\\Signatures')
        firma_html = None

        if os.path.isdir(ruta_firmas):
            for archivo in os.listdir(ruta_firmas):
                if archivo.lower().endswith('.htm'):
                    ruta_completa = os.path.join(ruta_firmas, archivo)
                    try:
                        with open(ruta_completa, 'r', encoding='utf-8', errors='replace') as f:
                            firma_html = f.read()
                        # Asegúrate de ajustar las rutas de imágenes u otros recursos en la firma
                        firma_html = firma_html.replace('src="', f'src="file:///{ruta_firmas}/')
                        break  # Solo usa la primera firma encontrada
                    except Exception as e:
                        print(f"Error al leer el archivo de firma: {e}")
        return firma_html
    def enviar_correo(self):
        try:
            dt_seleccionado = self.listado_dt.currentText()
            datos_dt = self.obtener_datos_dt(dt_seleccionado)
            cuerpo_correo = self.crear_cuerpo_correo(datos_dt, dt_seleccionado)

            # Obtener 'Vía (Texto)'
            via_texto = next(iter(datos_dt["Vía (Texto)"]), '').lower()

            # Seleccionar los destinatarios según 'Vía (Texto)'
            if via_texto == "aereo":
                destinatarios = "scarlette.tapia.dre@teamworkchile.cl;danielamardones@derco.cl;lst_cl_analistaabastecimientooem@derco.cl;lst_comex_dercoparts@derco.cl; Lst_Administracion_ComprasCd@derco.cl"
            elif via_texto == "courier":
                destinatarios = "sandrarojas@derco.cl;lst_cl_analistaabastecimientooem@derco.cl;lst_comex_dercoparts@derco.cl; Lst_Administracion_ComprasCd@derco.cl"
            elif via_texto in ["maritimo", "terrestre", "camión"]:
                destinatarios = "danielamardones@derco.cl;lst_comex_dercoparts@derco.cl;lst_cl_analista_abastecimiento_aftermarket@derco.cl; Lst_Administracion_ComprasCd@derco.cl"
            else:
                destinatarios = "lst_comex_dercoparts@derco.cl"  # Coloca aquí un correo por defecto o gestiona esta situación

            # Construcción del Subject del correo
            doc_embarque = self.procesar_texto_excel(next(iter(datos_dt["Documento de embarque"]), ''))
            proveedor = self.procesar_texto_excel(next(iter(datos_dt["Proveedor"]), ''))
            subject = f"(MEMO COMEX) EMBARQUE {via_texto.upper()} DT {dt_seleccionado} // {doc_embarque} // {proveedor}"

            # Configuración y envío del correo
            outlook = win32.Dispatch('outlook.application')
            correo = outlook.CreateItem(0)
            correo.To = destinatarios
            correo.Subject = subject
            correo.HTMLBody = cuerpo_correo

            # Agregar archivos adjuntos
            nombre_usuario = os.getlogin()
            carpetas_base = [
                f"C:/Users/{nombre_usuario}/OneDrive - Inchcape/AFM(Recuperado ok)",
                f"C:/Users/{nombre_usuario}/OneDrive - Inchcape/OEM(Recuperado ok)"
            ]

            for carpeta_base in carpetas_base:
                if os.path.isdir(carpeta_base):
                    patron_busqueda = f"DT {dt_seleccionado}"
                    for dirpath, dirnames, filenames in os.walk(carpeta_base):
                        if patron_busqueda in dirpath:
                            for archivo in filenames:
                                if not archivo.lower().startswith('carga'):
                                    ruta_completa = os.path.join(dirpath, archivo)
                                    correo.Attachments.Add(ruta_completa)

            correo.Display()

        except Exception as e:
            logging.error(f"Error al enviar el correo: {e}")
            QMessageBox.warning(self, "Error", f"Ocurrió un error al enviar el correo: {e}")


def main():
    app = QApplication(sys.argv)
    ventana = MiVentana()
    ventana.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()


