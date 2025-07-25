import mysql.connector
import pandas as pd
from openpyxl import load_workbook
import os
from datetime import datetime

class OCSInventoryToExcel:
    def __init__(self, db_config, template_path):
        """
        Inicializa la clase con configuración de BD y ruta de plantilla
        
        Args:
            db_config (dict): Configuración de la base de datos
            template_path (str): Ruta de la plantilla Excel
        """
        self.db_config = db_config
        self.template_path = template_path
        self.connection = None
        
    def connect_database(self):
        """Conecta a la base de datos MySQL de OCS Inventory"""
        try:
            self.connection = mysql.connector.connect(**self.db_config)
            print("Conexión exitosa a la base de datos OCS Inventory")
            return True
        except mysql.connector.Error as err:
            print(f"Error conectando a la base de datos: {err}")
            return False
    
    def get_devices_data(self):
        """
        Extrae información de dispositivos desde OCS Inventory
        
        Returns:
            list: Lista de diccionarios con información de cada dispositivo
        """
        if not self.connection:
            print("No hay conexión a la base de datos")
            return []
        
        query = """
        SELECT DISTINCT
            h.NAME as username,
            h.OSNAME as device_type,
            b.SMANUFACTURER as manufacturer,
            b.SMODEL as model,
            b.SSN as serial_number,
            h.ID as hardware_id
        FROM hardware h
        LEFT JOIN bios b ON h.ID = b.HARDWARE_ID
        WHERE h.NAME IS NOT NULL AND h.NAME != ''
        ORDER BY h.NAME
        """
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute(query)
            devices = cursor.fetchall()
            cursor.close()
            
            # Para cada dispositivo, obtener monitores, teclados y mouse
            for device in devices:
                device['monitors'] = self.get_monitors(device['hardware_id'])
                device['keyboards'] = self.get_keyboards(device['hardware_id'])
                device['mice'] = self.get_mice(device['hardware_id'])
            
            print(f"Se encontraron {len(devices)} dispositivos")
            return devices
            
        except mysql.connector.Error as err:
            print(f"Error ejecutando consulta: {err}")
            return []
    
    def get_monitors(self, hardware_id):
        """Obtiene información de monitores conectados"""
        query = """
        SELECT 
            MANUFACTURER as brand,
            CAPTION as identifier,
            SERIAL as serial_number
        FROM monitors 
        WHERE HARDWARE_ID = %s
        """
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute(query, (hardware_id,))
            monitors = cursor.fetchall()
            cursor.close()
            return monitors
        except:
            return []
    
    def get_keyboards(self, hardware_id):
        """Obtiene información de teclados conectados"""
        query = """
        SELECT 
            NAME as brand,
            DESCRIPTION as identifier,
            '' as serial_number
        FROM inputs 
        WHERE HARDWARE_ID = %s AND TYPE = 'Keyboard'
        """
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute(query, (hardware_id,))
            keyboards = cursor.fetchall()
            cursor.close()
            return keyboards
        except:
            return []
    
    def get_mice(self, hardware_id):
        """Obtiene información de mouse conectados"""
        query = """
        SELECT 
            NAME as brand,
            DESCRIPTION as identifier,
            '' as serial_number
        FROM inputs 
        WHERE HARDWARE_ID = %s AND TYPE = 'Mouse'
        """
        
        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute(query, (hardware_id,))
            mice = cursor.fetchall()
            cursor.close()
            return mice
        except:
            return []
    
    def create_excel_for_user(self, device_data, output_folder):
        """
        Crea un Excel individual para cada usuario usando la plantilla
        
        Args:
            device_data (dict): Datos del dispositivo y usuario
            output_folder (str): Carpeta donde guardar los archivos
        """
        try:
            # Cargar la plantilla
            workbook = load_workbook(self.template_path)
            worksheet = workbook.active
            
            # Basándome en la plantilla PDF, necesito que me confirmes las celdas exactas
            # Datos del colaborador que recibe
            worksheet['O21'] = device_data.get('username', '')  # Nombre del colaborador
            
            # Fecha actual
            worksheet['O22'] = datetime.now().strftime('%d-%m-%Y')  # Fecha
            worksheet['O23'] = datetime.now().strftime('%H:%M')    # Hora
            
            # Datos del equipo principal en la tabla
            # Fila del primer equipo (ajustar según la plantilla Excel real)
            equipment_row = 15  # Estimado, necesita confirmación
            
            # Columnas de la tabla de equipos (estimadas, necesitan confirmación)
            worksheet[f'A{equipment_row}'] = '1'  # Número
            worksheet[f'B{equipment_row}'] = self.determine_equipment_type(device_data)  # Descripción/Tipo
            worksheet[f'C{equipment_row}'] = 'En funcionamiento'  # Estado
            worksheet[f'D{equipment_row}'] = device_data.get('manufacturer', '')  # Marca
            worksheet[f'E{equipment_row}'] = device_data.get('model', '')  # Modelo
            worksheet[f'F{equipment_row}'] = device_data.get('serial_number', '')  # Serie
            worksheet[f'G{equipment_row}'] = f"Usuario: {device_data.get('username', '')}"  # Observaciones
            
            # Agregar monitores como equipos adicionales
            current_row = equipment_row + 1
            monitors = device_data.get('monitors', [])
            for i, monitor in enumerate(monitors):
                if current_row <= equipment_row + 10:  # Limitar a 10 filas adicionales
                    worksheet[f'A{current_row}'] = str(i + 2)
                    worksheet[f'B{current_row}'] = 'Monitor'
                    worksheet[f'C{current_row}'] = 'En funcionamiento'
                    worksheet[f'D{current_row}'] = monitor.get('brand', '')
                    worksheet[f'E{current_row}'] = monitor.get('identifier', '')
                    worksheet[f'F{current_row}'] = monitor.get('serial_number', '')
                    worksheet[f'G{current_row}'] = f"Monitor {i+1}"
                    current_row += 1
            
            # Agregar teclados
            keyboards = device_data.get('keyboards', [])
            for i, keyboard in enumerate(keyboards):
                if current_row <= equipment_row + 10:
                    worksheet[f'A{current_row}'] = str(current_row - equipment_row + 1)
                    worksheet[f'B{current_row}'] = 'Teclado'
                    worksheet[f'C{current_row}'] = 'En funcionamiento'
                    worksheet[f'D{current_row}'] = keyboard.get('brand', '')
                    worksheet[f'E{current_row}'] = keyboard.get('identifier', '')
                    worksheet[f'F{current_row}'] = keyboard.get('serial_number', 'N/A')
                    worksheet[f'G{current_row}'] = f"Teclado {i+1}"
                    current_row += 1
            
            # Agregar mouse
            mice = device_data.get('mice', [])
            for i, mouse in enumerate(mice):
                if current_row <= equipment_row + 10:
                    worksheet[f'A{current_row}'] = str(current_row - equipment_row + 1)
                    worksheet[f'B{current_row}'] = 'Mouse'
                    worksheet[f'C{current_row}'] = 'En funcionamiento'
                    worksheet[f'D{current_row}'] = mouse.get('brand', '')
                    worksheet[f'E{current_row}'] = mouse.get('identifier', '')
                    worksheet[f'F{current_row}'] = mouse.get('serial_number', 'N/A')
                    worksheet[f'G{current_row}'] = f"Mouse {i+1}"
                    current_row += 1
            
            # Crear nombre de archivo seguro
            username = device_data.get('username', 'Usuario_Desconocido')
            safe_filename = "".join(c for c in username if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filename = f"Acta_Entrega_{safe_filename}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            filepath = os.path.join(output_folder, filename)
            
            # Guardar el archivo
            workbook.save(filepath)
            print(f"Acta creada: {filename}")
            
        except Exception as e:
            print(f"Error creando acta para {device_data.get('username', 'usuario')}: {e}")
    
    def determine_equipment_type(self, device_data):
        """Determina el tipo de equipo basado en el OS"""
        os_name = device_data.get('device_type', '').lower()
        if 'windows' in os_name:
            if 'server' in os_name:
                return 'Servidor'
            else:
                return 'PC Escritorio/Laptop'
        elif 'linux' in os_name:
            return 'Estación Linux'
        elif 'mac' in os_name:
            return 'Mac'
        else:
            return 'Equipo Informático'
    
    def generate_all_excel_files(self, output_folder="output_inventarios"):
        """
        Genera todos los archivos Excel automáticamente
        
        Args:
            output_folder (str): Carpeta donde guardar todos los archivos
        """
        # Crear carpeta de salida si no existe
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        # Conectar a la base de datos
        if not self.connect_database():
            return False
        
        # Obtener datos de todos los dispositivos
        devices_data = self.get_devices_data()
        
        if not devices_data:
            print("No se encontraron datos para procesar")
            return False
        
        # Generar Excel para cada usuario
        print(f"\nGenerando {len(devices_data)} archivos Excel...")
        for device in devices_data:
            self.create_excel_for_user(device, output_folder)
        
        # Cerrar conexión
        if self.connection:
            self.connection.close()
        
        print(f"\nProceso completado. Archivos guardados en: {output_folder}")
        return True

# Configuración y uso del script
if __name__ == "__main__":
    # Configuración de la base de datos (AJUSTAR CON TUS DATOS)
    db_config = {
        'host': 'localhost',
        'database': 'ocsweb',  # Nombre típico de la BD de OCS
        'user': 'ocsuser',  # Cambiar por tu usuario
        'password': 'ocspass',  # Cambiar por tu password
        'port': 3306
    }
    
    # Ruta de tu plantilla Excel (CAMBIAR POR LA RUTA REAL)
    template_path = "plantilla_inventario.xlsx"
    
    # Crear instancia y generar archivos
    generator = OCSInventoryToExcel(db_config, template_path)
    
    # Generar todos los archivos Excel automáticamente
    generator.generate_all_excel_files("inventarios_generados")