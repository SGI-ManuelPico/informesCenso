import mysql.connector
from tkinter import messagebox

class conexionDB:
    def __init__(self):
        self.host = ''
        self.user = ''
        self.password = ''
        self.database = ''
        self.conexion = None
        
    def establecer_conexion(self):
        try:
            self.conexion = mysql.connector.connect(
                host = self.host,
                user = self.user,
                password = self.password,
                database = self.database
            )
            return self.conexion
        
        except mysql.connector.Error as e:
            messagebox.showerror(message=f'Error de conexión: {e}', title='Mensaje')
            return None
        
    def cerrar_conexion(self):
        if self.conexion:
            self.conexion.close()