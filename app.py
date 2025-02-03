
import pymysql
import paramiko
import tkinter as tk
from ttkthemes import ThemedTk
from tkinter import ttk, messagebox, Toplevel, Label, StringVar
import os
import threading
import time

usuario = os.getlogin()

def conectar_bd():
    try:
        conn = pymysql.connect(
            host='192.168.40.181',
            user='bases_digicom',
            password='@DIGICOM2025*',
            database='correo'
        )
        print("conexion establecida")
        return conn
    except pymysql.MySQLError as e:
        print("Error al conectarse: {e}")
        return None

def registrar_transferencia(usuario, estado):
    con = conectar_bd()
    cur = con.cursor()
    cur.execute("INSERT INTO registro_correo (usuario, estado) VALUES (%s, %s)", (usuario, estado))
    con.commit()
    cur.close()
    con.close()
    
def mostrar_datos():
    con = conectar_bd()
    cur = con.cursor()
    cur.execute("SELECT * FROM registro_correo")
    registros = cur.fetchall()
    cur.close()
    con.close()
    
    #limpia el tree
    for item in tree.get_children():
        tree.delete(item)
    
    #inserta datos
    for registro in registros:
        tree.insert("", "end", values=(registro[1], registro[2], registro[3]))

def SFTP_conectar():
    servidor = '10.10.10.40'
    port = 22
    usuario = 't-i'
    contrasena = 'd1g1c0m2012'
    
    transporte = paramiko.Transport((servidor, port))
    transporte.connect(username=usuario, password=contrasena)
    sftp = paramiko.SFTPClient.from_transport(transporte)
    return sftp

def origen(): ##INCLUIR SUBcarpetas
    archivos = []
    documentos_pst = os.path.join(os.environ["USERPROFILE"], "Documents", "Archivos de Outlook")
    documentos_ost = os.path.join(os.environ["USERPROFILE"], "AppData", "Local", "Microsoft", "Outlook")
    
    #archivos pst
    for root, dirs, files in os.walk(documentos_pst):
        for file in files:
            if file.endswith('.pst'):
                archivos.append(os.path.join(root, file))
    #archivos ost
    for root, dirs, files in os.walk(documentos_ost):
        for file in files:
            if file.endswith('.ost'):
                archivos.append(os.path.join(root, file))
    return archivos

def destino(): ##INCLUIR SUBCARPETAS
    sftp = SFTP_conectar()
    ruta_base = '/shares/Backup/Correos'
    sftp.chdir(ruta_base)
    ruta_destino = []
    try:
        sftp.chdir(usuario)
    except IOError:
        sftp.mkdir(usuario)
    documentos_pst = os.path.join(os.environ["USERPROFILE"], "Documents", "Archivos de Outlook")
    documentos_ost = os.path.join(os.environ["USERPROFILE"], "AppData", "Local", "Microsoft", "Outlook")
    
    #archovos pst
    for root, dirs, files in os.walk(documentos_pst):
        for file in files:
            if file.endswith('.pst'):
                relative_path = os.path.relpath(os.path.join(root, file), documentos_pst)
                destino = f'{ruta_base}/{usuario}/{relative_path}'
                ruta_destino.append(destino)
    #archivo ost
    for root, dirs, files in os.walk(documentos_ost):
        for file in files:
            if file.endswith('.ost'):
                relative_path = os.path.relpath(os.path.join(root, file), documentos_ost)
                destino = f'{ruta_base}/{usuario}/{relative_path}'
                ruta_destino.append(destino)
    return ruta_destino


def calcular_tamano_archivos(rutas):
    total_tamano = 0
    for ruta in rutas:
        if os.path.exists(ruta):
            total_tamano += os.path.getsize(ruta)
    return total_tamano

def subir():
    progreso = Toplevel(root)
    progreso.title("Transferencia en progreso")
    progreso.geometry("400x200")
    progreso.grab_set()
    
    lblEstado = Label(progreso, text="Iniciando transferencia....")
    lblEstado.pack(pady=10)
    
    barra = ttk.Progressbar(progreso, orient="horizontal", length=300, mode="determinate")
    barra.pack(pady=10)
    
    porcentaje_var = StringVar(value="Progreso: 0%")
    lblProcentaje = Label(progreso, textvariable=porcentaje_var)
    lblProcentaje.pack(pady=10)
    
    tiempo_estimado_var = StringVar(value="Tiempo estimado: Calculando...")
    lblTiempo = Label(progreso, textvariable=tiempo_estimado_var)
    lblTiempo.pack(pady=10)
    
    def transferencia():
        try:
            cerrar_outlook()
            sftp = SFTP_conectar()
            origen_local = origen()
            destino_remoto = destino()
            total_tamano = calcular_tamano_archivos(origen_local)
            tamano_transferido = 0
            
            if len(origen_local) == len(destino_remoto):
                inicio = time.time()
                for archivo, nombre_destino in zip(origen_local, destino_remoto):
                    if os.path.exists(archivo):
                        tamano_archivo = os.path.getsize(archivo)
                        
                        #verifica si el archivo ya existe y tiene el mismo tamano
                        try:
                            tamano_destino = sftp.stat(nombre_destino).st_size
                        except FileNotFoundError:
                            tamano_destino = None
                        
                        if tamano_destino == tamano_archivo:
                            print(f"Archivo '{os.path.basename(archivo)}' ya esta actualizado")
                            tamano_transferido += tamano_archivo
                            continue
                        
                        def actualizar_progreso(sent, total):
                            nonlocal tamano_transferido
                            tamano_transferido += sent
                            if total_tamano > 0:
                                progreso_global = min((tamano_transferido / total_tamano) * 100, 100)
                                barra["value"] = progreso_global
                                porcentaje_var.set(f"Progreso: {int(progreso_global)}%")
                                
                                tiempo_transcurrido = time.time() - inicio
                                velocidad_promedio = tamano_transferido / tiempo_transcurrido if tiempo_transcurrido > 0 else 1
                                tiempo_restante = max((total_tamano - tamano_transferido) / velocidad_promedio, 0)
                                tiempo_estimado_var.set(f"Tiempo estimado: {int(tiempo_restante)}")
                                progreso.update()
                        
                        sftp.put(archivo, nombre_destino, callback=actualizar_progreso)
                estado = "Correcto"
                registrar_transferencia(usuario, estado)
                if usuario in ['cjalfonso', 'JOCASTRO', 'Yhurtado']:
                    mostrar_datos()
                messagebox.showinfo("Exito", "Se ha realizado la copia con exito")
            else:
                estado = "Error"
                registrar_transferencia(usuario, estado)
                mostrar_datos()
                messagebox.showerror("Error", "Se ha producido un error")
        except Exception as e:
            messagebox.showerror("Error", f"Error durante la transferencia: {e}")
            print(f"error durante la trasnferencia: {e}")
        finally:
            progreso.destroy()
    threading.Thread(target=transferencia).start()
    

def cerrar_outlook():
    os.system("taskkill /f /im outlook.exe")
    
root = ThemedTk(theme="adapata")
root.title("Correo")

bntSubir = ttk.Button(root, text="Realizar Backup", command=subir).pack()

if usuario in ['cjalfonso', 'JOCASTRO', 'Yhurtado']:
    tree_frame = ttk.Frame(root, relief="solid")
    tree_frame.pack()
    
    tree = ttk.Treeview(tree_frame)
    tree["columns"] = ("Usuario", "Ultimo Backup", "Estado")
    
    tree.column("#0", width=0, stretch=tk.NO)
    tree.column("Usuario", anchor=tk.W, width=100)
    tree.column("Ultimo Backup", anchor=tk.W, width=100)
    tree.column("Estado", anchor=tk.W, width=100)
    
    tree.heading("#0", text="", anchor=tk.W)
    tree.heading("Usuario", text="Usuario", anchor=tk.W)
    tree.heading("Ultimo Backup", text="Ultimo Backup", anchor=tk.W)
    tree.heading("Estado", text="Estado", anchor=tk.W)
    
    #agrega datos
    mostrar_datos()
    
    tree.pack(padx=10, pady=10)
    
root.mainloop()
    