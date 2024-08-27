import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import customtkinter as ctk
import pyodbc
from tksheet import Sheet
from datetime import datetime
import logging
from dotenv import load_dotenv
import os
from ldap3 import Server, Connection, ALL, NTLM, SUBTREE
from cryptography.fernet import Fernet, InvalidToken

# Configure logging
logging.basicConfig(filename='app.log', level=logging.ERROR)
logging.basicConfig(filename='auth.log', level=logging.INFO)

# Check if .env file exists in the current directory
env_file_path = '.env' if os.path.isfile('.env') else '_internal/.env'

# Load .env file
load_dotenv(env_file_path)

# AD settings
AD_SERVER = os.getenv('AD_SERVER')
AD_DOMAIN = os.getenv('AD_DOMAIN')
AD_USER = os.getenv('AD_USER')
AD_PASSWORD = os.getenv('AD_PASSWORD')
ALLOWED_GROUPS = os.getenv('ALLOWED_GROUPS')
ALLOWED_USERS = os.getenv('ALLOWED_USERS')
# Retrieve the encryption key from the .env file
ENCRYPTION_KEY = os.getenv('ENCRYPTION_KEY')
fernet = Fernet(ENCRYPTION_KEY)

# Ensure the encryption key is loaded
if ENCRYPTION_KEY is None:
    raise ValueError("No encryption key found in environment variables.")


class ScrollableFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)


class MyFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # Create Tksheet widget
        self.sheet = Sheet(self)
        self.sheet.pack(fill="both", expand=True)

        # fases_produccion
        self.fases = [
            "Dispensacion",
            "Pesaje",
            "Fabricacion",
            "Microbiologia",
            "Envasado",
            "Acondicionamiento",
            "Embalaje",
            "Despacho",
            "Reproceso"
        ]
        # plantas de produccion
        self.plantas = ["01", "02"]

        # Configure column headers
        headers = [
            "# OP",
            "# PEDIDO",
            "CODIGO ITEM",
            "DESCRIPCION ITEM",
            "FECHA REQUERIDA",
            "FECHA ENTREGA PLANTA",
            "FECHA ESTIMADO FIN",
            "CANTIDAD PEDIDA",
            "ESTADO OP",
            "COMPANIA",
            "FP_ID",
            "CANTIDAD EN PRODUCCION",
            "FASE DE PRODUCCION",
            "PLANTA",
            "COMENTARIOS/OBSERVACIONES"
        ]
        self.sheet.headers(headers)

        # Enable row selection
        self.sheet.enable_bindings(("single_select", "row_select"))

        # Create a scrollable frame
        self.scrollable_frame = ScrollableFrame(self)
        self.scrollable_frame.pack(fill="both", expand=True)

        # Create filter entry
        self.filter_entry = ctk.CTkEntry(
            self.scrollable_frame, placeholder_text="Filtrar por #OP, #PEDIDO o CODIGO ITEM")
        self.filter_entry.pack(padx=10, pady=10, fill="x")
        self.filter_entry.bind("<Return>", self.filter_data)

        # Create buttons
        self.button_frame = ctk.CTkFrame(self.scrollable_frame)
        self.button_frame.pack(padx=10, pady=10, fill="x")

        self.create_child_button = ctk.CTkButton(
            self.button_frame, text="Crear Registro", command=self.create_child_record)
        self.create_child_button.pack(side="left", padx=5)

        self.edit_child_button = ctk.CTkButton(
            self.button_frame, text="Editar Registro", command=self.edit_child_record)
        self.edit_child_button.pack(side="left", padx=5)

        self.hot_reload_button = ctk.CTkButton(
            self.button_frame, text="Refrescar", command=self.reload_data)
        self.hot_reload_button.pack(side="left", padx=5)
        # Load data from the database
        self.load_data()

    def load_data(self):
        try:
            # Connect to the database
            conn_str = (
                f"DRIVER={os.getenv('DB1_DRIVER')};"
                f"SERVER={os.getenv('DB1_SERVER')};"
                f"DATABASE={os.getenv('DB2_DATABASE')};"
                f"UID={os.getenv('DB1_UID')};"
                f"PWD={os.getenv('DB1_PWD')}"
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            # Fetch data from the database using parameterized query
            query = """
                SELECT
                    pd_ordenproceso.orpconsecutivo AS [# OP]
                    ,pd_ordenproceso.orpconspedi AS [# PEDIDO]
                    ,pd_ordenproceso.orpcodiitem AS [CODIGO ITEM]
                    ,in_items.itedesclarg AS [DESCRIPCION ITEM]
                    ,pd_ordenproceso.orpfecharequ AS [FECHA REQUERIDA]
                    ,pd_ordenproceso.orpfechaentrega AS [FECHA ENTREGA PLANTA]
                    ,pd_ordenproceso.orpfechestifin AS [FECHA ESTIMADO FIN]
                    ,pd_ordenproceso.orpcantrequump AS [CANTIDAD PEDIDA]
                    ,pd_ordenproceso.eobnombre AS [ESTADO OP]
                    ,pd_ordenproceso.orpcompania
                    ,FP_PROGRES.FP_ID
                    ,FP_PROGRES.CANTIDAD_FP AS [CANTIDAD EN PRODUCCION]
                    ,FP_PROGRES.FASE_PODUCC AS [FASE DE PRODUCCION]
                    ,FP_PROGRES.PLANTA AS PLANTA
                    ,FP_PROGRES.COMENTARIES AS [COMENTARIOS/OBSERVACIONES]
                    FROM ssf_genericos.dbo.pd_ordenproceso
                    INNER JOIN ssf_genericos.dbo.in_items
                    ON pd_ordenproceso.orpcodiitem = in_items.itecodigo
                        AND pd_ordenproceso.orpcompania = in_items.itecompania
                    LEFT OUTER JOIN SIIAPP.dbo.FP_PROGRES
                    ON pd_ordenproceso.orpconsecutivo = FP_PROGRES.orpconsecutivo COLLATE Latin1_General_CI_AS
                    LEFT OUTER JOIN SIIAPP.dbo.FP_TIMES
                    ON FP_TIMES.FP_ID = FP_PROGRES.FP_ID
                    WHERE pd_ordenproceso.orpcompania = ?
                    AND in_items.itecompania = ?
                    AND pd_ordenproceso.eobcodigo IN (?, ?, ?)
                    ORDER BY [# OP]
            """
            params = ('01', '01', 'EF', 'PE', 'EE')
            cursor.execute(query, params)
            data = cursor.fetchall()

            # Insert data into the Tksheet
            formatted_data = []
            for row in data:
                parent_row = [
                    str(value) if value is not None else "" for value in row[:10]]
                fp_progres_values = row[10:]

                if any(fp_progres_values):
                    parent_row.extend(
                        str(value) if value is not None else "" for value in fp_progres_values)
                else:
                    # Add empty cells for FP_PROGRES columns
                    parent_row.extend([""] * 5)

                formatted_data.append(parent_row)

            self.original_data = formatted_data
            self.sheet.set_sheet_data(formatted_data)
            # Configure column widths
            self.column_widths = [120, 120, 120, 500, 140,
                                  140, 140, 120, 120, 120, 120, 220, 200, 120, 600]

            for i, width in enumerate(self.column_widths):
                self.sheet.column_width(column=i, width=width)

            # Highlight FP_PROGRES columns
            for i in range(10, 15):
                self.sheet.highlight_columns(
                    columns=[i], bg="lightgray", fg="black")

        except pyodbc.Error as e:
            logging.error(f"An error occurred while loading data: {str(e)}")
            messagebox.showerror(
                "Error", "An error occurred while loading data. Please check the logs for more information.")
        finally:
            # Close the database connection and cursor
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def filter_data(self, event):
        self.column_widths = [120, 120, 120, 500, 140,
                              140, 140, 120, 120, 120, 120, 220, 200, 120, 600]
        op_filter = self.filter_entry.get().lower()
        if op_filter:
            filtered_data = [
                row for row in self.original_data
                if op_filter in str(row[0]).lower() or  # Search by "# OP"
                op_filter in str(row[1]).lower() or  # Search by "# PEDIDO"
                op_filter in str(row[2]).lower()  # Search by "Codigo Item"
            ]
            self.sheet.set_sheet_data(filtered_data)
            for i, width in enumerate(self.column_widths):
                self.sheet.column_width(column=i, width=width)
        else:
            self.sheet.set_sheet_data(self.original_data)
            for i, width in enumerate(self.column_widths):
                self.sheet.column_width(column=i, width=width)

    def create_child_record(self):
        selected_rows = self.sheet.get_selected_rows()
        if selected_rows:
            # Get the first selected row
            selected_row = next(iter(selected_rows))
            row_data = self.sheet.get_row_data(selected_row)
            op_value = row_data[0]  # Assuming '# OP' is at index 0
            it_comp = row_data[9]  # Assuming 'orpcompania' is at index 9

            # Create a new window for entering child record data
            child_window = ctk.CTkToplevel(self)
            child_window.title("Crear Registro de Fase de Produccion")

            # Add input fields for child record data
            cantidad_fp_entry = ctk.CTkEntry(
                child_window, placeholder_text="Cantidad en fase de produccion")
            fase_producc_entry = ctk.CTkComboBox(
                child_window, values=self.fases, state="readonly")
            planta_entry = ctk.CTkComboBox(
                child_window,  values=self.plantas, state="readonly")
            comentarios_entry = ctk.CTkTextbox(
                child_window, height=50, width=200)
            comentarios_entry.configure(
                border_color='blue', border_width=0.5)
            # Grid view
            cantidad_fp_label = ctk.CTkLabel(
                child_window, text="Cantidad en fase de produccion:")
            cantidad_fp_label.grid(row=0, column=0, padx=5, pady=5)
            cantidad_fp_entry.grid(row=0, column=1, padx=5, pady=5)

            fase_producc_label = ctk.CTkLabel(
                child_window, text="Fase de Produccion:")
            fase_producc_label.grid(row=1, column=0, padx=5, pady=5)
            fase_producc_entry.grid(row=1, column=1, padx=5, pady=5)

            planta_label = ctk.CTkLabel(child_window, text="Planta:")
            planta_label.grid(row=2, column=0, padx=5, pady=5)
            planta_entry.grid(row=2, column=1, padx=5, pady=5)

            comentarios_label = ctk.CTkLabel(
                child_window, text="Observasiones/Comentarios:")
            comentarios_label.grid(row=3, column=0, padx=5, pady=5)
            comentarios_entry.grid(row=3, column=1, padx=5, pady=5)

            def save_child_record():
                cantidad_fp = cantidad_fp_entry.get()
                fase_producc = fase_producc_entry.get()
                planta = planta_entry.get()
                comentarios = comentarios_entry.get("0.0", "end")
                if not all([cantidad_fp, fase_producc, planta]):
                    messagebox.showerror(
                        "Error", "Por favor llene todos los campos antes de guardar el registro.")
                    return child_window.destroy()

                # Define cursor variable outside try block
                cursor = None
                try:

                    # Insert the child record into the database using parameterized query
                    conn_str = (
                        f"DRIVER={os.getenv('DB1_DRIVER')};"
                        f"SERVER={os.getenv('DB1_SERVER')};"
                        f"DATABASE={os.getenv('DB1_DATABASE')};"
                        f"UID={os.getenv('DB1_UID')};"
                        f"PWD={os.getenv('DB1_PWD')}"
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    insert_query = """
                        INSERT INTO FP_PROGRES (orpconsecutivo, orpcompania, CANTIDAD_FP, FASE_PODUCC, PLANTA, COMENTARIES)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """
                    params = (op_value, it_comp, cantidad_fp,
                              fase_producc, planta, comentarios)
                    cursor.execute(insert_query, params)
                    conn.commit()

                    # Get the last inserted FP_ID
                    cursor.execute("SELECT @@IDENTITY")
                    fp_id = cursor.fetchone()[0]

                    # Insert data into FP_TIMES table for the corresponding phase with current datetime
                    if fase_producc == "Despacho":
                        insert_times_query = f"""
                            INSERT INTO FP_TIMES (FP_ID, {fase_producc}_ST, {fase_producc}_ET)
                            VALUES (?, ?, ?)
                        """
                        current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        cursor.execute(
                            insert_times_query, (fp_id, current_datetime, current_datetime))
                    else:
                        insert_times_query = f"""
                            INSERT INTO FP_TIMES (FP_ID, {fase_producc}_ST)
                            VALUES (?, ?)
                        """
                        current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        cursor.execute(insert_times_query,
                                       (fp_id, current_datetime))
                    conn.commit()

                except pyodbc.Error as e:
                    logging.error(
                        f"An error occurred while saving child record: {str(e)}")
                    messagebox.showerror(
                        "Error", "An error occurred while saving the child record. Please check the logs for more information.")
                finally:
                    # Close the database connection and cursor
                    if cursor:
                        cursor.close()
                    if conn:
                        conn.close()

                child_window.destroy()
                self.reload_data()

            save_button = ctk.CTkButton(
                child_window, text="Guardar", command=save_child_record)
            save_button.grid(row=4, column=0, columnspan=2, pady=10)
        else:
            messagebox.showinfo(
                "Sin seleccion", "Porfavor eliga una fila para Crear un registro")

    def edit_child_record(self):
        selected_rows = self.sheet.get_selected_rows()
        if selected_rows:
            # Get the first selected row
            selected_row = next(iter(selected_rows))
            row_data = self.sheet.get_row_data(selected_row)
            op_value = row_data[0]  # Assuming '# OP' is at index 0
            fp_id = row_data[10]  # Assuming 'FP_ID' is at index 10

            if fp_id == "":
                messagebox.showerror(
                    "Error", "No se puede editar el registro porque no se ha creado.")
                return

            # Create a new window for editing child record data
            edit_window = ctk.CTkToplevel(self)
            edit_window.title("Editar registro de fase de produccion")

            # Add input fields for child record data
            cantidad_fp_entry = ctk.CTkEntry(
                edit_window, placeholder_text="Cantidad en fase de produccion")
            # Pre-fill with existing data
            cantidad_fp_entry.insert(0, row_data[11])
            fase_producc_entry = ctk.CTkComboBox(
                edit_window, values=self.fases, state="readonly")
            # Pre-fill with existing data
            fase_producc_entry.set(row_data[12])
            planta_entry = ctk.CTkComboBox(
                edit_window,  values=self.plantas, state="readonly")
            planta_entry.set(row_data[13])
            # Pre-fill with existing data
            comentarios_entry = ctk.CTkTextbox(
                edit_window, height=50, width=200)
            # Pre-fill with existing data
            comentarios_entry.insert("0.0", row_data[14])
            comentarios_entry.configure(
                border_color='blue', border_width=0.5)
            # Grid view
            cantidad_fp_label = ctk.CTkLabel(
                edit_window, text="Cantidad en fase de produccion:")
            cantidad_fp_label.grid(row=0, column=0, padx=5, pady=5)
            cantidad_fp_entry.grid(row=0, column=1, padx=5, pady=5)

            fase_producc_label = ctk.CTkLabel(
                edit_window, text="Fase de Produccion:")
            fase_producc_label.grid(row=1, column=0, padx=5, pady=5)
            fase_producc_entry.grid(row=1, column=1, padx=5, pady=5)

            planta_label = ctk.CTkLabel(edit_window, text="Planta:")
            planta_label.grid(row=2, column=0, padx=5, pady=5)
            planta_entry.grid(row=2, column=1, padx=5, pady=5)

            comentarios_label = ctk.CTkLabel(
                edit_window, text="Observaciones/Comentarios:")
            comentarios_label.grid(row=3, column=0, padx=5, pady=5)
            comentarios_entry.grid(row=3, column=1, padx=5, pady=5)

            def save_edited_child_record():
                try:
                    cantidad_fp = cantidad_fp_entry.get()
                    fase_producc = fase_producc_entry.get()
                    planta = planta_entry.get()
                    comentarios = comentarios_entry.get("0.0", "end")

                    # Update the child record in the database using parameterized query
                    conn_str = (
                        f"DRIVER={os.getenv('DB1_DRIVER')};"
                        f"SERVER={os.getenv('DB1_SERVER')};"
                        f"DATABASE={os.getenv('DB1_DATABASE')};"
                        f"UID={os.getenv('DB1_UID')};"
                        f"PWD={os.getenv('DB1_PWD')}"
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    select_query = """
                        SELECT FASE_PODUCC
                        FROM FP_PROGRES
                        WHERE FP_ID = ?
                    """
                    cursor.execute(select_query, (fp_id,))
                    existing_data = cursor.fetchone()
                    if existing_data:
                        prev_fase = existing_data[0]
                    else:
                        prev_fase = None

                    update_query = """
                        UPDATE FP_PROGRES
                        SET CANTIDAD_FP = ?, FASE_PODUCC = ?, PLANTA = ?, COMENTARIES = ?
                        WHERE FP_ID = ?
                    """
                    params = (cantidad_fp, fase_producc,
                              planta, comentarios, fp_id)
                    cursor.execute(update_query, params)
                    conn.commit()

                    # Update the corresponding phase start and end times in FP_TIMES table with current datetime
                    if fase_producc == "Despacho":
                        update_times_query = f"""
                            UPDATE FP_TIMES
                            SET {fase_producc}_ST = CASE WHEN {fase_producc}_ST IS NULL THEN ? ELSE {fase_producc}_ST END,
                                {fase_producc}_ET = ?,
                                {prev_fase}_ET = CASE WHEN {prev_fase}_ST IS NOT NULL THEN ? ELSE {prev_fase}_ET END
                            WHERE FP_ID = ?
                        """
                        current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        cursor.execute(update_times_query,
                                       (current_datetime, current_datetime, current_datetime, fp_id))
                    else:
                        update_times_query = f"""
                            UPDATE FP_TIMES
                            SET {fase_producc}_ST = CASE WHEN {fase_producc}_ST IS NULL THEN ? ELSE {fase_producc}_ST END,
                                {prev_fase}_ET = CASE WHEN {prev_fase}_ST IS NOT NULL THEN ? ELSE {prev_fase}_ET END
                            WHERE FP_ID = ?
                        """
                        current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        cursor.execute(update_times_query,
                                       (current_datetime, current_datetime, fp_id))
                    conn.commit()

                except pyodbc.Error as e:
                    logging.error(
                        f"An error occurred while updating child record: {str(e)}")
                    messagebox.showerror(
                        "Error", "An error occurred while updating the child record. Please check the logs for more information.")
                finally:
                    # Close the database connection and cursor
                    if cursor:
                        cursor.close()
                    if conn:
                        conn.close()
                edit_window.destroy()
                self.reload_data()

            save_button = ctk.CTkButton(
                edit_window, text="Guardar Cambios", command=save_edited_child_record)
            save_button.grid(row=4, column=0, columnspan=2, pady=10)
        else:
            messagebox.showinfo(
                "Sin seleccion", "Porfavor eliga una fila para editar un registro")

    def reload_data(self):
        # Clear existing data
        self.sheet.set_sheet_data([])

        # Load updated data from the database
        self.load_data()


class LoginFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.username_entry = ctk.CTkEntry(
            self, placeholder_text="Nombre de usuario")
        self.username_entry.pack(pady=10)
        self.password_entry = ctk.CTkEntry(
            self, placeholder_text="Contraseña", show="*")
        self.password_entry.pack(pady=10)
        self.remember_var = tk.BooleanVar()  # Variable to track the checkbox state
        self.remember_checkbox = ctk.CTkCheckBox(
            self, text="Recordar mis credenciales", variable=self.remember_var)
        self.remember_checkbox.pack(pady=5)
        self.login_button = ctk.CTkButton(
            self, text="Login", command=self.authenticate)
        self.login_button.pack(pady=10)

        # Load saved credentials if available
        self.load_credentials()

    def save_credentials(self):
        if self.remember_var.get():
            encrypted_username = fernet.encrypt(
                self.username_entry.get().encode())
            encrypted_password = fernet.encrypt(
                self.password_entry.get().encode())
            with open("credentials.txt", "wb") as f:
                f.write(encrypted_username + b"," + encrypted_password)

    def load_credentials(self):
        try:
            with open("credentials.txt", "rb") as f:
                data = f.read()
                encrypted_username, encrypted_password = data.split(b",")
                self.username = fernet.decrypt(encrypted_username).decode()
                self.password = fernet.decrypt(encrypted_password).decode()
                self.username_entry.insert(0, self.username)
                self.password_entry.insert(0, self.password)
        except FileNotFoundError:
            pass
        except (ValueError, InvalidToken):
            messagebox.showerror(
                "Error", "Unable to decrypt credentials. Please enter the correct password.")

    def authenticate(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if authenticate_user(username, password):
            messagebox.showinfo("Login Exitoso", "Bienvenido!")
            self.save_credentials()  # Save credentials before showing the app frame
            self.master.show_app_frame()
        else:
            messagebox.showerror(
                "Login Fallido", "Credenciales invalidas o acceso denegado.")


def authenticate_user(username, password):
    server = Server(os.getenv('AD_SERVER'), get_info=ALL)
    user = f'{os.getenv("AD_DOMAIN")}\\{username}'

    try:
        conn = Connection(server, user=user, password=password,
                          authentication='NTLM', auto_bind=True)
        logging.info(f"LDAP bind successful for {username}.")

        # Check if user is in allowed users
        allowed_users = os.getenv('ALLOWED_USERS').split(',')
        if username in allowed_users:
            return True

        # Search base is set to the root of the domain
        search_base = f'DC={AD_DOMAIN.replace(".", ",DC=")}'

        conn.search(
            search_base,
            f'(sAMAccountName={username})',
            attributes=['memberOf'],
            search_scope=SUBTREE
        )

        if not conn.entries:
            logging.warning(f"User {username} not found in LDAP search.")
            return False

        user_groups = [entry.memberOf.values if isinstance(entry.memberOf, list) else [
            entry.memberOf] for entry in conn.entries]
        user_groups = [item for sublist in user_groups for item in sublist]

        allowed_groups = os.getenv('ALLOWED_GROUPS').split(',')

        for group in allowed_groups:
            if any(group in str(user_group) for user_group in user_groups):
                return True

    except Exception as e:
        logging.error(f"LDAP error for {username}: {e}")

    logging.warning(f"Access denied for {username}.")
    return False


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("1000x600")
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.login_frame = LoginFrame(master=self)
        self.login_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

    def show_app_frame(self):
        self.login_frame.destroy()
        self.geometry("1000x600")
        self.my_frame = MyFrame(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")


app = App()
app.title("SIIAPP FASES PRODUCCION")
app.mainloop()
