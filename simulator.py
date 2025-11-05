
import c104
import time
import random
import pandas as pd
import datetime
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk, filedialog
import threading 
import subprocess
from pyModbusTCP.client import ModbusClient
from pyModbusTCP.server import ModbusServer
import struct

iec104_type_ids = {
    1: c104.Type.M_SP_NA_1,
    13: c104.Type.M_ME_NC_1,
    30: c104.Type.M_SP_TB_1,
    36: c104.Type.M_ME_TF_1,
    45: c104.Type.C_SC_NA_1,
    50: c104.Type.C_SE_NC_1,
}

def float_to_registers(value):
    packed = struct.pack('<f', value)
    return list(struct.unpack('<HH', packed))

def unsigned_integer_to_register(value): 
    if value < 0 or value > 0xFFFFFFFF:
        raise ValueError("Value must be an unsigned 32-bit integer (0 to 4294967295).")
    packed = struct.pack('<I', value)
    return list(struct.unpack('<HH', packed))

def signed_integer_to_register(value):
    if value < -0x80000000 or value > 0x7FFFFFFF:
        raise ValueError("Value must be a signed 32-bit integer (-2147483648 to 2147483647).")
    packed = struct.pack('<i', value)
    return list(struct.unpack('<HH', packed))

def unsigned_16bit_to_register(value):
    if value < 0 or value > 0xFFFF:
        raise ValueError("Value must be an unsigned 16-bit integer (0 to 65535).")
    packed = struct.pack('<H', value)
    return list(struct.unpack('<H', packed))

def signed_16bit_to_register(value):
    if value < -0x8000 or value > 0x7FFF:
        raise ValueError("Value must be a signed 16-bit integer (-32768 to 32767).")
    packed = struct.pack('<h', value)
    return list(struct.unpack('<H', packed))

def swap(my_list):
    my_list[0],my_list[1] = my_list[1], my_list[0]
    return my_list

#***************Functions for big endian float to register************************** 

def float_to_registers_be(value):
    packed = struct.pack('>f', value)
    return list(struct.unpack('>HH', packed))

def unsigned_integer_to_register_be(value):
    if value < 0 or value > 0xFFFFFFFF:
        raise ValueError("Value must be an unsigned 32-bit integer (0 to 4294967295).")
    packed = struct.pack('>I', value)
    return list(struct.unpack('>HH', packed))

def signed_integer_to_register_be(value):
    if value < -0x80000000 or value > 0x7FFFFFFF:
        raise ValueError("Value must be a signed 32-bit integer (-2147483648 to 2147483647).")
    packed = struct.pack('>i', value)
    return list(struct.unpack('>HH', packed))

def unsigned_16bit_to_register_be(value):
    if value < 0 or value > 0xFFFF:
        raise ValueError("Value must be an unsigned 16-bit integer (0 to 65535).")
    packed = struct.pack('>H', value)
    return list(struct.unpack('>H', packed))

def signed_16bit_to_register_be(value):
    if value < -0x8000 or value > 0x7FFF:
        raise ValueError("Value must be a signed 16-bit integer (-32768 to 32767).")
    packed = struct.pack('>h', value)
    return list(struct.unpack('>H', packed))

#***************Functions for little endian register to float**************************

def registers_to_float(registers):
    packed = struct.pack('<HH', *registers)
    return struct.unpack('<f', packed)[0]

def registers_to_unsigned_integer(registers):
    packed = struct.pack('<HH', *registers)
    return struct.unpack('<I', packed)[0]

def registers_to_signed_integer(registers):
    packed = struct.pack('<HH', *registers)
    return struct.unpack('<i', packed)[0]

def registers_to_unsigned_16bit(registers):
    packed = struct.pack('<H', registers[0])
    return struct.unpack('<H', packed)[0]

def registers_to_signed_16bit(registers):
    packed = struct.pack('<H', registers[0])
    return struct.unpack('<h', packed)[0]

#***************Functions for big endian register to float**************************

def registers_to_float_be(registers):
    packed = struct.pack('>HH', *registers)
    return struct.unpack('>f', packed)[0]

def registers_to_unsigned_integer_be(registers):
    packed = struct.pack('>HH', *registers)
    return struct.unpack('>I', packed)[0]

def registers_to_signed_integer_be(registers):
    packed = struct.pack('>HH', *registers)
    return struct.unpack('>i', packed)[0]

def registers_to_unsigned_16bit_be(registers):
    packed = struct.pack('>H', registers[0])
    return struct.unpack('>H', packed)[0]

def registers_to_signed_16bit_be(registers):
    packed = struct.pack('>H', registers[0])
    return struct.unpack('>h', packed)[0]

class IEC104SlaveSingle:
    def __init__(self, master):
        self.master = master
        self.current_dialog = None

        master.title("IEC 104 Slave Simulator")
        master.geometry("700x720+20+20")

        self.server = None
        self.xls = None
        self.log_data = []
        self.signal_data = {}
        self.current_signal_index = 0

        self.label = tk.Label(master, text=" 104 Single Device Simulator ", font=("Arial", 14))
        self.label.pack(pady=10)

        self.log_frame = tk.Frame(self.master)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(self.log_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(self.log_frame, wrap='word', state='disabled', height=15, yscrollcommand=self.scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar.config(command=self.log_text.yview)

        self.upload_button = tk.Button(master, text="Upload Excel File", command=self.upload_excel)
        self.upload_button.pack(pady=10)

        self.file_name_label = tk.Label(master, text="", font=("Arial", 10))
        self.file_name_label.pack(pady=5)

        self.sheet_frame = tk.Frame(master)
        self.sheet_frame.pack(padx=20, pady=10)

        self.sheet_label = tk.Label(self.sheet_frame, text="Select Sheet:", font=("Arial", 10))
        self.sheet_label.pack(side=tk.LEFT)

        self.sheet_combo = ttk.Combobox(self.sheet_frame, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=(10, 0))

        self.load_ip_button = tk.Button(master, text="Load IPs", command=self.load_ips, state="disabled")
        self.load_ip_button.pack(pady=10)

        self.ip_frame = tk.Frame(master)
        self.ip_frame.pack(padx=20, pady=10)

        self.ip_label = tk.Label(self.ip_frame, text="Select IP Address :", font=("Arial", 10))
        self.ip_label.pack(side=tk.LEFT)

        self.ip_combo = ttk.Combobox(self.ip_frame, state="readonly")
        self.ip_combo.pack(side=tk.LEFT, padx=(10, 0))

        self.port_label = tk.Label(self.ip_frame, text="Port :", font=("Arial", 10))
        self.port_label.pack(side=tk.LEFT, padx=(20, 0))

        self.port_entry = tk.Entry(self.ip_frame, font=("Arial", 10), width = 7, relief="sunken")
        self.port_entry.insert(0,2404)
        self.port_entry.pack(side=tk.LEFT, padx=(10, 0))

        self.asdu_label = tk.Label(self.ip_frame, text="ASDU : ", font=("Arial", 10))
        self.asdu_label.pack(side=tk.LEFT, padx=(20, 0))

        self.asdu_entry = tk.Entry(self.ip_frame, font=("Arial", 10), width = 7, relief="sunken")
        self.asdu_entry.insert(0,1)
        self.asdu_entry.pack(side=tk.LEFT, padx=(10, 0))

        self.connect_button = tk.Button(master, text="Connect", command=self.connect)
        self.connect_button.pack(pady=10)

        self.stop_button = tk.Button(master, text="STOP", fg="red", command=self.stop_server)
        self.stop_button.pack(pady=10)

        self.report_frame = tk.Frame(master)
        self.report_frame.pack(padx=20, pady=10)

        self.report_button_csv = tk.Button(self.report_frame, text="Generate Report CSV ", state="disabled", command = self.generate_report_csv)
        self.report_button_csv.pack(side = tk.LEFT)
        
        self.report_button_text = tk.Button(self.report_frame, text="Generate Report TEXT ", state="disabled", command = self.generate_report_txt)
        self.report_button_text.pack(side=tk.LEFT, padx=(20, 0))

        self.status_label = tk.Label(master, text="", fg="green", font=("Arial", 10))
        self.status_label.pack(pady=10)

        self.master.protocol("WM_DELETE_WINDOW", self.close_simulator)

    def log(self, message):
        """Logs a message to the log text area."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.master.update()

        self.log_data.append({"Timestamp": datetime.datetime.now(), "Message": message})

    def generate_report_txt(self):
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(report_filename, "w") as report_file:
            report_file.write(f"IEC 104 Simulator Report\n")
            report_file.write(f"Generated on: {datetime.datetime.now()}\n\n")
            report_file.write(self.log_text.get("1.0", tk.END))

        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    def generate_report_csv(self):
        df = pd.DataFrame(self.log_data)
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(report_filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name="Log", index=False)

            # Customize Excel sheet appearance (optional)
            workbook = writer.book
            worksheet = writer.sheets["Log"]
            worksheet.set_column('A:B', 20)  # Adjust column width as needed
        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    def upload_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.xls = pd.ExcelFile(file_path)
            self.sheet_combo['values'] = self.xls.sheet_names
            self.sheet_combo.current(0)
            self.load_ip_button['state'] = "normal"

            file_name = file_path.split("/")[-1]
            self.file_name_label.config(text=f"Uploaded File: {file_name}")

    def load_ips(self):
        selected_sheet = self.sheet_combo.get()
        if not selected_sheet:
            messagebox.showwarning("Warning", "Please select a sheet first.")
            return

        df = pd.read_excel(self.xls, sheet_name=selected_sheet)
        unique_ips = df['IP Address'].dropna().unique()
        ip_list = [str(ip).strip() for ip in unique_ips]

        self.ip_combo['values'] = ip_list
        if ip_list:
            self.ip_combo.current(0)

    def connect(self):
        selected_sheet = self.sheet_combo.get()
        port = self.port_entry.get()
        asdu = self.asdu_entry.get()
        ip_address = self.ip_combo.get()
        if not selected_sheet or not ip_address:
            messagebox.showwarning("Warning", "Please select a sheet and IP address.")
            return
        self.reset_server()
        self.log(f"Setting up server for {ip_address}")
        self.server = c104.Server(ip=ip_address, port=int(port))
        self.station = self.server.add_station(common_address=int(asdu))
        self.server.start()

        self.waiting_dots = 0
        self.log(f"Waiting for connection to {ip_address} ")
        self.check_connection()

    def check_connection(self):
        ip_address = self.ip_combo.get()
        if self.server.has_active_connections:
            self.log(f"{self.ip_combo.get()} Connected ")
            self.choose_processing_mode()
        else:
            '''self.waiting_dots = (self.waiting_dots + 1) % 4
            dots = "." * self.waiting_dots
            self.log(f"Waiting for connection{dots}")
            self.master.after(500, self.check_connection)
            '''
            self.log(f"Waiting for connection to {ip_address}")
            time.sleep(2)
            self.check_connection()

    # To display something in label
    def update_status(self, message):
        self.status_label.config(text=message)

    def choose_processing_mode(self):
        """Show a radio button dialog for choosing processing mode."""
        mode_dialog = tk.Toplevel(self.master)
        mode_dialog.title("Select Processing Mode")

        mode_dialog.geometry("300x200+100+300")

        self.selected_mode = tk.StringVar(value="one")

        one_radio = tk.Radiobutton(mode_dialog, text="One-by-One", variable=self.selected_mode, value="one", font=("Arial", 12))
        one_radio.pack(pady=10)

        all_radio = tk.Radiobutton(mode_dialog, text="All-at-Once", variable=self.selected_mode, value="all", font=("Arial", 12))
        all_radio.pack(pady=10)

        specific_ioa_radio = tk.Radiobutton(mode_dialog, text="Specific IOA", variable=self.selected_mode, value="specific_ioa", font=("Arial", 12))
        specific_ioa_radio.pack(pady=10)

        confirm_button = tk.Button(mode_dialog, text="Confirm", command=lambda: self.process_selected_mode(mode_dialog))
        confirm_button.pack(pady=10)

        mode_dialog.wait_window()  # Wait for the dialog to close

    def process_selected_mode(self, mode_dialog):
        """Process the selected mode and close the dialog."""
        mode_dialog.destroy()  
        self.mode = self.selected_mode.get()  

        # Redirect to the corresponding processing logic
        if self.mode in  ["specific_ioa", "one"]:
            self.process_signals_one_by_one()
        elif self.mode== "all":
            self.process_signals_all_at_once()
        else:
            messagebox.showerror("Error", "Invalid mode selected")

    def process_signals_one_by_one(self):
        selected_sheet = self.sheet_combo.get()
        df = pd.read_excel(self.xls, sheet_name=selected_sheet)
        grouped = df.groupby('IP Address')

        for ip_address, group in grouped:
            if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either server stopped or client disconnected")
                break
            
            self.signal_data[ip_address] = list(group.iterrows()) # Store for navigation
            self.current_signal_index = 0

            if str(ip_address).strip() == self.ip_combo.get().strip():
                for _, row in self.signal_data[ip_address]:
                    ioa = row['IOA']
                    name = row['Object Text']
                    type_id = row['Type ID']
                    point_type = iec104_type_ids[row['Type ID']]
                    if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                        self.log("Processing stopped either server stopped or client disconnected")
                        break
                    if type_id in [1,13,30,36,45,50]:
                        point_type = iec104_type_ids[row['Type ID']]
                        point = self.station.add_point(io_address=int(row['IOA']), type=point_type, report_ms=0)
                    else:
                        self.log(f"Invalid Type ID {row['Type ID']} for IOA {row['IOA']}")
        
        self.log("All points created, Ready for update ")
        if self.mode== "one":
            self.update_signals( ip_address)
        elif self.mode == "specific_ioa":
            self.process_specific_ioa()
        self.log("         Processing completed            ")
        self.master.update_idletasks()
        time.sleep(5)
        self.stop_server()

    def process_signals_all_at_once(self):
        selected_sheet = self.sheet_combo.get()
        df = pd.read_excel(self.xls, sheet_name=selected_sheet)

        if 'UpdatedValue' in df.columns :
            df = df.drop(columns=['UpdatedValue'])
            df['UpdatedValue'] = df['value']
        else:
            df['UpdatedValue'] = df['value']

        grouped = df.groupby('IP Address')
        
        for ip_address, group in grouped:
            if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either server stopped or client disconnected")
                break

            points={}

            if str(ip_address).strip() == self.ip_combo.get().strip():
                for _, row in group.iterrows():
                    if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                        self.log("Processing stopped either server stopped or client disconnected")
                        break
                    if row['Type ID'] in [1,13,30,36,45,50]:
                        point_type = iec104_type_ids[row['Type ID']]
                        point = self.station.add_point(io_address=int(row['IOA']), type=point_type, report_ms=0)

                self.log("All points created, Ready for update ")

                self.log(f"Starting Continuous Value Updates for {ip_address}")

                while self.server.is_running:
                    for _, row in group.iterrows():
                        if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                            self.log("Processing stopped either server stopped or client disconnected")
                            break
                        ioa = int(row['IOA'])
                        type_id = row['Type ID']
                        name = row['Object Text']
                        point = self.station.get_point(io_address = ioa)
                        value = df.loc[row.name, 'UpdatedValue'] if pd.notna(df.loc[row.name, 'UpdatedValue']) else 0

                        # Update the value for the point based on its type
                        if type_id == 45 :
                            val = point.value
                            value = bool(val)
                            self.log(f"Received point IOA : {ioa} : {name} : {value}")
                            df.loc[row.name, 'UpdatedValue'] = value
                        elif type_id == 50:
                            val = point.value
                            value = round(val,5)
                            self.log(f"Received point IOA : {ioa} : {name} : {value}")
                            df.loc[row.name, 'UpdatedValue'] = float(value)
                        elif type_id in [1,30]:
                            point.report_ms = 1000
                            point.value = (bool(value))
                            self.log(f"Set point IOA : {ioa} : {name} : {bool(value)}")
                            value = not bool(value)
                            df.loc[row.name, 'UpdatedValue'] = int(value)
                        elif type_id in [13,36]:
                            point.report_ms = 1000
                            point.value = (float(value))
                            self.log(f"Set point IOA : {ioa} : {name} : {value} ")
                            value = random.randint(10,100)
                            df.loc[row.name, 'UpdatedValue'] = value
                        else:
                            self.log(f"Invalid type id {type_id} for IOA {ioa}")
                        
                        time.sleep(1.2)
                        if type_id in [1,13,30,36]:
                            point.report_ms = 0

                    if not self.server.is_running or not self.server.has_active_connections:
                        break
                    df.to_excel(self.xls, sheet_name=selected_sheet, index=False)
                    self.log("Saved updated values to Excel.")
                    self.log ("Next set of Update is starting........")
                    time.sleep(5)  # Sleep for 4 seconds before the next update loop

                if not self.server.is_running or not self.server.has_active_connections :
                    self.log("Stopped updates for All-at-Once mode.")
                    break

    def update_signals(self, ip_address):
        for _, row in self.signal_data[ip_address]:
            ioa = row['IOA']
            name = row['Object Text']
            type_id = row['Type ID']
            point = self.station.get_point(io_address = ioa)
            if type_id in [45,50]:
                if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either server stopped or client disconnected")
                self.get_command_signal_dialog(point, name, type_id, ip_address)
            elif type_id in [1,30]:
                if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either server stopped or client disconnected")
                self.show_binary_input_dialog(point, name, type_id, ip_address)
            elif type_id in [13,36]:
                if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either server stopped or client disconnected")
                self.show_numeric_input_dialog(point, name, type_id, ip_address)
            else:
                self.log(f"Invalid type id {type_id} for IOA {ioa}")
                continue
            
            self.current_signal_index +=1

    def process_specific_ioa(self):
        ioa_input = simpledialog.askinteger("Input", "Please enter the starting IOA:", minvalue=0)
        if ioa_input is None:
            self.choose_processing_mode()
            return  # User cancelled the input
        
        selected_ip = self.ip_combo.get()
        if not selected_ip:
            messagebox.showwarning("Warning", "Please select a valid IP address.")
            return

        df = pd.read_excel(self.xls, sheet_name=self.sheet_combo.get())
        group = df[df['IP Address'].astype(str).str.strip() == selected_ip]
        if group.empty:
            messagebox.showinfo("Info", f"No signals found for IP : {selected_ip}. and IOA : {ioa_input}")
            return

        result = group[group['IOA'] == ioa_input]
        if result.empty:
            messagebox.showinfo("Info", f"No signal found for IOA: {ioa_input}.")
            self.choose_processing_mode()
             
        row = result.iloc[0]
        name = row['Object Text']
        type_id = row['Type ID']
        ip_address = row['IP Address']
        if type_id in [1,13,30,36,45,50]:
            point = self.station.get_point(io_address = ioa_input)
            if type_id in [45,50] :
                if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either server stopped or client disconnected")
                self.get_command_signal_dialog(point, name, type_id, ip_address)
            elif type_id in [1, 30]: # Binary input signals
                if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either server stopped or client disconnected")
                self.show_binary_input_dialog(point, name, type_id, ip_address)
            elif type_id in [13, 36]:  # Numeric input signals                        
                if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either server stopped or client disconnected")
                self.show_numeric_input_dialog(point, name,type_id, ip_address)
        
        time.sleep(2)
        if self.server.is_running and self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
            self.choose_processing_mode()

    def show_previous_signal(self, point, type_id, ip_address):
        if self.mode == "specific_ioa":
            self.current_dialog.destroy()
        else:
            self.current_signal_index -= 1
            if self.current_signal_index < 0:
                self.current_signal_index = 0
                return # at the beginning
            self.current_dialog.destroy()
            self.process_current_signal(ip_address) # new function

    def show_next_signal(self, point, type_id, ip_address):
        if self.mode == "specific_ioa":
            self.current_dialog.destroy()
        else:
            self.current_signal_index += 1
            if self.current_signal_index >= len(self.signal_data[ip_address]):
                self.current_signal_index = len(self.signal_data[ip_address]) -1
                return # at the end
            self.current_dialog.destroy()
            self.process_current_signal(ip_address) # new function

    def process_current_signal(self, ip_address):
        _, row = self.signal_data[ip_address][self.current_signal_index]
        ioa = row['IOA']
        name = row['Object Text']
        type_id = row['Type ID']
        point = self.station.get_point(io_address = ioa)
        if type_id in [45,50]:
            if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either server stopped or client disconnected")
            self.get_command_signal_dialog(point, name, type_id, ip_address)
        elif type_id in [1,30]:
            if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either server stopped or client disconnected")
            self.show_binary_input_dialog(point, name, type_id, ip_address)
        elif type_id in [13,36]:
            if not self.server.is_running or not self.server.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either server stopped or client disconnected")
            self.show_numeric_input_dialog(point, name, type_id, ip_address)
        else:
            self.log(f"Invalid type id {type_id} for IOA {ioa}")
            self.current_signal_index += 1
            self.process_current_signal(ip_address)

    def get_command_signal_dialog(self, point, name, type_id, ip_address):
        if not self.server.is_running:
            return
        command_dialog = tk.Toplevel(self.master)
        self.current_dialog = command_dialog
        command_dialog.title("Command Signal")

        command_dialog.geometry("500x200+100+390")
        label = tk.Label(command_dialog, text=f"Signal: {name}\nIOA: {point.io_address}", font=("Arial", 12))
        label.pack(pady=10)

        self.display_text = tk.Text(command_dialog, height=2, width=25)
        self.display_text.pack(pady=5)

        fetch_button = tk.Button(command_dialog, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_command_signal(point, name, type_id,))
        fetch_button.pack(pady = 5)

        prev_button = tk.Button(command_dialog, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(point, type_id, ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=15)

        next_button = tk.Button(command_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(point, type_id, ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=15)

        command_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        command_dialog.wait_window()

    def get_command_signal(self, point, name, type_id):
        val = point.value
        if type_id == 45 :
            value = bool(val)
        else:
            value = round(val,5)
        self.display_text.delete('1.0',tk.END)
        self.display_text.insert('1.0', f"{value}")
        self.log(f"Received point IOA : {point.io_address} : {name} : {value}")
        time.sleep(2)

    def show_binary_input_dialog(self, point, name, type_id, ip_address):
        if not self.server.is_running:
            return
        binary_dialog = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog
        binary_dialog.title("Binary Input")

        '''binary_dialog.geometry("500x200+100+390")
        label = tk.Label(binary_dialog, text=f"Signal: {name}\nIOA: {point.io_address}", font=("Arial", 12))
        label.pack(pady=10)'''

        label = tk.Label(binary_dialog, text=f"Signal: {name}\nIOA: {point.io_address}", font=("Arial", 12))
        label.update_idletasks()  # Force label to calculate its size

        min_width = 500
        min_height = 200

        dialog_width = max(label.winfo_reqwidth() + 40, min_width)  # Use max to ensure minimum width
        dialog_height = min_height

        binary_dialog.geometry(f"{dialog_width}x{dialog_height}+100+390")  # Use f-string for dynamic width
        label.pack(pady=10)

        self.onoff_frame = tk.Frame(binary_dialog)
        self.onoff_frame.pack(padx=20, pady=10)

        on_button = tk.Button(self.onoff_frame, text="  On  ", fg="green", font=("Arial", 10),command=lambda: self.set_point_value(point, name, bool(1)))
        on_button.pack(side=tk.LEFT, padx = 10)

        off_button = tk.Button(self.onoff_frame, text="  Off  ", fg="red", font=("Arial", 10),command=lambda: self.set_point_value(point, name, bool(0)))
        off_button.pack(side= tk.RIGHT, padx = 10)

        prev_button = tk.Button(binary_dialog, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(point, type_id, ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=30)

        next_button = tk.Button(binary_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(point, type_id, ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=30)

        binary_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog.wait_window()

    def show_numeric_input_dialog(self, point, name, type_id, ip_address):
        if not self.server.is_running:
            return
        numeric_dialog = tk.Toplevel(self.master)
        self.current_dialog = numeric_dialog
        numeric_dialog.title("Analog Input")

        '''numeric_dialog.geometry("500x200+100+390")
        label = tk.Label(numeric_dialog, text=f"Signal: {name}\nIOA: {point.io_address}", font=("Arial", 12))
        label.pack(pady=10)'''

        label = tk.Label(numeric_dialog, text=f"Signal: {name}\nIOA: {point.io_address}", font=("Arial", 12))
        label.update_idletasks()

        min_width = 500
        min_height = 400

        dialog_width = max(label.winfo_reqwidth() + 40, min_width)  # Use max to ensure minimum width
        dialog_height = min_height

        numeric_dialog.geometry(f"{dialog_width}x{dialog_height}+100+390")  # Use f-string for dynamic width
        label.pack(pady=10)

        value_entry = tk.Entry(numeric_dialog, font=("Arial", 12))
        value_entry.pack(pady=5)

        confirm_button = tk.Button(numeric_dialog, text="Confirm", command=lambda: self.set_numeric_value(point, name, value_entry))
        confirm_button.pack(pady=5)

        prev_button = tk.Button(numeric_dialog, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(point, type_id, ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=22)

        next_button = tk.Button(numeric_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(point, type_id, ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=22)

        numeric_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        numeric_dialog.wait_window()

    def set_numeric_value(self, point, name, value_entry):
        point.report_ms = 1000
        value = float(value_entry.get())
        point.value = value
        self.log(f"Set point IOA : {point.io_address} : {name} : {value}")
        time.sleep(0.8)
        point.report_ms = 0

    def set_point_value(self, point, name, value):
        point.report_ms = 1000
        point.value = value
        self.log(f"Set point IOA : {point.io_address} : {name} : {value}")
        time.sleep(0.8)
        point.report_ms = 0
   
    def dialog_closed(self):
        """Clear the dialog reference when it is manually closed."""
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None

    def close_simulator(self):
        """Stop all servers and close the main window."""
        if messagebox.askokcancel("Quit", "Do you want to stop all servers and close the simulator?"):
            self.stop_server()
            self.master.destroy()

    def stop_server(self):
        """Stop the IEC 104 server."""
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None  # Clear the reference

        if self.server:
            self.server.stop()
            self.log("Server stopped .")
            print("Server stopped.")
            self.report_button_csv['state'] = "normal"
            self.report_button_text['state'] = "normal"

    def reset_server(self):
        self.stop_server()
        self.server = None 
        self.station = None  

class IEC104SlaveMultiple:
    def __init__(self, master):
        self.master = master
        self.current_dialog = None
        self.server_running = True

        self.master.title("IEC 104 Slave Simulator")
        self.master.geometry("700x720+20+20")

        self.servers = []
        self.file_paths = []
        self.log_data = []
        self.all_points = {}

        self.label = tk.Label(master, text=" 104 Mutliple Device Simulator ", font=("Arial", 14))
        self.label.pack(pady=10)

        # Frame to contain the log area and its scrollbar
        self.log_frame = tk.Frame(self.master)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH)

        # Scrollbar for the log area
        self.log_scrollbar = tk.Scrollbar(self.log_frame, orient=tk.VERTICAL)
        self.log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Log text area with the scrollbar
        self.log_text = tk.Text(self.log_frame, wrap='word', state='disabled', height=15, yscrollcommand=self.log_scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH)

        # Configure the scrollbar to scroll the Text widget
        self.log_scrollbar.config(command=self.log_text.yview)

        self.upload_button = tk.Button(master, text="Upload Excel File", command=self.upload_excel)
        self.upload_button.pack(pady=10)

        self.ip_frame = tk.Frame(master)
        self.ip_frame.pack(padx=20, pady=10)

        self.port_label = tk.Label(self.ip_frame, text="Port :", font=("Arial", 10) )
        self.port_label.pack(side=tk.LEFT, padx=(20, 0))

        self.port_entry = tk.Entry(self.ip_frame, font=("Arial", 10), width = 7, relief="sunken")
        self.port_entry.insert(0,2404)
        self.port_entry.pack(side=tk.LEFT, padx=(10, 0))

        self.asdu_label = tk.Label(self.ip_frame, text="ASDU : ", font=("Arial", 10) )
        self.asdu_label.pack(side=tk.LEFT, padx=(20, 0))

        self.asdu_entry = tk.Entry(self.ip_frame, font=("Arial", 10), width = 7, relief="sunken")
        self.asdu_entry.insert(0,1)
        self.asdu_entry.pack(side=tk.LEFT, padx=(10, 0))

        self.start_servers_button = tk.Button(self.master, text="Start Servers", command=self.update_signals, state='disabled')
        self.start_servers_button.pack(pady=5)

        self.stop_button = tk.Button(master, text="STOP", fg="red", command=self.stop_servers)
        self.stop_button.pack(pady=10)

        self.report_frame = tk.Frame(master)
        self.report_frame.pack(padx=20, pady=10)

        self.report_button_csv = tk.Button(self.report_frame, text="Generate Report CSV ", state="disabled", command = self.generate_report_csv)
        self.report_button_csv.pack(side = tk.LEFT)
        
        self.report_button_text = tk.Button(self.report_frame, text="Generate Report TEXT ", state="disabled", command = self.generate_report_txt)
        self.report_button_text.pack(side=tk.LEFT, padx=(20, 0))

        self.status_label = tk.Label(master, text="", fg="green", font=("Arial", 10))
        self.status_label.pack(pady=10)

        self.master.protocol("WM_DELETE_WINDOW", self.close_simulator)

    def upload_excel(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if file_paths:
            self.file_paths = file_paths
            self.log(f"Selected files: {self.file_paths}, Ready to start servers")
            self.start_servers_button.config(state='normal')

    def log(self, message):
        """Logs a message to the log text area."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.master.update()

        self.log_data.append({"Timestamp": datetime.datetime.now(), "Message": message})

    def generate_report_txt(self):
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(report_filename, "w") as report_file:
            report_file.write(f"IEC 104 Simulator Report\n")
            report_file.write(f"Generated on: {datetime.datetime.now()}\n\n")
            report_file.write(self.log_text.get("1.0", tk.END))

        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    def generate_report_csv(self):
        df = pd.DataFrame(self.log_data)
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(report_filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name="Log", index=False)

            # Customize Excel sheet appearance (optional)
            workbook = writer.book
            worksheet = writer.sheets["Log"]
            worksheet.set_column('A:B', 20)  # Adjust column width as needed
        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    # To display something in label
    def update_status(self, message):
        self.status_label.config(text=message)

    def setup_servers(self):
        """Starts the servers on a separate thread."""
        thread = threading.Thread(target=self.start_servers)
        thread.start()

    def start_servers(self):
        self.reset_server()
        self.report_button_csv['state'] = "disabled"
        self.report_button_text['state'] = "disabled"
        threads = []
        port = self.port_entry.get()
        asdu = self.asdu_entry.get()
        for file_path in self.file_paths:
            xls = pd.ExcelFile(file_path)
            df = pd.read_excel(xls)
            df = df[df['IP Address'] != '-']
            grouped = df.groupby('IP Address')

            for ip_address, group in grouped:
                self.log(f"Setting up server for IP: {ip_address}")
                server = c104.Server(ip=ip_address, port=int(port))
                station = server.add_station(common_address=int(asdu))
                server.start()
                time.sleep(1)

                self.servers.append(server)

                # Create a thread to handle connection and data processing for this server
                thread = threading.Thread(target=self.handle_server_connection, args=(server, station, group))
                threads.append(thread)
                thread.start()

        # Wait for all threads to finish
        for thread in threads:
            thread.join()

        self.log("All servers are running...")
        self.master.update_idletasks()
        mode = self.selected_mode.get()
        if mode == "one":
            self.process_signals_one_by_one()
        elif mode == "specific_ioa":
            self.process_specific_ioa()

    def handle_server_connection(self, server, station, group):
        while not server.has_active_connections:
            self.log(f"Waiting for connection to IP: {server.ip}")
            time.sleep(3)

        self.log(f"Connected to IP: {server.ip}")
        points = {}

        for _, row in group.iterrows():
            if row['Type ID'] in [1, 13, 30, 36, 45, 50]:
                point_type = iec104_type_ids[row['Type ID']]
                point = station.add_point(io_address=int(row['IOA']), type=point_type, report_ms=0)
                points[row['IOA']] = (point, row['Object Text'])
            else:
                self.log(f"Invalid Type ID {row['Type ID']} for IOA {row['IOA']}")

        # Collect all points across servers
        for ioa, (point, name) in points.items():
            if ioa not in self.all_points:
                self.all_points[ioa] = []
            self.all_points[ioa].append((point, name))

    def update_signals(self):
        mode_dialog = tk.Toplevel(self.master)
        mode_dialog.title("Select Processing Mode")

        mode_dialog.geometry("300x200+100+200")

        # Variable to store selected mode
        self.selected_mode = tk.StringVar(value="one")

        # Create radio buttons for selection
        one_radio = tk.Radiobutton(mode_dialog, text="One-by-One", variable=self.selected_mode, value="one", font=("Arial", 12))
        one_radio.pack(pady=10)

        all_radio = tk.Radiobutton(mode_dialog, text="All-at-Once", variable=self.selected_mode, value="all", font=("Arial", 12))
        all_radio.pack(pady=10)

        specific_ioa_radio = tk.Radiobutton(mode_dialog, text="Specific IOA", variable=self.selected_mode, value="specific_ioa", font=("Arial", 12))
        specific_ioa_radio.pack(pady=10)

        # Confirm button
        confirm_button = tk.Button(mode_dialog, text="Confirm", command=lambda: self.process_selected_mode(mode_dialog))
        confirm_button.pack(pady=10)

        mode_dialog.wait_window()

    def process_selected_mode(self, mode_dialog):
        """Process the selected mode and close the dialog."""
        mode_dialog.destroy()  # Close the mode selection dialog
        mode = self.selected_mode.get()  # Get the selected mode

        # Redirect to the corresponding processing logic
        if mode in ("specific_ioa", "one"):
            self.setup_servers()
        elif mode == "all":
            self.process_signals_all_at_once()
        else:
            messagebox.showerror("Error", "Invalid mode selected")

    def process_signals_one_by_one(self) : 
        for ioa, points in self.all_points.items():
            name = points[0][1]  # Get the signal name from the first point in the list
            if any(point.type in [c104.Type.C_SC_NA_1, c104.Type.C_SE_NC_1] for point, _ in points):
                self.get_command_signal_dialog(points, ioa, name)
            elif any(point.type in [c104.Type.M_SP_NA_1, c104.Type.M_SP_TB_1] for point, _ in points):
                self.show_binary_input_dialog(points, ioa, name)
            elif any(point.type in [c104.Type.M_ME_NC_1, c104.Type.M_ME_TF_1] for point, _ in points):
                self.show_numeric_input_dialog(points, ioa, name)

        self.log("Processing Completed")
        self.master.update_idletasks()
        time.sleep(5)
        self.stop_servers()
    
    def get_command_signal_dialog(self, points, ioa, name):
        if not self.server_running:
            return
        command_dialog = tk.Toplevel(self.master)
        self.current_dialog = command_dialog
        command_dialog.title("Command Signal")

        command_dialog.geometry("600x200+70+530")
        label = tk.Label(command_dialog, text=f"Signal: {name}\nIOA: {ioa}", font=("Arial", 12))
        label.pack(pady=10)

        fetch_button = tk.Button(command_dialog, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_command_signal(points))
        fetch_button.pack(pady = 10)

        next_button = tk.Button(command_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.break_loop_c(points, command_dialog))
        next_button.pack(padx=20, pady=10)

        command_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        command_dialog.wait_window()

    def show_binary_input_dialog(self, points, ioa, name):
        """Dialog for binary input (On/Off buttons) with signal details."""
        if not self.server_running:
            return
        binary_dialog = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog
        binary_dialog.title("Binary Input")

        binary_dialog.geometry("600x200+70+530")

        label = tk.Label(binary_dialog, text=f"Signal: {name}\nIOA: {ioa}", font=("Arial", 12))
        label.pack(pady=10)

        for point,name in points:
            point.report_ms = 1000

        on_button = tk.Button(binary_dialog, text="  On  ", fg="green", font=("Arial", 10),command=lambda: self.set_point_value(points, bool(1)))
        on_button.pack(side=tk.LEFT, padx=20, pady=10)

        off_button = tk.Button(binary_dialog, text="  Off  ", fg="Red", font=("Arial", 10),command=lambda: self.set_point_value(points, bool(0)))
        off_button.pack(side=tk.RIGHT, padx=20, pady=10)

        next_button = tk.Button(binary_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.break_loop_b(points, binary_dialog))
        next_button.pack(padx=20, pady=10)

        binary_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog.wait_window()

    def show_numeric_input_dialog(self, points, ioa, name):
        """Dialog for numeric input with signal details."""
        if not self.server_running:
            return
        numeric_dialog = tk.Toplevel(self.master)
        self.current_dialog = numeric_dialog
        numeric_dialog.title("Analog Input")

        numeric_dialog.geometry("600x200+70+530")

        label = tk.Label(numeric_dialog, text=f"Signal: {name}\nIOA: {ioa}", font=("Arial", 12))
        label.pack(pady=10)

        for point,name in points:
            point.report_ms = 1000

        value_entry = tk.Entry(numeric_dialog, font=("Arial", 12))
        value_entry.pack(pady=5)

        confirm_button = tk.Button(numeric_dialog, text="Confirm",
                                   command=lambda: self.confirm_numeric_input(points, value_entry))
        confirm_button.pack(pady=10)

        next_button = tk.Button(numeric_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.break_loop_n(points, numeric_dialog))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        numeric_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        numeric_dialog.wait_window()

    def get_command_signal(self, points):
        for point,name in points:
            val = point.value
            if point.type in [c104.Type.C_SC_NA_1]:
                value = bool(val)
            else:
                value = round(val,5)
            ip = point.station.server.ip
            self.log(f"IP:{ip} : Received point IOA : {point.io_address} : {name} : {value}")

    def set_point_value(self, points, value):
        """Set binary point values for all devices and close dialog."""
        for point, name in points:
            point.value = value
            ip = point.station.server.ip
            self.log(f"IP:{ip} : Set point IOA : {point.io_address} : {name} : {value}")

    def confirm_numeric_input(self, points, value_entry):
        value = float(value_entry.get())
        for point, name in points:
            point.value = value
            ip = point.station.server.ip
            self.log(f"IP:{ip} : Set point IOA : {point.io_address} : {name} : {value}")

    def break_loop_b(self, points, binary_dialog):
        for point, name in points:
            point.report_ms = 0
        binary_dialog.destroy()

    def break_loop_n(self, points, numeric_dialog):
        for point,name in points:
            point.report_ms = 0
        numeric_dialog.destroy()

    def break_loop_c(self, points, command_dialog):
        command_dialog.destroy()

    def process_signals_all_at_once(self):
        self.stop_servers()
        thread = threading.Thread(target=self.run_bat_file)
        thread.start()

    def run_bat_file(self):
        """Run a .bat file selected by the user."""
        try:
            # Create a hidden root window for the file dialog
            bat_file_path = filedialog.askopenfilename(
            title="Select BAT File",
            filetypes=[("Batch Files", "*.bat"), ("All Files", "*.*")])

            # If the user cancels the dialog, exit the function            
            if not bat_file_path:
                self.log("No file selected.")
                messagebox.showinfo("Info", "No file selected. Operation canceled.")
                return

            # Log the file path for debugging
            self.log(f"Selected BAT file: {bat_file_path}")

            # Execute the selected BAT file
            result = subprocess.run([bat_file_path], capture_output=True, text=True, shell=True)

            # Check for success or failure
            if result.returncode == 0:
                self.log(f"Successfully executed the BAT file: {bat_file_path}")
            else:
                self.log(f"Error while executing the BAT file: {bat_file_path}")
                self.log(f"Error Output:\n{result.stderr}")
                messagebox.showerror("Execution Error", f"Error while running BAT file:\n{result.stderr}")

        except subprocess.SubprocessError as e:
            self.log(f"Error running BAT file: {e}")
            messagebox.showerror("Error", f"An error occurred while running the BAT file:\n{e}")
        except Exception as e:
            self.log(f"Unexpected error: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
        
    def process_specific_ioa(self):
        """Process a specific IOA by updating its value across all devices."""
        if not self.server_running:
            return
        specific_ioa_dialog = tk.Toplevel(self.master)
        self.current_dialog = specific_ioa_dialog

        specific_ioa_dialog.title("Specific IOA Selection")
        specific_ioa_dialog.geometry("400x200+100+300")

        # Label to prompt user input
        label = tk.Label(specific_ioa_dialog, text="Enter IOA to update:", font=("Arial", 12))
        label.pack(pady=10)

        # Entry field for IOA
        ioa_entry = tk.Entry(specific_ioa_dialog, font=("Arial", 12))
        ioa_entry.pack(pady=10)

        # Confirm button
        confirm_button = tk.Button(specific_ioa_dialog,text="Confirm",command=lambda: self.check_ioa_and_show_dialog(ioa_entry.get(), specific_ioa_dialog))
        confirm_button.pack(pady=10)
        specific_ioa_dialog.wait_window()

    def check_ioa_and_show_dialog(self, ioa_input, dialog):
        """Check if IOA exists and determine whether to show binary or numeric dialog."""
        try:
            ioa = int(ioa_input)  # Parse IOA as an integer
            if ioa not in self.all_points:
                self.log(f"Error....No signals found for IOA: {ioa}")
                dialog.destroy()
                self.process_specific_ioa()
                return

            # Determine type (binary or numeric) based on the first point's type
            points = self.all_points[ioa]
            point_type = points[0][0].type

            if point_type in [c104.Type.C_SC_NA_1, c104.Type.C_SE_NC_1]:
                self.get_command_signal_dialog(points, ioa, points[0][1])
            elif point_type in [c104.Type.M_SP_NA_1, c104.Type.M_SP_TB_1]:
                self.show_binary_input_dialog(points, ioa, points[0][1])
            elif point_type in [c104.Type.M_ME_NC_1, c104.Type.M_ME_TF_1]:
                self.show_numeric_input_dialog(points, ioa, points[0][1])
            else:
                messagebox.showerror("Error", "Unsupported Type ID for the selected IOA.")
            dialog.destroy()

        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid numeric IOA.")
            dialog.destroy()
        
        self.process_specific_ioa()
        
    def dialog_closed(self):
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None

    def close_simulator(self):
        """Stop all servers and close the main window."""
        if messagebox.askokcancel("Quit", "Do you want to stop all servers and close the simulator?"):
            self.stop_servers()
            self.master.destroy()

    def stop_servers(self):
        """Stop all servers."""
        self.server_running = False
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None

        for server in self.servers:
            server.stop()
            self.log("Server stopped.")

        self.all_points.clear()
        self.servers.clear()
        self.report_button_csv['state'] = "normal"
        self.report_button_text['state'] = "normal"

    def reset_server(self):
        self.stop_servers()
        self.server_running = True
        self.all_points = {}  # Reset points dictionary
        self.servers = []

class IEC104client:
    def __init__(self, master):
        self.master = master
        self.current_dialog = None

        master.title("IEC 104 Master Simulator")
        master.geometry("700x720+20+20")

        self.client = None
        self.xls = None
        self.log_data = []
        self.signal_data = {}
        self.current_signal_index = 0

        self.label = tk.Label(master, text=" 104 Master Simulator ", font=("Arial", 14))
        self.label.pack(pady=10)

        self.log_frame = tk.Frame(self.master)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(self.log_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(self.log_frame, wrap='word', state='disabled', height=15, yscrollcommand=self.scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar.config(command=self.log_text.yview)

        self.upload_button = tk.Button(master, text="Upload Excel File", command=self.upload_excel)
        self.upload_button.pack(pady=10)

        self.file_name_label = tk.Label(master, text="", font=("Arial", 10))
        self.file_name_label.pack(pady=5)

        self.sheet_frame = tk.Frame(master)
        self.sheet_frame.pack(padx=20, pady=10)

        self.sheet_label = tk.Label(self.sheet_frame, text="Select Sheet:", font=("Arial", 10))
        self.sheet_label.pack(side=tk.LEFT)

        self.sheet_combo = ttk.Combobox(self.sheet_frame, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=(10, 0))

        self.load_ip_button = tk.Button(master, text="Load IPs", command=self.load_ips, state="disabled")
        self.load_ip_button.pack(pady=10)

        self.ip_frame = tk.Frame(master)
        self.ip_frame.pack(padx=20, pady=10)

        self.ip_label = tk.Label(self.ip_frame, text="Select IP Address :", font=("Arial", 10))
        self.ip_label.pack(side=tk.LEFT)

        self.ip_combo = ttk.Combobox(self.ip_frame, state="readonly")
        self.ip_combo.pack(side=tk.LEFT, padx=(10, 0))

        self.port_label = tk.Label(self.ip_frame, text="Port :", font=("Arial", 10) )
        self.port_label.pack(side=tk.LEFT, padx=(20, 0))

        self.port_entry = tk.Entry(self.ip_frame, font=("Arial", 10), width = 7, relief="sunken")
        self.port_entry.insert(0,2404)
        self.port_entry.pack(side=tk.LEFT, padx=(10, 0))

        self.asdu_label = tk.Label(self.ip_frame, text="ASDU : ", font=("Arial", 10) )
        self.asdu_label.pack(side=tk.LEFT, padx=(20, 0))

        self.asdu_entry = tk.Entry(self.ip_frame, font=("Arial", 10), width = 7, relief="sunken")
        self.asdu_entry.insert(0,1)
        self.asdu_entry.pack(side=tk.LEFT, padx=(10, 0))

        self.connect_button = tk.Button(master, text="Connect", command=self.connect)
        self.connect_button.pack(pady=10)

        self.stop_button = tk.Button(master, text="STOP", fg="red", command=self.stop_client)
        self.stop_button.pack(pady=10)

        self.report_frame = tk.Frame(master)
        self.report_frame.pack(padx=20, pady=10)

        self.report_button_csv = tk.Button(self.report_frame, text="Generate Report CSV ", state="disabled", command = self.generate_report_csv)
        self.report_button_csv.pack(side = tk.LEFT)
        
        self.report_button_text = tk.Button(self.report_frame, text="Generate Report TEXT ", state="disabled", command = self.generate_report_txt)
        self.report_button_text.pack(side=tk.LEFT, padx=(20, 0))

        self.status_label = tk.Label(master, text="", fg="green", font=("Arial", 10))
        self.status_label.pack(pady=10)

        self.master.protocol("WM_DELETE_WINDOW", self.close_simulator)

    def log(self, message):
        """Logs a message to the log text area."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.master.update()

        self.log_data.append({"Timestamp": datetime.datetime.now(), "Message": message})

    def generate_report_txt(self):
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(report_filename, "w") as report_file:
            report_file.write(f"IEC 104 Simulator Report\n")
            report_file.write(f"Generated on: {datetime.datetime.now()}\n\n")
            report_file.write(self.log_text.get("1.0", tk.END))

        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    def generate_report_csv(self):
        df = pd.DataFrame(self.log_data)
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(report_filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name="Log", index=False)

            # Customize Excel sheet appearance (optional)
            workbook = writer.book
            worksheet = writer.sheets["Log"]
            worksheet.set_column('A:B', 20)  # Adjust column width as needed
        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    def upload_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.xls = pd.ExcelFile(file_path)
            self.sheet_combo['values'] = self.xls.sheet_names
            self.sheet_combo.current(0)
            self.load_ip_button['state'] = "normal"

            file_name = file_path.split("/")[-1]
            self.file_name_label.config(text=f"Uploaded File: {file_name}")

    def load_ips(self):
        selected_sheet = self.sheet_combo.get()
        if not selected_sheet:
            messagebox.showwarning("Warning", "Please select a sheet first.")
            return

        df = pd.read_excel(self.xls, sheet_name=selected_sheet)
        unique_ips = df['IP Address'].dropna().unique()
        ip_list = [str(ip).strip() for ip in unique_ips]

        self.ip_combo['values'] = ip_list
        if ip_list:
            self.ip_combo.current(0)

    def connect(self):
        selected_sheet = self.sheet_combo.get()
        port = self.port_entry.get()
        asdu = self.asdu_entry.get()
        ip_address = self.ip_combo.get()
        if not selected_sheet or not ip_address:
            messagebox.showwarning("Warning", "Please select a sheet and IP address.")
            return
        self.reset_client()
        self.log(f"Setting up client for {ip_address}")
        self.client = c104.Client()
        self.connection = self.client.add_connection(ip=ip_address, port=int(port), init=c104.Init.ALL)
        self.station = self.connection.add_station(common_address=int(asdu))
        self.client.start()

        self.waiting_dots = 0
        self.log(f"Waiting for connection to {ip_address} ")
        self.check_connection()

    def check_connection(self):
        ip_address = self.ip_combo.get()
        if self.client.has_active_connections:
            self.log(f"{self.ip_combo.get()} Connected ")
            self.choose_processing_mode()
        else:
            '''self.waiting_dots = (self.waiting_dots + 1) % 4
            dots = "." * self.waiting_dots
            self.log(f"Waiting for connection{dots}")
            self.master.after(500, self.check_connection)
            '''
            self.log(f"Waiting for connection to {ip_address}")
            time.sleep(2)
            self.check_connection()

    # To display something in label
    def update_status(self, message):
        self.status_label.config(text=message)

    def choose_processing_mode(self):
        """Show a radio button dialog for choosing processing mode."""
        mode_dialog = tk.Toplevel(self.master)
        mode_dialog.title("Select Processing Mode")

        mode_dialog.geometry("300x200+100+300")

        self.selected_mode = tk.StringVar(value="one")

        one_radio = tk.Radiobutton(mode_dialog, text="One-by-One", variable=self.selected_mode, value="one", font=("Arial", 12))
        one_radio.pack(pady=10)

        all_radio = tk.Radiobutton(mode_dialog, text="All-at-Once", variable=self.selected_mode, value="all", font=("Arial", 12))
        all_radio.pack(pady=10)

        specific_ioa_radio = tk.Radiobutton(mode_dialog, text="Specific IOA", variable=self.selected_mode, value="specific_ioa", font=("Arial", 12))
        specific_ioa_radio.pack(pady=10)

        confirm_button = tk.Button(mode_dialog, text="Confirm", command=lambda: self.process_selected_mode(mode_dialog))
        confirm_button.pack(pady=10)

        mode_dialog.wait_window()  # Wait for the dialog to close

    def process_selected_mode(self, mode_dialog):
        """Process the selected mode and close the dialog."""
        mode_dialog.destroy()  
        self.mode = self.selected_mode.get()  

        # Redirect to the corresponding processing logic
        if self.mode == "specific_ioa":
            self.process_specific_ioa()
        elif self.mode== "one":
            self.process_signals_one_by_one()
        elif self.mode== "all":
            self.process_signals_all_at_once()
        else:
            messagebox.showerror("Error", "Invalid mode selected")

    def process_signals_one_by_one(self):
        selected_sheet = self.sheet_combo.get()
        df = pd.read_excel(self.xls, sheet_name=selected_sheet)
        grouped = df.groupby('IP Address')

        for ip_address, group in grouped:
            if not self.client.is_running or not self.client.has_active_connections :  
                self.log("Processing stopped either client stopped or client disconnected")
                break
            
            self.signal_data[ip_address] = list(group.iterrows()) 
            self.current_signal_index = 0

            if str(ip_address).strip() == self.ip_combo.get().strip():
                for _, row in self.signal_data[ip_address]:
                    ioa = row['IOA']
                    name = row['Object Text']
                    type_id = row['Type ID']
                    point_type = iec104_type_ids[row['Type ID']]
                    if not self.client.is_running or not self.client.has_active_connections :  
                        self.log("Processing stopped either client stopped or client disconnected")
                        break
                    if type_id in [1,13,30,36]:
                        self.station.add_point(io_address=int(row['IOA']), type=point_type)
                    elif type_id in [45,50] :
                        self.station.add_point(io_address=int(row['IOA']), type=point_type, command_mode = c104.CommandMode.SELECT_AND_EXECUTE)
                    else:
                        self.log(f"Invalid Type ID {row['Type ID']} for IOA {row['IOA']}")
        
        self.log("All points created, Ready for update ")
        self.update_signals( ip_address)
        self.log("         Processing completed            ")
        self.master.update_idletasks()
        time.sleep(5)
        self.stop_client()

    def process_signals_all_at_once(self):
        selected_sheet = self.sheet_combo.get()
        df = pd.read_excel(self.xls, sheet_name=selected_sheet)

        if 'UpdatedValue' in df.columns :
            df = df.drop(columns=['UpdatedValue'])
            df['UpdatedValue'] = df['value']
        else:
            df['UpdatedValue'] = df['value']

        grouped = df.groupby('IP Address')
        
        for ip_address, group in grouped:
            if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either client stopped or client disconnected")
                break

            points={}

            if str(ip_address).strip() == self.ip_combo.get().strip():
                for _, row in group.iterrows():
                    if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                        self.log("Processing stopped either client stopped or client disconnected")
                        break
                    point_type = iec104_type_ids[row['Type ID']]
                    if row['Type ID'] in [1,13,30,36]:
                        self.station.add_point(io_address=int(row['IOA']), type=point_type)
                    elif row['Type ID'] in [45,50] :
                        self.station.add_point(io_address=int(row['IOA']), type=point_type, command_mode = c104.CommandMode.SELECT_AND_EXECUTE)
                self.log("All points created, Ready for update ")

                self.log(f"Starting Continuous Value Updates for {ip_address}")

                while self.client.is_running:
                    for _, row in group.iterrows():
                        if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                            self.log("Processing stopped either client stopped or client disconnected")
                            break
                        ioa = int(row['IOA'])
                        type_id = row['Type ID']
                        name = row['Object Text']
                        point = self.station.get_point(io_address = ioa)
                        value = df.loc[row.name, 'UpdatedValue'] if pd.notna(df.loc[row.name, 'UpdatedValue']) else 0

                        if type_id == 45 :
                            point.value = (bool(value))
                            point.transmit(cause=c104.Cot.ACTIVATION)
                            self.log(f"Set point IOA : {ioa} : {name} : {bool(value)}")
                            value = not bool(value)
                            df.loc[row.name, 'UpdatedValue'] = int(value)
                        elif type_id == 50:
                            point.value = (float(value))
                            point.transmit(cause=c104.Cot.ACTIVATION)
                            self.log(f"Set point IOA : {ioa} : {name} : {value} ")
                            value = random.randint(10,100)
                            df.loc[row.name, 'UpdatedValue'] = value
                        elif type_id in [1,30]:
                            value = point.value
                            self.log(f"Received point IOA : {ioa} : {name} : {value}")
                            df.loc[row.name, 'UpdatedValue'] = value
                        elif type_id in [13,36]:
                            val = point.value
                            value = round(val,5)
                            self.log(f"Received point IOA : {ioa} : {name} : {value}")
                            df.loc[row.name, 'UpdatedValue'] = float(value)
                        else:
                            self.log(f"Invalid type id {type_id} for IOA {ioa}") 
                        
                        time.sleep(2)
                    if not self.client.is_running or not self.client.has_active_connections:
                        break
                    df.to_excel(self.xls, sheet_name=selected_sheet, index=False)
                    self.log("Saved updated values to Excel.")
                    self.log ("Next set of Update is starting........")
                    time.sleep(5)  # Sleep for 4 seconds before the next update loop

                if not self.client.is_running or not self.client.has_active_connections :
                    self.log("Stopped updates for All-at-Once mode.")
                    break

    def update_signals(self, ip_address):
        for _, row in self.signal_data[ip_address]:
            ioa = row['IOA']
            name = row['Object Text']
            type_id = row['Type ID']
            point = self.station.get_point(io_address = ioa)
            if type_id in [1,13,30,36]:
                if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either client stopped or client disconnected")
                self.show_input_signal_dialog(point, name, type_id, ip_address)
            elif type_id == 45:
                if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either client stopped or client disconnected")
                self.get_binary_command_dialog(point, name, type_id, ip_address)
            elif type_id == 50:
                if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                    self.log("Processing stopped either client stopped or client disconnected")
                self.get_numeric_command_dialog(point, name, type_id, ip_address)
            else:
                self.log(f"Invalid type id {type_id} for IOA {ioa}")
                continue
            
            self.current_signal_index +=1

    def process_specific_ioa(self):
        ioa_input = simpledialog.askinteger("Input", "Please enter the IOA:", minvalue=0)
        if ioa_input is None:
            self.choose_processing_mode()
            return  # User cancelled the input
        
        selected_ip = self.ip_combo.get()
        if not selected_ip:
            messagebox.showwarning("Warning", "Please select a valid IP address.")
            return

        df = pd.read_excel(self.xls, sheet_name=self.sheet_combo.get())
        group = df[df['IP Address'].astype(str).str.strip() == selected_ip]
        if group.empty:
            messagebox.showinfo("Info", f"No signals found for IP : {selected_ip}. and IOA : {ioa_input}")
            return

        result = group[group['IOA'] == ioa_input]
        if result.empty:
            messagebox.showinfo("Info", f"No signal found for IOA: {ioa_input}.")
            self.choose_processing_mode()
             
        row = result.iloc[0]
        name = row['Object Text']
        type_id = row['Type ID']
        ip_address = row['IP Address']
        point_type = iec104_type_ids[row['Type ID']]
        if type_id in [1,13,30,36]:
            self.station.add_point(io_address=int(ioa_input), type=point_type)
            point = self.station.get_point(io_address = ioa_input)
            if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either client stopped or client disconnected")
            self.show_input_signal_dialog(point, name, type_id, ip_address)
        elif type_id == 45: 
            self.station.add_point(io_address=int(row['IOA']), type=point_type, command_mode = c104.CommandMode.SELECT_AND_EXECUTE)
            point = self.station.get_point(io_address = ioa_input)
            if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either client stopped or client disconnected")
            self.get_binary_command_dialog(point, name, type_id, ip_address)
        elif type_id == 50:  
            self.station.add_point(io_address=int(row['IOA']), type=point_type, command_mode = c104.CommandMode.SELECT_AND_EXECUTE)
            point = self.station.get_point(io_address = ioa_input)                  
            if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either client stopped or client disconnected")
            self.get_numeric_command_dialog(point, name,type_id, ip_address)

        
        time.sleep(2)
        if self.client.is_running and self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
            self.choose_processing_mode()

    def show_previous_signal(self, point, type_id, ip_address):
        if self.mode == "specific_ioa":
            self.current_dialog.destroy()
        else:
            self.current_signal_index -= 1
            if self.current_signal_index < 0:
                self.current_signal_index = 0
                return # at the beginning
            self.current_dialog.destroy()
            self.process_current_signal(ip_address) # new function

    def show_next_signal(self, point, type_id, ip_address):
        if self.mode == "specific_ioa":
            self.current_dialog.destroy()
        else:
            self.current_signal_index += 1
            if self.current_signal_index >= len(self.signal_data[ip_address]):
                self.current_signal_index = len(self.signal_data[ip_address]) -1
                return # at the end
            self.current_dialog.destroy()
            self.process_current_signal(ip_address) # new function

    def process_current_signal(self, ip_address):
        _, row = self.signal_data[ip_address][self.current_signal_index]
        ioa = row['IOA']
        name = row['Object Text']
        type_id = row['Type ID']
        point = self.station.get_point(io_address = ioa)
        if type_id in [1,13,30,36]:
            if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either client stopped or client disconnected")
            self.show_input_signal_dialog(point, name, type_id, ip_address)
        elif type_id == 45:
            if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either client stopped or client disconnected")
            self.get_binary_command_dialog(point, name, type_id, ip_address)
        elif type_id == 50:
            if not self.client.is_running or not self.client.has_active_connections :  # Check if stop button was pressed or client disconnects
                self.log("Processing stopped either client stopped or client disconnected")
            self.get_numeric_command_dialog(point, name, type_id, ip_address)
        else:
            self.log(f"Invalid type id {type_id} for IOA {ioa}")
            self.current_signal_index += 1
            self.process_current_signal(ip_address)

    def show_input_signal_dialog(self, point, name, type_id, ip_address):
        if not self.client.is_running:
            return
        command_dialog = tk.Toplevel(self.master)
        self.current_dialog = command_dialog
        command_dialog.title("Command Signal")

        command_dialog.geometry("500x200+100+390")
        label = tk.Label(command_dialog, text=f"Signal: {name}\nIOA: {point.io_address}", font=("Arial", 12))
        label.pack(pady=10)

        self.display_text = tk.Text(command_dialog, height=2, width=25)
        self.display_text.pack(pady=5)

        fetch_button = tk.Button(command_dialog, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_input_signal(point, name, type_id,))
        fetch_button.pack(pady = 5)

        prev_button = tk.Button(command_dialog, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(point, type_id, ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=15)

        next_button = tk.Button(command_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(point, type_id, ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=15)

        command_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        command_dialog.wait_window()

    def get_input_signal(self, point, name, type_id):
        value = point.value
        self.display_text.delete('1.0',tk.END)
        self.display_text.insert('1.0', f"{value}")
        self.log(f"Received point IOA : {point.io_address} : {name} : {value}")
        time.sleep(2)

    def get_binary_command_dialog(self, point, name, type_id, ip_address):
        if not self.client.is_running:
            return
        binary_dialog = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog
        binary_dialog.title("Binary Input")

        binary_dialog.geometry("500x200+100+390")
        label = tk.Label(binary_dialog, text=f"Signal: {name}\nIOA: {point.io_address}", font=("Arial", 12))
        label.pack(pady=10)

        self.onoff_frame = tk.Frame(binary_dialog)
        self.onoff_frame.pack(padx=20, pady=10)

        on_button = tk.Button(self.onoff_frame, text="  On  ", fg="green", font=("Arial", 10),command=lambda: self.set_point_value(point, name, bool(1)))
        on_button.pack(side=tk.LEFT, padx = 10)

        off_button = tk.Button(self.onoff_frame, text="  Off  ", fg="red", font=("Arial", 10),command=lambda: self.set_point_value(point, name, bool(0)))
        off_button.pack(side= tk.RIGHT, padx = 10)

        prev_button = tk.Button(binary_dialog, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(point, type_id, ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=30)

        next_button = tk.Button(binary_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(point, type_id, ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=30)

        binary_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog.wait_window()

    def get_numeric_command_dialog(self, point, name, type_id, ip_address):
        if not self.client.is_running:
            return
        numeric_dialog = tk.Toplevel(self.master)
        self.current_dialog = numeric_dialog
        numeric_dialog.title("Analog Input")

        numeric_dialog.geometry("500x200+100+390")
        label = tk.Label(numeric_dialog, text=f"Signal: {name}\nIOA: {point.io_address}", font=("Arial", 12))
        label.pack(pady=10)

        value_entry = tk.Entry(numeric_dialog, font=("Arial", 12))
        value_entry.pack(pady=5)

        confirm_button = tk.Button(numeric_dialog, text="Confirm", command=lambda: self.set_numeric_value(point, name, value_entry))
        confirm_button.pack(pady=5)

        prev_button = tk.Button(numeric_dialog, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(point, type_id, ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=22)

        next_button = tk.Button(numeric_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(point, type_id, ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=22)

        numeric_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        numeric_dialog.wait_window()

    def set_numeric_value(self, point, name, value_entry):
        value = float(value_entry.get())
        point.value = value
        point.transmit(cause=c104.Cot.ACTIVATION)
        self.log(f"Set point IOA : {point.io_address} : {name} : {value}")
        time.sleep(2)

    def set_point_value(self, point, name, value):
        point.value = value
        point.transmit(cause=c104.Cot.ACTIVATION)
        self.log(f"Set point IOA : {point.io_address} : {name} : {value}")
        time.sleep(2)
   
    def dialog_closed(self):
        """Clear the dialog reference when it is manually closed."""
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None

    def close_simulator(self):
        """Stop all clients and close the main window."""
        if messagebox.askokcancel("Quit", "Do you want to stop all clients and close the simulator?"):
            self.stop_client()
            self.master.destroy()

    def stop_client(self):
        """Stop the IEC 104 client."""
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None  # Clear the reference

        if self.client:
            self.client.stop()
            self.log("client stopped .")
            print("client stopped.")
            self.report_button_csv['state'] = "normal"
            self.report_button_text['state'] = "normal"

    def reset_client(self):
        self.stop_client()
        self.client = None 
        self.station = None  

class Mbus_Master_Simulator:
    def __init__(self,master):
        self.master = master 
        self.client = None
        self.current_dialog = None
        self.client_running = True
        self.signal_data = {}
        self.log_data = []
        self.current_signal_index = 0

        master.title("Modbus Master Simulator")
        master.geometry("1000x720+20+20")

        self.label = tk.Label(master, text=" Modbus Master Simulator ", font=("Arial", 14))
        self.label.pack(pady=5)

        self.ip_frame = tk.Frame(master)
        self.ip_frame.pack(padx=10, pady=5)

        self.ip_label = tk.Label(self.ip_frame, text="Slave IP Address :", font=("Arial", 10))
        self.ip_label.pack(side=tk.LEFT)

        self.ip_entry = tk.Entry(self.ip_frame)
        self.ip_entry.pack(side= tk.LEFT, padx= (10,0))

        self.slave_id_label = tk.Label(self.ip_frame, text="Slave ID :", font=("Arial", 10))
        self.slave_id_label.pack(side=tk.LEFT)

        self.slave_id_entry = tk.Entry(self.ip_frame)
        self.slave_id_entry.insert(0,1)
        self.slave_id_entry.pack(side= tk.LEFT, padx= (10,0))

        self.port_label = tk.Label(self.ip_frame, text="Port :", font=("Arial", 10))
        self.port_label.pack(side=tk.LEFT)

        self.port_entry = tk.Entry(self.ip_frame)
        self.port_entry.insert(0,502)
        self.port_entry.pack(side= tk.LEFT, padx= (10,1))

        self.addr_label = tk.Label(self.ip_frame, text=" Starting Address :", font=("Arial", 10))
        self.addr_label.pack(side=tk.LEFT)

        self.addr_entry = tk.Entry(self.ip_frame)
        self.addr_entry.pack(side= tk.LEFT, padx= (10,1))

        self.connect_button = tk.Button(self.ip_frame, text="Connect" , fg= 'green', command=self.connect)
        self.connect_button.pack(pady=10)

        self.disconnect_button = tk.Button(self.ip_frame, text="Disonnect" , fg= 'Red', state = 'disabled', command=self.stop_client)
        self.disconnect_button.pack(padx=10)

        self.log_frame = tk.Frame(self.master)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(self.log_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(self.log_frame, wrap='word', state='disabled', height=15, yscrollcommand=self.scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar.config(command=self.log_text.yview)

        self.clear_button = tk.Button(self.log_frame, text="Clear Log", command=self.clear_log) 
        self.clear_button.pack(side = "bottom", pady = 10)
        self.clear_button.place(x= 895 ,y=340)

        self.upload_button = tk.Button(master, text="Upload Excel File", command=self.upload_excel)
        self.upload_button.pack(pady=10)

        self.file_name_label = tk.Label(master, text="", font=("Arial", 10))
        self.file_name_label.pack(pady=5)

        self.sheet_frame = tk.Frame(master)
        self.sheet_frame.pack(padx=20, pady=10)

        self.sheet_label = tk.Label(self.sheet_frame, text="Select Sheet:", font=("Arial", 10))
        self.sheet_label.pack(side=tk.LEFT)

        self.sheet_combo = ttk.Combobox(self.sheet_frame, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=(10, 0))

        self.update_button = tk.Button(self.master, text="Update Values" , state= 'disabled',  command=self.process_data)
        self.update_button.pack(pady=10)

        self.stop_button = tk.Button(self.master, text="STOP", fg="red", command=self.stop_client)
        self.stop_button.pack(pady=10)

        self.report_frame = tk.Frame(master)
        self.report_frame.pack(padx=20, pady=20)

        self.report_button_csv = tk.Button(self.report_frame, text="Generate Report CSV ", state="disabled", command = self.generate_report_csv)
        self.report_button_csv.pack(side = tk.LEFT)
        
        self.report_button_text = tk.Button(self.report_frame, text="Generate Report TEXT ", state="disabled", command = self.generate_report_txt)
        self.report_button_text.pack(side=tk.LEFT, padx=(20, 0))

        self.master.protocol("WM_DELETE_WINDOW", self.close_simulator)
        
    def log(self, message):
        """Logs a message to the log text area."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.master.update()

        self.log_data.append({"Timestamp": datetime.datetime.now(), "Message": message})

    def generate_report_txt(self):
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(report_filename, "w") as report_file:
            report_file.write(f"IEC 104 Simulator Report\n")
            report_file.write(f"Generated on: {datetime.datetime.now()}\n\n")
            report_file.write(self.log_text.get("1.0", tk.END))

        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    def generate_report_csv(self):
        df = pd.DataFrame(self.log_data)
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(report_filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name="Log", index=False)

            # Customize Excel sheet appearance (optional)
            workbook = writer.book
            worksheet = writer.sheets["Log"]
            worksheet.set_column('A:B', 20)  # Adjust column width as needed
        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")


    def clear_log(self):
        """Clears the log text area."""
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state='disabled')

    def upload_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.xls = pd.ExcelFile(file_path)
            self.sheet_combo['values'] = self.xls.sheet_names
            self.sheet_combo.current(0)

            file_name = file_path.split("/")[-1]
            self.file_name_label.config(text=f"Uploaded File: {file_name}")

    def connect(self):
        ip = self.ip_entry.get()
        slave_id = int(self.slave_id_entry.get())
        port = int(self.port_entry.get())

        if len(ip.split('.')) != 4:
            self.log("Invalid IP address format")
        else:
            self.reset_client()
            self.client = ModbusClient(ip, port, slave_id,timeout = 30, auto_open=True, auto_close= False)
            self.log("Modbus Master is started....")
            self.validate_connection()

    def check_connection(self):
        address = int(self.addr_entry.get())
        try:
            result = self.client.read_holding_registers(address,1)
            print(result)
            return result is not None
        except Exception as e:
            self.log(f"Connection error: {e}")
            return False
        
    def validate_connection(self):
        if self.check_connection():
            self.log("Slave Device connected.")
            self.connect_button['state'] = "disabled"
            self.disconnect_button['state'] = "normal"
            self.update_button['state'] = "normal"

        else:
            self.log(f"Waiting for connection to IP: {self.ip_entry.get()}")
            time.sleep(2)
            self.validate_connection()

    def process_data(self):
        selected_sheet = self.sheet_combo.get()
        df = pd.read_excel(self.xls, sheet_name=selected_sheet)
        grouped = df.groupby('IP Address')

        for ip_address, group in grouped:
            self.log(f"processing signals for IP : {ip_address} ")

            self.signal_data[ip_address] = list(group.iterrows()) # Store for navigation
            self.current_signal_index = 0 # reset for each IP

            self.process_signals_for_ip(ip_address) # new function
            
        self.log("All points Updated")
        time.sleep(3)
        self.stop_client()
    
    def process_signals_for_ip(self,ip_address):
        for _, row in self.signal_data[ip_address]:
            address = row['Index']
            function_code = row['Function Code']
            name = row['Name']

            if function_code == 1:
                self.get_coil_data(address, name, ip_address)
            elif function_code == 2:
                self.get_discrete_data(address, name, ip_address)
            elif function_code == 3:
                self.get_holding_data(address, name, ip_address)
            elif function_code == 4:
                self.get_input_data(address, name, ip_address)
            elif function_code == 5:
                self.show_binary_input_dialog(address, name, ip_address)
            elif function_code == 6:
                self.show_numeric_input_dialog(address, name, ip_address)
            elif function_code == 16:
                self.show_numeric_input_dialog_m(address, name, ip_address)
            else:
                self.log("Invalid Function code")

            self.current_signal_index +=1 # increment after each signal

    def show_previous_signal(self, ip_address):
        self.current_signal_index -= 1
        if self.current_signal_index < 0:
            self.current_signal_index = 0
            return # at the beginning
        self.current_dialog.destroy()
        self.process_current_signal(ip_address) # new function

    def show_next_signal(self, ip_address):
         self.current_signal_index += 1
         if self.current_signal_index >= len(self.signal_data[ip_address]):
            self.current_signal_index = len(self.signal_data[ip_address]) -1
            return # at the end
         self.current_dialog.destroy()
         self.process_current_signal(ip_address) # new function

    def process_current_signal(self, ip_address):
        _, row = self.signal_data[ip_address][self.current_signal_index]
        address = row['Index']
        function_code = row['Function Code']
        name = row['Name']

        if function_code == 1:
            self.get_coil_data(address, name, ip_address)
        elif function_code == 2:
            self.get_discrete_data(address, name, ip_address)
        elif function_code == 3:
            self.get_holding_data(address, name, ip_address)
        elif function_code == 4:
            self.get_input_data(address, name, ip_address)
        elif function_code == 5:
            self.show_binary_input_dialog(address, name, ip_address)
        elif function_code == 6:
            self.show_numeric_input_dialog(address, name, ip_address)
        elif function_code == 16:
            self.show_numeric_input_dialog_m(address, name, ip_address)
        else:
            self.log("Invalid Function code")
    
    def get_coil_data(self,address, name, ip_address):
        if not self.client_running:
            return
        binary_dialog_coil = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog_coil
        binary_dialog_coil.title("Read Coil Status")
        binary_dialog_coil.geometry("500x220+600+200")

        label = tk.Label(binary_dialog_coil, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.display_text = tk.Text(binary_dialog_coil, height=2, width=25)
        self.display_text.pack(pady=10)

        get_button = tk.Button(binary_dialog_coil, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_coil_value(address, binary_dialog_coil))
        get_button.pack(pady = 10)

        prev_button = tk.Button(binary_dialog_coil, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(binary_dialog_coil, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        get_button.focus_set()
        next_button.focus_set()

        binary_dialog_coil.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog_coil.wait_window()
    
    def get_coil_value(self, address, dialog = None):
        value = self.client.read_coils(address,1)
        self.display_text.delete('1.0',tk.END)
        rec_val = value[0]
        self.display_text.insert('1.0', f"{rec_val}")
        self.log(f"Read Input status at  Index : {address}, value {value}")

    def break_loop_c(self, binary_dialog_coil):
        binary_dialog_coil.destroy()

    def get_discrete_data(self,address, name, ip_address):
        if not self.client_running:
            return
        binary_dialog_o = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog_o
        binary_dialog_o.title("Read Input Status")
        binary_dialog_o.geometry("500x220+600+200")

        label = tk.Label(binary_dialog_o, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.display_text = tk.Text(binary_dialog_o, height=2, width=25)
        self.display_text.pack(pady=10)

        get_button = tk.Button(binary_dialog_o, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_discrete_value(address, binary_dialog_o))
        get_button.pack(pady=10)

        prev_button = tk.Button(binary_dialog_o, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(binary_dialog_o, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        binary_dialog_o.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog_o.wait_window()
    
    def get_discrete_value(self, address, dialog = None):
        value = self.client.read_discrete_inputs(address,1)
        self.display_text.delete('1.0',tk.END)
        rec_val = value[0]
        self.display_text.insert('1.0', f"{rec_val}")
        self.log(f"Read Input status at  Index : {address}, value {value}")

    def break_loop_d(self, binary_dialog_o):
        binary_dialog_o.destroy()

    def get_holding_data(self,address,name,ip_address):
        if not self.client_running:
            return
        analog_dialog_h = tk.Toplevel(self.master)
        self.current_dialog = analog_dialog_h
        analog_dialog_h.title("Read Holding Register")
        analog_dialog_h.geometry("500x300+600+200")

        label = tk.Label(analog_dialog_h, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.big_endian_c_mode = tk.BooleanVar()
        big_endian_c_button = tk.Checkbutton(analog_dialog_h, text = "Big_endian", variable = self.big_endian_c_mode)
        big_endian_c_button.pack()

        self.selected_dt_c = tk.StringVar()
        options = ["Float","Swapped Float","16bit signed Integer","16bit unsigned Integer","32bit signed Integer","32bit unsigned Integer"]
        self.data_type_combo = ttk.Combobox(analog_dialog_h, textvariable=self.selected_dt_c, state="readonly")
        self.data_type_combo['values'] = options
        self.data_type_combo.pack(pady=10)

        self.display_text_c = tk.Text(analog_dialog_h, height=2, width=25)
        self.display_text_c.pack(pady=10)

        get_button = tk.Button(analog_dialog_h, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_holding_value(address, analog_dialog_h))
        get_button.pack(pady=10)

        prev_button = tk.Button(analog_dialog_h, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(analog_dialog_h, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        analog_dialog_h.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        analog_dialog_h.wait_window()

    def get_holding_value(self, address, dialog = None):
        if not self.client_running:
            return
        registers = self.client.read_holding_registers( address, reg_nb = 2)
        big_endian_c = self.big_endian_c_mode.get()
        dt_c  = self.selected_dt_c.get()

        if dt_c == "Float":
            if big_endian_c:
                value = registers_to_float_be(registers)
            else:
                value = registers_to_float(registers)
            value = round(value,4)
            self.log(f"Read holding register {address}:  float value {value}")

        elif dt_c == "Swapped Float":
            registers_new = swap(registers)
            if big_endian_c:
                value = registers_to_float_be(registers_new)
            else:
                value = registers_to_float(registers_new)
            value = round(value,4)
            self.log(f"Read holding register {address} : swapped float value {value}")

        elif dt_c == "32bit unsigned Integer":
            if big_endian_c:
                value = registers_to_unsigned_integer_be(registers)
            else:
                value = registers_to_unsigned_integer(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "32bit signed Integer":
            if big_endian_c:
                value = registers_to_signed_integer_be(registers)
            else:
                value = registers_to_signed_integer(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "16bit unsigned Integer":
            if big_endian_c:
                value = registers_to_unsigned_16bit_be(registers)
            else:
                value = registers_to_unsigned_16bit(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "16bit signed Integer":
            if big_endian_c:
                value = registers_to_signed_16bit_be(registers)
            else:
                value = registers_to_signed_16bit(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        self.display_text_c.delete('1.0',tk.END)
        self.display_text_c.insert('1.0', f"{value}")

    def break_loop_holding(self, analog_dialog_h):
        analog_dialog_h.destroy()

    def get_input_data(self,address,name,ip_address):
        if not self.client_running:
            return
        analog_dialog_i = tk.Toplevel(self.master)
        self.current_dialog = analog_dialog_i
        analog_dialog_i.title("Read Holding Register")
        analog_dialog_i.geometry("500x300+600+200")

        label = tk.Label(analog_dialog_i, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.big_endian_c_mode = tk.BooleanVar()
        big_endian_c_button = tk.Checkbutton(analog_dialog_i, text = "Big_endian", variable = self.big_endian_c_mode)
        big_endian_c_button.pack()

        self.selected_dt_c = tk.StringVar()
        options = ["Float","Swapped Float","16bit signed Integer","16bit unsigned Integer","32bit signed Integer","32bit unsigned Integer"]
        self.data_type_combo = ttk.Combobox(analog_dialog_i, textvariable=self.selected_dt_c, state="readonly")
        self.data_type_combo['values'] = options
        self.data_type_combo.pack(pady=10)

        self.display_text_c = tk.Text(analog_dialog_i, height=2, width=25)
        self.display_text_c.pack(pady=10)

        get_button = tk.Button(analog_dialog_i, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_input_value(address, analog_dialog_i))
        get_button.pack(pady=10)

        prev_button = tk.Button(analog_dialog_i, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(analog_dialog_i, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        analog_dialog_i.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        analog_dialog_i.wait_window()

    def get_input_value(self, address, dialog = None):
        if not self.client_running:
            return
        registers = self.client.read_input_registers( address, reg_nb = 2)
        big_endian_c = self.big_endian_c_mode.get()
        dt_c  = self.selected_dt_c.get()

        if dt_c == "Float":
            if big_endian_c:
                value = registers_to_float_be(registers)
            else:
                value = registers_to_float(registers)
            value = round(value,4)
            self.log(f"Read holding register {address}:  float value {value}")

        elif dt_c == "Swapped Float":
            registers_new = swap(registers)
            if big_endian_c:
                value = registers_to_float_be(registers_new)
            else:
                value = registers_to_float(registers_new)
            value = round(value,4)
            self.log(f"Read holding register {address} : swapped float value {value}")

        elif dt_c == "32bit unsigned Integer":
            if big_endian_c:
                value = registers_to_unsigned_integer_be(registers)
            else:
                value = registers_to_unsigned_integer(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "32bit signed Integer":
            if big_endian_c:
                value = registers_to_signed_integer_be(registers)
            else:
                value = registers_to_signed_integer(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "16bit unsigned Integer":
            if big_endian_c:
                value = registers_to_unsigned_16bit_be(registers)
            else:
                value = registers_to_unsigned_16bit(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "16bit signed Integer":
            if big_endian_c:
                value = registers_to_signed_16bit_be(registers)
            else:
                value = registers_to_signed_16bit(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        self.display_text_c.delete('1.0',tk.END)
        self.display_text_c.insert('1.0', f"{value}")

    def break_loop_input(self, analog_dialog_i):
        analog_dialog_i.destroy()

    def show_binary_input_dialog(self, address, name, ip_address):
        if not self.client_running:
            return
        binary_dialog = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog
        binary_dialog.title("Write Single coil")

        binary_dialog.geometry("500x200+600+200")
        label = tk.Label(binary_dialog, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.onoff_frame = tk.Frame(binary_dialog)
        self.onoff_frame.pack(padx=20, pady=10)

        on_button = tk.Button(self.onoff_frame, text="  On  ", fg="green", font=("Arial", 10),
                              command=lambda: self.set_bool_value(address, 1, binary_dialog))
        on_button.pack(side=tk.LEFT)

        off_button = tk.Button(self.onoff_frame, text="  Off  ", fg="red", font=("Arial", 10),
                               command=lambda: self.set_bool_value(address, 0, binary_dialog))
        off_button.pack(side= tk.RIGHT)

        prev_button = tk.Button(binary_dialog, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=30)

        next_button = tk.Button(binary_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=30)

        binary_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog.wait_window()

    def set_bool_value(self, address, value, dialog=None):
        value_di = [bool(int(value))]
        self.client.write_single_coil(address, value_di[0])
        self.log(f"Set Coil status at  Index : {address} , with value {value_di}")

    def break_loop_bb(self, binary_dialog):
        binary_dialog.destroy()

    def show_numeric_input_dialog(self, address, name, ip_address):
        if not self.client_running:
            return
        numeric_dialog = tk.Toplevel(self.master)
        self.current_dialog = numeric_dialog
        numeric_dialog.title("Analog Input")

        numeric_dialog.geometry("500x300+600+200")
        label = tk.Label(numeric_dialog, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.big_endian_mode = tk.BooleanVar()
        big_endian_button = tk.Checkbutton(numeric_dialog, text = "Big_endian", variable = self.big_endian_mode)
        big_endian_button.pack()

        self.selected_dt = tk.StringVar()
        options = ["Float","Swapped Float","16bit signed Integer","16bit unsigned Integer","32bit signed Integer","32bit unsigned Integer"]
        self.data_type_combo = ttk.Combobox(numeric_dialog, textvariable=self.selected_dt, state="readonly")
        self.data_type_combo['values'] = options
        self.data_type_combo.pack(pady=10)

        value_entry = tk.Entry(numeric_dialog, font=("Arial", 12))
        value_entry.pack(pady=5)

        confirm_button = tk.Button(numeric_dialog, text="Confirm",
                                   command=lambda: self.set_numeric_input(address, value_entry , numeric_dialog))
        confirm_button.pack(pady=10)

        prev_button = tk.Button(numeric_dialog, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(numeric_dialog, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        numeric_dialog.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        numeric_dialog.wait_window()
    
    def set_numeric_input(self, address, value_entry, dialog=None):
        big_endian = self.big_endian_mode.get()
        dt  = self.selected_dt.get()
        value = value_entry.get()
        if dt == "Float":
            valuen = float(value)
            if big_endian:
                registers = float_to_registers_be(valuen)
            else:
                registers = float_to_registers(valuen)
            self.client.write_single_register( address  , registers[0])
            self.log(f"Updated holding register {address} with float value {value}")

        elif dt == "Swapped Float":
            valuen = float(value)
            if big_endian:
                registers = float_to_registers_be(valuen)
            else:
                registers = float_to_registers(valuen)
            registers_new = swap(registers)
            self.client.write_single_register( address  , registers_new[0])
            self.log(f"Updated holding register {address} with swapped float value {value}")

        elif dt == "32bit unsigned Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = unsigned_integer_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = unsigned_integer_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.client.write_single_register(address, registers[0])
            self.log(f"Updated holding register {address} with Unsigned Integer value {value}")

        elif dt == "32bit signed Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = signed_integer_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = signed_integer_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.client.write_single_register(address, registers[0])
            self.log(f"Updated holding register {address} with Signed Integer value {value}")

        elif dt == "16bit unsigned Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = unsigned_16bit_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = unsigned_16bit_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.client.write_single_register(address, registers[0])
            self.log(f"Updated holding register {address} with Unsigned Integer value {value}")

        elif dt == "16bit signed Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = signed_16bit_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = signed_16bit_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.client.write_single_register(address, registers[0])
            self.log(f"Updated holding register {address} with Signed Integer value {value}")

    def break_loop_aa(self, numeric_dialog):
        numeric_dialog.destroy()

    def show_numeric_input_dialog_m(self, address, name, ip_address):
        if not self.client_running:
            return
        numeric_dialog_m = tk.Toplevel(self.master)
        self.current_dialog = numeric_dialog_m
        numeric_dialog_m.title("Analog Input")

        numeric_dialog_m.geometry("500x300+600+200")
        label = tk.Label(numeric_dialog_m, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.big_endian_mode = tk.BooleanVar()
        big_endian_button = tk.Checkbutton(numeric_dialog_m, text = "Big_endian", variable = self.big_endian_mode)
        big_endian_button.pack()

        self.selected_dt = tk.StringVar()
        options = ["Float","Swapped Float","16bit signed Integer","16bit unsigned Integer","32bit signed Integer","32bit unsigned Integer"]
        self.data_type_combo = ttk.Combobox(numeric_dialog_m, textvariable=self.selected_dt, state="readonly")
        self.data_type_combo['values'] = options
        self.data_type_combo.pack(pady=10)

        value_entry = tk.Entry(numeric_dialog_m, font=("Arial", 12))
        value_entry.pack(pady=5)

        confirm_button = tk.Button(numeric_dialog_m, text="Confirm",
                                   command=lambda: self.set_numeric_input_m(address, value_entry , numeric_dialog_m))
        confirm_button.pack(pady=10)

        prev_button = tk.Button(numeric_dialog_m, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(numeric_dialog_m, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        numeric_dialog_m.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        numeric_dialog_m.wait_window()
    
    def set_numeric_input_m(self, address, value_entry, dialog=None):
        big_endian = self.big_endian_mode.get()
        dt  = self.selected_dt.get()
        value = value_entry.get()
        if dt == "Float":
            valuen = float(value)
            if big_endian:
                registers = float_to_registers_be(valuen)
            else:
                registers = float_to_registers(valuen)
            self.client.write_multiple_registers( address  , registers)
            self.log(f"Updated holding register {address} with float value {value}")

        elif dt == "Swapped Float":
            valuen = float(value)
            if big_endian:
                registers = float_to_registers_be(valuen)
            else:
                registers = float_to_registers(valuen)
            registers_new = swap(registers)
            self.client.write_multiple_registers( address  , registers_new)
            self.log(f"Updated holding register {address} with swapped float value {value}")

        elif dt == "32bit unsigned Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = unsigned_integer_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = unsigned_integer_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.client.write_multiple_registers(address, registers)
            self.log(f"Updated holding register {address} with Unsigned Integer value {value}")

        elif dt == "32bit signed Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = signed_integer_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = signed_integer_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.client.write_multiple_registers(address, registers)
            self.log(f"Updated holding register {address} with Signed Integer value {value}")

        elif dt == "16bit unsigned Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = unsigned_16bit_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = unsigned_16bit_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.client.write_multiple_registers(address, registers)
            self.log(f"Updated holding register {address} with Unsigned Integer value {value}")

        elif dt == "16bit signed Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = signed_16bit_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = signed_16bit_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.client.write_multiple_registers(address, registers)
            self.log(f"Updated holding register {address} with Signed Integer value {value}")

    def break_loop_aa_m(self, numeric_dialog_m):
        numeric_dialog_m.destroy()

    def dialog_closed(self):
        """Clear the dialog reference when it is manually closed."""
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None

    def close_simulator(self):
        if messagebox.askokcancel("Quit", "Do you want to stop all servers and close the simulator?"):
            self.stop_client()
            self.master.destroy()
    
    def stop_client(self):
        self.client_running = False
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None  # Clear the reference

        if self.client:
            self.client.close()
            self.log("Modbus Master stopped .")
            print("Modbus Master stopped.")
            self.connect_button['state'] = "normal"
            self.disconnect_button['state'] = "disabled"
            self.report_button_csv['state'] = "normal"
            self.report_button_text['state'] = "normal"

    def reset_client(self):
        self.stop_client()
        self.client_running = True
        self.client = None  
        self.all_points = {}

class Mbus_Slave_Single:
    def __init__(self, master):
        self.master = master
        self.current_dialog = None
        self.server_running = True
        self.signal_data = {}
        self.current_signal_index = 0

        master.title("Modbus Slave Simulator")
        master.geometry("700x770+20+20")

        self.label = tk.Label(master, text=" Modbus Single Device Simulator ", font=("Arial", 14))
        self.label.pack(pady=10)

        self.server = None
        self.xls = None
        self.all_points = {}  # Store all points with IOA keys for updates
        self.log_data = []

        self.log_frame = tk.Frame(self.master)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(self.log_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(self.log_frame, wrap='word', state='disabled', height=15, yscrollcommand=self.scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar.config(command=self.log_text.yview)

        self.upload_button = tk.Button(master, text="Upload Excel File", command=self.upload_excel)
        self.upload_button.pack(pady=10)

        self.file_name_label = tk.Label(master, text="", font=("Arial", 10))
        self.file_name_label.pack(pady=5)

        self.sheet_frame = tk.Frame(master)
        self.sheet_frame.pack(padx=20, pady=10)

        self.sheet_label = tk.Label(self.sheet_frame, text="Select Sheet:", font=("Arial", 10))
        self.sheet_label.pack(side=tk.LEFT)

        self.sheet_combo = ttk.Combobox(self.sheet_frame, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=(10, 0))

        self.load_ip_button = tk.Button(master, text="Load IPs", command=self.load_ips, state="disabled")
        self.load_ip_button.pack(pady=10)

        self.ip_frame = tk.Frame(master)
        self.ip_frame.pack(padx=20, pady=10)

        self.ip_label = tk.Label(self.ip_frame, text="Select IP Address :", font=("Arial", 10))
        self.ip_label.pack(side=tk.LEFT)

        self.ip_combo = ttk.Combobox(self.ip_frame, state="readonly")
        self.ip_combo.pack(side=tk.LEFT, padx=(10, 0))

        self.port_label = tk.Label(self.ip_frame, text="Port :", font=("Arial", 10))
        self.port_label.pack(side=tk.LEFT, padx=(20, 0))

        self.port_entry = tk.Entry(self.ip_frame, font=("Arial", 10), width = 7, relief="sunken")
        self.port_entry.insert(0,502)
        self.port_entry.pack(side=tk.LEFT, padx=(10, 0))

        self.connect_button = tk.Button(master, text="Connect", command=self.connect)
        self.connect_button.pack(pady=10)

        self.update_button = tk.Button(master, text="Update Values", command=self.choose_processing_mode, state = "disabled")
        self.update_button.pack(pady=10)

        self.stop_button = tk.Button(master, text="STOP", fg="red", command=self.stop_server)
        self.stop_button.pack(pady=10)

        self.report_frame = tk.Frame(master)
        self.report_frame.pack(padx=20, pady=20)

        self.report_button_csv = tk.Button(self.report_frame, text="Generate Report CSV ", state="disabled", command = self.generate_report_csv)
        self.report_button_csv.pack(side = tk.LEFT)
        
        self.report_button_text = tk.Button(self.report_frame, text="Generate Report TEXT ", state="disabled", command = self.generate_report_txt)
        self.report_button_text.pack(side=tk.LEFT, padx=(20, 0))

        self.status_label = tk.Label(master, text="", fg="green", font=("Arial", 10))
        self.status_label.pack(pady=10)

        self.master.protocol("WM_DELETE_WINDOW", self.close_simulator)

    def log(self, message):
        """Logs a message to the log text area."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.master.update()

        self.log_data.append({"Timestamp": datetime.datetime.now(), "Message": message})

    def generate_report_txt(self):
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(report_filename, "w") as report_file:
            report_file.write(f"IEC 104 Simulator Report\n")
            report_file.write(f"Generated on: {datetime.datetime.now()}\n\n")
            report_file.write(self.log_text.get("1.0", tk.END))

        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    def generate_report_csv(self):
        df = pd.DataFrame(self.log_data)
        report_filename = f"iec104_simulator_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(report_filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name="Log", index=False)

            # Customize Excel sheet appearance (optional)
            workbook = writer.book
            worksheet = writer.sheets["Log"]
            worksheet.set_column('A:B', 20)  # Adjust column width as needed
        messagebox.showinfo("Report Generated", f"Report saved as: {report_filename}")

    def upload_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.xls = pd.ExcelFile(file_path)
            self.sheet_combo['values'] = self.xls.sheet_names
            self.sheet_combo.current(0)
            self.load_ip_button['state'] = "normal"

            file_name = file_path.split("/")[-1]
            self.file_name_label.config(text=f"Uploaded File: {file_name}")

    def load_ips(self):
        selected_sheet = self.sheet_combo.get()
        if not selected_sheet:
            messagebox.showwarning("Warning", "Please select a sheet first.")
            return

        df = pd.read_excel(self.xls, sheet_name=selected_sheet)
        unique_ips = df['IP Address'].dropna().unique()
        ip_list = [str(ip).strip() for ip in unique_ips]

        self.ip_combo['values'] = ip_list
        if ip_list:
            self.ip_combo.current(0)

    def connect(self):
        selected_sheet = self.sheet_combo.get()
        port = self.port_entry.get()
        ip_address = self.ip_combo.get()
        if not selected_sheet or not ip_address:
            messagebox.showwarning("Warning", "Please select a sheet and IP address.")
            return
        self.reset_server()
        self.log(f"Setting up server for IP: {ip_address}")
        self.server = ModbusServer(ip_address, int(port), no_block=True)
        #self.station = self.server.add_station(common_address=1)
        self.server.start()
        self.log(f"Modbus Server started at IP : {ip_address}")
        time.sleep(2)
        self.update_button['state'] = "normal"

    def choose_processing_mode(self):
        """Show a radio button dialog for choosing processing mode."""
        mode_dialog = tk.Toplevel(self.master)
        mode_dialog.title("Select Processing Mode")

        mode_dialog.geometry("300x200+600+200")

        self.selected_mode = tk.StringVar(value="one")

        one_radio = tk.Radiobutton(mode_dialog, text="One-by-One", variable=self.selected_mode, value="one", font=("Arial", 12))
        one_radio.pack(pady=10)

        all_radio = tk.Radiobutton(mode_dialog, text="All-at-Once", variable=self.selected_mode, value="all", font=("Arial", 12))
        all_radio.pack(pady=10)

        specific_ioa_radio = tk.Radiobutton(mode_dialog, text="Specific Index", variable=self.selected_mode, value="specific_ioa", font=("Arial", 12))
        specific_ioa_radio.pack(pady=10)

        confirm_button = tk.Button(mode_dialog, text="Confirm", command=lambda: self.process_selected_mode(mode_dialog))
        confirm_button.pack(pady=10)

        mode_dialog.wait_window()  # Wait for the dialog to close

    def process_selected_mode(self, mode_dialog):
        """Process the selected mode and close the dialog."""
        mode_dialog.destroy()  
        mode = self.selected_mode.get()  

        # Redirect to the corresponding processing logic
        if mode == "specific_ioa":
            self.process_specific_ioa()
        elif mode == "one":
            self.process_signals_one_by_one()
        elif mode == "all":
            self.process_signals_all_at_once()
        else:
            messagebox.showerror("Error", "Invalid mode selected")

    def process_signals_one_by_one(self):
        selected_sheet = self.sheet_combo.get()
        df = pd.read_excel(self.xls, sheet_name=selected_sheet)
        grouped = df.groupby('IP Address')

        for ip_address, group in grouped:
            self.log(f"processing signals for IP : {ip_address} ")

            self.signal_data[ip_address] = list(group.iterrows()) # Store for navigation
            self.current_signal_index = 0 # reset for each IP

            self.process_signals_for_ip(ip_address) # new function

        self.log("All points Updated")
        time.sleep(3)
        self.stop_server()
    
    def process_signals_for_ip(self,ip_address):
        if str(ip_address).strip() == self.ip_combo.get().strip():
            for _, row in self.signal_data[ip_address]:
                address = row['Index']
                function_code = row['Function Code']
                name = row['Name']

                if function_code == 1:
                    self.coil_dialog(address, name, ip_address)
                elif function_code == 2:
                    self.discrete_dialog(address, name, ip_address)
                elif function_code == 3:
                    self.holding_dialog(address, name, ip_address)
                elif function_code == 4:
                    self.input_dialog(address, name, ip_address)
                elif function_code == 5:
                    self.get_binary_data(address, name, ip_address)
                elif function_code in[6, 16] :
                    self.get_analog_data(address, name, function_code, ip_address)
                else:
                    self.log("Invalid Function code")

                self.current_signal_index +=1 # increment after each signal

    def show_previous_signal(self, ip_address):
        self.current_signal_index -= 1
        if self.current_signal_index < 0:
            self.current_signal_index = 0
            return # at the beginning
        self.current_dialog.destroy()
        self.process_current_signal(ip_address) # new function

    def show_next_signal(self, ip_address):
         self.current_signal_index += 1
         if self.current_signal_index >= len(self.signal_data[ip_address]):
            self.current_signal_index = len(self.signal_data[ip_address]) -1
            return # at the end
         self.current_dialog.destroy()
         self.process_current_signal(ip_address) # new function

    def process_current_signal(self, ip_address):
        _, row = self.signal_data[ip_address][self.current_signal_index]
        address = row['Index']
        function_code = row['Function Code']
        name = row['Name']

        if function_code == 1:
            self.coil_dialog(address, name, ip_address)
        elif function_code == 2:
            self.discrete_dialog(address, name, ip_address)
        elif function_code == 3:
            self.holding_dialog(address, name, ip_address)
        elif function_code == 4:
            self.input_dialog(address, name, ip_address)
        elif function_code == 5:
            self.get_binary_data(address, name, ip_address)
        elif function_code in[6, 16] :
            self.get_analog_data(address, name, function_code, ip_address)
        else:
            self.log("Invalid Function code")
    
    def process_specific_ioa(self):
        if not self.server_running:
            return
        specific_ioa_dialog = tk.Toplevel(self.master)
        self.current_dialog = specific_ioa_dialog

        specific_ioa_dialog.title("Specific Index Selection")
        specific_ioa_dialog.geometry("400x200+600+300")

        label = tk.Label(specific_ioa_dialog, text="Enter Index and Point Type:", font=("Arial", 12))
        label.pack(pady=10)

        point_type_frame = tk.Frame(specific_ioa_dialog)
        point_type_frame.pack(padx=20, pady=10)

        point_type_label = tk.Label(point_type_frame, text="MODBUS Point Type :", font=("Arial", 10))
        point_type_label.pack(side=tk.LEFT)

        self.point_type_combo = ttk.Combobox(point_type_frame, state="readonly")
        self.point_type_combo.pack(side=tk.LEFT, padx=(10, 0))

        ioa_entry_frame = tk.Frame(specific_ioa_dialog)
        ioa_entry_frame.pack(padx=20, pady=10)

        ioa_entry_label = tk.Label(ioa_entry_frame, text="Index  :", font=("Arial", 10))
        ioa_entry_label.pack(side=tk.LEFT)

        ioa_entry = tk.Entry(ioa_entry_frame, font=("Arial", 10))
        ioa_entry.pack(padx=(10,0))

        confirm_button = tk.Button(specific_ioa_dialog, text="Confirm", command=lambda: self.check_ioa_and_show_dialog(ioa_entry.get(), specific_ioa_dialog))
        confirm_button.pack(pady=10)

        selected_sheet = self.sheet_combo.get()
        df = pd.read_excel(self.xls, sheet_name=selected_sheet)
        unique_point_type = df['Function Code'].dropna().unique()
        point_type_list = [str(pt).strip() for pt in unique_point_type]
        self.point_type_combo['values'] = point_type_list
        if point_type_list:
            self.point_type_combo.current(0)

        specific_ioa_dialog.wait_window()

    def check_ioa_and_show_dialog(self,ioa_input,dialog):

        point_type  = str(self.point_type_combo.get())
        ip_address = self.ip_combo.get().strip()
        ioa_input = int(ioa_input)

        if not ioa_input or not point_type:
            messagebox.showwarning("Warning", "Please select a point type and enter Index ")
            return  # User cancelled the input
        
        selected_ip = self.ip_combo.get()
        if not selected_ip:
            messagebox.showwarning("Warning", "Please select a valid IP address.")
            return
        
        df = pd.read_excel(self.xls, sheet_name=self.sheet_combo.get())
        group_ip = df[df['IP Address'].astype(str).str.strip() == selected_ip]
        if group_ip.empty:
            messagebox.showinfo("Info", f"No signals found for IP : {selected_ip}. and Index : {ioa_input}")
            return
        
        group_pt = group_ip[group_ip['Function Code'].astype(str).str.strip() == point_type]
        if group_pt.empty:
            messagebox.showinfo("Info", f"No signals found for Function Code : {point_type}. and Index : {ioa_input}")
            return

        result = group_pt[group_pt['Index'] == ioa_input]
        if result.empty:
            messagebox.showinfo("Info", f"No signal found for Index: {ioa_input}.")
             
        row = result.iloc[0]
        address = int(row['Index'])
        function_code = row['Function Code']
        name = row['Name']

        if function_code == 1:
            self.coil_dialog(address, name, ip_address)
        elif function_code == 2:
            self.discrete_dialog(address, name, ip_address)
        elif function_code == 3:
            self.holding_dialog(address, name, ip_address)
        elif function_code == 4:
            self.input_dialog(address, name, ip_address)
        elif function_code == 5:
            self.get_binary_data(address, name, ip_address)
        elif function_code in[6, 16] :
            self.get_analog_data(address, name, function_code, ip_address)
        else:
            self.log(" Invalid Input, Please enter valid signal detail. ")
            dialog.destroy()
        
        dialog.destroy()
        time.sleep(1)
        self.choose_processing_mode()
 
    def process_signals_all_at_once(self):
        selected_sheet = self.sheet_combo.get()
        df = pd.read_excel(self.xls, sheet_name=selected_sheet)

        if 'UpdatedValue' in df.columns :
            df = df.drop(columns=['UpdatedValue'])
            df['UpdatedValue'] = df['Value']
        else:
            df['UpdatedValue'] = df['Value']

        grouped = df.groupby('IP Address')

        for ip_address, group in grouped:
            if str(ip_address).strip() == self.ip_combo.get().strip():
                while self.server_running:
                    for _, row in group.iterrows():
                        if not self.server_running:  # Check if stop button was pressed
                            self.log("Processing stopped by user.")
                            break
                        address = int(row['Index'])
                        function_code = row['Function Code']
                        value = df.loc[row.name, 'UpdatedValue'] if pd.notna(df.loc[row.name, 'UpdatedValue']) else 0
                        mtype = row['Type']
                        endian = row['Endian']

                        if function_code ==1: # Coil Input
                            value_DI =[bool(int(value))]
                            self.server.data_bank.set_coils(address-1, value_DI)
                            self.log(f"Updated Coil input {address} with value {value_DI}")
                            value = not bool(value)
                            df.loc[row.name, 'UpdatedValue'] = int(value)
                            time.sleep(1)

                        elif function_code == 2: # Binary Input signal
                            value_DI = [bool(int(value))]
                            self.server.data_bank.set_discrete_inputs(address - 1, value_DI)
                            self.log(f"Updated Discrete input {address} with value {value_DI}")
                            value = not bool(value)
                            df.loc[row.name, 'UpdatedValue'] = int(value)
                            time.sleep(1)

                        elif function_code ==  5: # Binary Output signal
                            for i in range(3):
                                value = self.server.data_bank.get_coils(address - 1)
                                value_b = bool(value)
                                self.log(f" Read coil status at {address} value : {value}")
                                time.sleep(2)

                        elif function_code == 3: # Holding Register
                            if mtype == 'Float':
                                valuen = float(value)
                                if endian == 'Big':
                                    registers = float_to_registers_be(valuen)
                                elif endian == 'Little':
                                    registers = float_to_registers(valuen)
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_holding_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with float value {value}")

                            elif mtype == 'Swapped Float':
                                valuen = float(value)
                                if endian == 'Big':
                                    registers = float_to_registers_be(valuen)
                                elif endian == 'Little':
                                    registers = float_to_registers(valuen)
                                else:
                                    self.log("Invalid Endian Format")
                                registers_new = swap(registers)
                                self.server.data_bank.set_holding_registers(address - 1, registers_new)
                                self.log(f"Updated Holding input {address} with Swapped float value {value}")

                            elif mtype == '32bit unsigned Integer':
                                valuen = int(value)
                                if endian == 'Big':
                                    try:
                                        registers = unsigned_integer_to_register_be(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                elif endian == 'Little':
                                    try:
                                        registers = unsigned_integer_to_register(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_holding_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with Integer value {value}")

                            elif mtype == '32bit signed Integer':
                                valuen = int(value)
                                if endian == 'Big':
                                    try:
                                        registers = signed_integer_to_register_be(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                elif endian == 'Little':
                                    try:
                                        registers = signed_integer_to_register(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_holding_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with Integer value {value}")

                            elif mtype == '16bit unsigned Integer':
                                valuen = int(value)
                                if endian == 'Big':
                                    try:
                                        registers = unsigned_16bit_to_register_be(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                elif endian == 'Little':
                                    try:
                                        registers = unsigned_16bit_to_register(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_holding_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with Integer value {value}")

                            elif mtype == '16bit signed Integer':
                                valuen = int(value)
                                if endian == 'Big':
                                    try:
                                        registers = signed_16bit_to_register_be(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                elif endian == 'Little':
                                    try:
                                        registers = signed_16bit_to_register(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_holding_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with Integer value {value}")

                            time.sleep(1)
                            value = random.randint(10,100)
                            df.loc[row.name, 'UpdatedValue'] = value

                        elif function_code == 4 : # Input register
                            if mtype == 'Float':
                                valuen = float(value)
                                if endian == 'Big':
                                    registers = float_to_registers_be(valuen)
                                elif endian == 'Little':
                                    registers = float_to_registers(valuen)
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_input_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with float value {value}")

                            elif mtype == 'Swapped Float':
                                valuen = float(value)
                                if endian == 'Big':
                                    registers = float_to_registers_be(valuen)
                                elif endian == 'Little':
                                    registers = float_to_registers(valuen)
                                else:
                                    self.log("Invalid Endian Format")
                                registers_new = swap(registers)
                                self.server.data_bank.set_input_registers(address - 1, registers_new)
                                self.log(f"Updated Holding input {address} with Swapped float value {value}")

                            elif mtype == '32bit unsigned Integer':
                                valuen = int(value)
                                if endian == 'Big':
                                    try:
                                        registers = unsigned_integer_to_register_be(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                elif endian == 'Little':
                                    try:
                                        registers = unsigned_integer_to_register(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_input_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with Integer value {value}")

                            elif mtype == '32bit signed Integer':
                                valuen = int(value)
                                if endian == 'Big':
                                    try:
                                        registers = signed_integer_to_register_be(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                elif endian == 'Little':
                                    try:
                                        registers = signed_integer_to_register(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_input_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with Integer value {value}")

                            elif mtype == '16bit unsigned Integer':
                                valuen = int(value)
                                if endian == 'Big':
                                    try:
                                        registers = unsigned_16bit_to_register_be(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                elif endian == 'Little':
                                    try:
                                        registers = unsigned_16bit_to_register(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_input_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with Integer value {value}")

                            elif mtype == '16bit signed Integer':
                                valuen = int(value)
                                if endian == 'Big':
                                    try:
                                        registers = signed_16bit_to_register_be(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                elif endian == 'Little':
                                    try:
                                        registers = signed_16bit_to_register(valuen)
                                    except ValueError as e:
                                        self.log(f"Error {str(e)}")
                                else:
                                    self.log("Invalid Endian Format")
                                self.server.data_bank.set_input_registers(address - 1, registers)
                                self.log(f"Updated Holding input {address} with Integer value {value}")

                            time.sleep(1)
                            value = random.randint(10,100)
                            df.loc[row.name, 'UpdatedValue'] = value

                        elif function_code == 6: # Analog Output signal
                            for i in range (3):
                                registers = self.server.data_bank.get_holding_registers(address -1, number = 1)
                                if mtype == 'Float':
                                    if endian == 'Big':
                                        try:
                                            value = registers_to_float_be(registers)
                                        except ValueError as e:
                                            self.log(f"Error {str(e)}")
                                    else :
                                        try:
                                            value = registers_to_float(registers)
                                        except ValueError as e:
                                            self.log(f"Error {str(e)}")
                                    self.log(f"Read holding register {address}, float value: {round(value,4)}")
                                    
                                elif mtype == 'Swapped Float':
                                    registers_new  = swap(registers)
                                    if endian == 'Big':
                                        value = registers_to_float_be(registers_new)
                                    else:
                                        value = registers_to_float(registers_new)
                                    self.log(f"Read holding register {address}, Swapped float value: {round(value,4)}")

                                elif mtype == '32bit unsigned Integer':
                                    if endian == 'Big':
                                        value = registers_to_unsigned_integer_be(registers)
                                    else:
                                        value = registers_to_unsigned_integer(registers)
                                    self.log(f"Read holding register {address}, Integer value: {value}")

                                elif mtype == '32bit signed Integer':
                                    if endian == 'Big':
                                        value = registers_to_signed_integer_be(registers)
                                    else:
                                        value = registers_to_signed_integer(registers)
                                    self.log(f"Read holding register {address}, Integer value: {value}")

                                elif mtype == '16bit unsigned Integer':
                                    if endian == 'Big':
                                        value = registers_to_unsigned_16bit_be(registers)
                                    else:
                                        value = registers_to_unsigned_16bit(registers)
                                    self.log(f"Read holding register {address}, Integer value: {value}")

                                elif mtype == '16bit signed Integer':
                                    if endian == 'Big':
                                        value = registers_to_signed_16bit_be(registers)
                                    else:
                                        value = registers_to_signed_16bit(registers)
                                    self.log(f"Read holding register {address}, Integer value: {value}")
                                
                                time.sleep(1)
                        
                        elif function_code == 16: # Analog Output signal
                            for i in range (3):
                                registers = self.server.data_bank.get_holding_registers(address -1, number = 2)
                                if mtype == 'Float':
                                    if endian == 'Big':
                                       value = registers_to_float_be(registers)
                                    else :
                                       value = registers_to_float(registers)
                                    self.log(f"Read holding register {address}, float value: {round(value,4)}")
                                    
                                elif mtype == 'Swapped Float':
                                    registers_new  = swap(registers)
                                    if endian == 'Big':
                                        value = registers_to_float_be(registers_new)
                                    else:
                                        value = registers_to_float(registers_new)
                                    self.log(f"Read holding register {address}, Swapped float value: {round(value,4)}")

                                elif mtype == '32bit unsigned Integer':
                                    if endian == 'Big':
                                        value = registers_to_unsigned_integer_be(registers)
                                    else:
                                        value = registers_to_unsigned_integer(registers)
                                    self.log(f"Read holding register {address}, Integer value: {value}")

                                elif mtype == '32bit signed Integer':
                                    if endian == 'Big':
                                        value = registers_to_signed_integer_be(registers)
                                    else:
                                        value = registers_to_signed_integer(registers)
                                    self.log(f"Read holding register {address}, Integer value: {value}")

                                elif mtype == '16bit unsigned Integer':
                                    if endian == 'Big':
                                        value = registers_to_unsigned_16bit_be(registers)
                                    else:
                                        value = registers_to_unsigned_16bit(registers)
                                    self.log(f"Read holding register {address}, Integer value: {value}")

                                elif mtype == '16bit signed Integer':
                                    if endian == 'Big':
                                        value = registers_to_signed_16bit_be(registers)
                                    else:
                                        value = registers_to_signed_16bit(registers)
                                    self.log(f"Read holding register {address}, Integer value: {value}")
                                
                                time.sleep(1)
                        
                        else:
                            self.log(" Invalid Function code ")
                    df.to_excel(self.xls, sheet_name=selected_sheet, index=False)
                    self.log("Saved updated values to excel")
                    time.sleep(5)

    def coil_dialog(self, address, name, ip_address):
        if not self.server_running:
            return
        binary_dialog_c = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog_c
        binary_dialog_c.title("Coil Status")

        binary_dialog_c.geometry("500x300+600+200")
        label = tk.Label(binary_dialog_c, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.onoff_frame = tk.Frame(binary_dialog_c)
        self.onoff_frame.pack(padx=20, pady=10)

        on_button = tk.Button(self.onoff_frame, text="  On  ", fg="green", font=("Arial", 10),
                              command=lambda: self.set_coil_value(address, 1, binary_dialog_c))
        on_button.pack(side=tk.LEFT, padx=20, pady=10)

        off_button = tk.Button(self.onoff_frame, text="  Off  ", fg="red", font=("Arial", 10),
                               command=lambda: self.set_coil_value(address, 0, binary_dialog_c))
        off_button.pack(side=tk.RIGHT, padx=20, pady=10)

        prev_button = tk.Button(binary_dialog_c, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=30)

        next_button = tk.Button(binary_dialog_c, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=30)

        binary_dialog_c.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog_c.wait_window()

    def set_coil_value(self, address, value, dialog=None):
        value_di = [bool(int(value))]
        self.server.data_bank.set_coils(address - 1, value_di)
        self.log(f"Set Coil at  Index : {address} , with value {value_di}")

    def break_loop_c(self,binary_dialog_c):
        binary_dialog_c.destroy()

    def discrete_dialog(self, address, name, ip_address):
        if not self.server_running:
            return
        binary_dialog_d = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog_d
        binary_dialog_d.title("Discrete Input")

        binary_dialog_d.geometry("500x300+600+200")
        label = tk.Label(binary_dialog_d, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.onoff_frame = tk.Frame(binary_dialog_d)
        self.onoff_frame.pack(padx=20, pady=10)

        on_button = tk.Button(self.onoff_frame, text="  On  ", fg="green", font=("Arial", 10),
                              command=lambda: self.set_coil_value(address, 1, binary_dialog_d))
        on_button.pack(side=tk.LEFT, padx=20, pady=10)

        off_button = tk.Button(self.onoff_frame, text="  Off  ", fg="red", font=("Arial", 10),
                               command=lambda: self.set_coil_value(address, 0, binary_dialog_d))
        off_button.pack(side=tk.RIGHT, padx=20, pady=10)

        prev_button = tk.Button(binary_dialog_d, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=30)

        next_button = tk.Button(binary_dialog_d, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=30)

        binary_dialog_d.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog_d.wait_window()

    def set_discrete_value(self, address, value, dialog=None):
        value_di = [bool(int(value))]
        self.server.data_bank.set_discrete_inputs(address - 1, value_di)
        self.log(f"Set input status at  Index : {address} , with value {value_di}")

    def break_loop_d(self,binary_dialog_d):
        binary_dialog_d.destroy()

    def holding_dialog(self, address, name, ip_address):
        if not self.server_running:
            return
        numeric_dialog_h = tk.Toplevel(self.master)
        self.current_dialog = numeric_dialog_h
        numeric_dialog_h.title("Holding Register")

        numeric_dialog_h.geometry("500x300+600+200")
        label = tk.Label(numeric_dialog_h, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.big_endian_mode = tk.BooleanVar()
        big_endian_button = tk.Checkbutton(numeric_dialog_h, text = "Big_endian", variable = self.big_endian_mode)
        big_endian_button.pack()

        self.selected_dt = tk.StringVar()
        options = ["Float","Swapped Float","16bit signed Integer","16bit unsigned Integer","32bit signed Integer","32bit unsigned Integer"]
        self.data_type_combo = ttk.Combobox(numeric_dialog_h, textvariable=self.selected_dt, state="readonly")
        self.data_type_combo['values'] = options
        self.data_type_combo.pack(pady=10)

        value_entry = tk.Entry(numeric_dialog_h, font=("Arial", 12))
        value_entry.pack(pady=5)

        confirm_button = tk.Button(numeric_dialog_h, text="Confirm",
                                   command=lambda: self.set_holding_value(address, value_entry , numeric_dialog_h))
        confirm_button.pack(pady=10)

        prev_button = tk.Button(numeric_dialog_h, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(numeric_dialog_h, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        numeric_dialog_h.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        numeric_dialog_h.wait_window()
    
    def set_holding_value(self, address, value_entry, dialog=None):
        big_endian = self.big_endian_mode.get()
        dt  = self.selected_dt.get()
        value = value_entry.get()
        if dt == "Float":
            valuen = float(value)
            if big_endian:
                registers = float_to_registers_be(valuen)
            else:
                registers = float_to_registers(valuen)
            self.server.data_bank.set_holding_registers( address - 1  , registers)
            self.log(f"Updated holding register {address} with float value {value}")

        elif dt == "Swapped Float":
            valuen = float(value)
            if big_endian:
                registers = float_to_registers_be(valuen)
            else:
                registers = float_to_registers(valuen)
            registers_new = swap(registers)
            self.server.data_bank.set_holding_registers( address - 1  , registers_new)
            self.log(f"Updated holding register {address} with swapped float value {value}")

        elif dt == "32bit unsigned Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = unsigned_integer_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = unsigned_integer_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.server.data_bank.set_holding_registers(address - 1, registers)
            self.log(f"Updated holding register {address} with Unsigned Integer value {value}")

        elif dt == "32bit signed Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = signed_integer_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = signed_integer_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.server.data_bank.set_holding_registers(address - 1, registers)
            self.log(f"Updated holding register {address} with Signed Integer value {value}")

        elif dt == "16bit unsigned Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = unsigned_16bit_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = unsigned_16bit_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.server.data_bank.set_holding_registers(address - 1, registers)
            self.log(f"Updated holding register {address} with Unsigned Integer value {value}")

        elif dt == "16bit signed Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = signed_16bit_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = signed_16bit_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.server.data_bank.set_holding_registers(address - 1, registers)
            self.log(f"Updated holding register {address} with Signed Integer value {value}")

    def break_loop_h(self,numeric_dialog_h):
        numeric_dialog_h.destroy()

    def input_dialog(self, address, name, ip_address):
        if not self.server_running:
            return
        numeric_dialog_i = tk.Toplevel(self.master)
        self.current_dialog = numeric_dialog_i
        numeric_dialog_i.title("Input Register")

        numeric_dialog_i.geometry("500x300+600+200")
        label = tk.Label(numeric_dialog_i, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.big_endian_mode = tk.BooleanVar()
        big_endian_button = tk.Checkbutton(numeric_dialog_i, text = "Big_endian", variable = self.big_endian_mode)
        big_endian_button.pack()

        self.selected_dt = tk.StringVar()
        options = ["Float","Swapped Float","16bit signed Integer","16bit unsigned Integer","32bit signed Integer","32bit unsigned Integer"]
        self.data_type_combo = ttk.Combobox(numeric_dialog_i, textvariable=self.selected_dt, state="readonly")
        self.data_type_combo['values'] = options
        self.data_type_combo.pack(pady=10)

        value_entry = tk.Entry(numeric_dialog_i, font=("Arial", 12))
        value_entry.pack(pady=5)

        confirm_button = tk.Button(numeric_dialog_i, text="Confirm",
                                   command=lambda: self.set_input_value(address, value_entry , numeric_dialog_i))
        confirm_button.pack(pady=10)

        prev_button = tk.Button(numeric_dialog_i, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(numeric_dialog_i, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        numeric_dialog_i.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        numeric_dialog_i.wait_window()
    
    def set_input_value(self, address, value_entry, dialog=None):
        big_endian = self.big_endian_mode.get()
        dt  = self.selected_dt.get()
        value = value_entry.get()
        if dt == "Float":
            valuen = float(value)
            if big_endian:
                registers = float_to_registers_be(valuen)
            else:
                registers = float_to_registers(valuen)
            self.server.data_bank.set_input_registers( address - 1  , registers)
            self.log(f"Updated holding register {address} with float value {value}")

        elif dt == "Swapped Float":
            valuen = float(value)
            if big_endian:
                registers = float_to_registers_be(valuen)
            else:
                registers = float_to_registers(valuen)
            registers_new = swap(registers)
            self.server.data_bank.set_input_registers( address - 1  , registers_new)
            self.log(f"Updated holding register {address} with swapped float value {value}")

        elif dt == "32bit unsigned Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = unsigned_integer_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = unsigned_integer_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.server.data_bank.set_input_registers(address - 1, registers)
            self.log(f"Updated holding register {address} with Unsigned Integer value {value}")

        elif dt == "32bit signed Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = signed_integer_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = signed_integer_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.server.data_bank.set_input_registers(address - 1, registers)
            self.log(f"Updated holding register {address} with Signed Integer value {value}")

        elif dt == "16bit unsigned Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = unsigned_16bit_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = unsigned_16bit_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.server.data_bank.set_input_registers(address - 1, registers)
            self.log(f"Updated holding register {address} with Unsigned Integer value {value}")

        elif dt == "16bit signed Integer":
            valuen = int(value)
            if big_endian:
                try:
                    registers = signed_16bit_to_register_be(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            else:
                try:
                    registers = signed_16bit_to_register(valuen)
                except ValueError as e:
                    self.log(f"Error {str(e)}")
            self.server.data_bank.set_input_registers(address - 1, registers)
            self.log(f"Updated holding register {address} with Signed Integer value {value}")
    
    def break_loop_i(self,numeric_dialog_i):
        numeric_dialog_i.destroy()

    def get_binary_data(self,address,name, ip_address):
        if not self.server_running:
            return
        binary_dialog_o = tk.Toplevel(self.master)
        self.current_dialog = binary_dialog_o
        binary_dialog_o.title("Binary Output")
        binary_dialog_o.geometry("500x300+600+200")

        label = tk.Label(binary_dialog_o, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.display_text = tk.Text(binary_dialog_o, height=2, width=25)
        self.display_text.pack(pady=10)

        get_button = tk.Button(binary_dialog_o, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_bool_value(address, binary_dialog_o))
        get_button.pack(pady = 10)

        prev_button = tk.Button(binary_dialog_o, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(binary_dialog_o, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)

        binary_dialog_o.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        binary_dialog_o.wait_window()
    
    def get_bool_value(self, address, dialog = None):
        value = self.server.data_bank.get_coils(address - 1)
        self.display_text.delete('1.0',tk.END)
        rec_val = value[0]
        self.display_text.insert('1.0', f"{rec_val}")
        self.log(f"Read coil status at  Index : {address}, value {value}")

    def break_loop_b(self, binary_dialog_o):
        binary_dialog_o.destroy()

    def get_analog_data(self,address,name,function_code,ip_address):
        if not self.server_running:
            return
        analog_dialog_o = tk.Toplevel(self.master)
        self.current_dialog = analog_dialog_o
        analog_dialog_o.title("Analog Output")
        analog_dialog_o.geometry("500x300+600+200")

        label = tk.Label(analog_dialog_o, text=f"Signal: {name}\nIndex: {address}", font=("Arial", 12))
        label.pack(pady=10)

        self.big_endian_c_mode = tk.BooleanVar()
        big_endian_c_button = tk.Checkbutton(analog_dialog_o, text = "Big_endian", variable = self.big_endian_c_mode)
        big_endian_c_button.pack()

        self.selected_dt_c = tk.StringVar()
        options = ["Float","Swapped Float","16bit signed Integer","16bit unsigned Integer","32bit signed Integer","32bit unsigned Integer"]
        self.data_type_combo = ttk.Combobox(analog_dialog_o, textvariable=self.selected_dt_c, state="readonly")
        self.data_type_combo['values'] = options
        self.data_type_combo.pack(pady=10)

        self.display_text_c = tk.Text(analog_dialog_o, height=2, width=25)
        self.display_text_c.pack(pady=10)

        get_button = tk.Button(analog_dialog_o, text=" FETCH ", font=("Arial", 10), state = 'normal', command=lambda: self.get_analog_value(address, function_code, analog_dialog_o))
        get_button.pack(pady=10)

        prev_button = tk.Button(analog_dialog_o, text=" PREV ", font=("Arial", 10), state = 'normal', command=lambda: self.show_previous_signal(ip_address))
        prev_button.pack(side=tk.LEFT, padx=20, pady=10)

        next_button = tk.Button(analog_dialog_o, text=" NEXT ", font=("Arial", 10), state = 'normal', command=lambda: self.show_next_signal(ip_address))
        next_button.pack(side=tk.RIGHT, padx=20, pady=10)
        get_button.focus_set()
        next_button.focus_set()

        analog_dialog_o.protocol("WM_DELETE_WINDOW", self.dialog_closed)
        analog_dialog_o.wait_window()

    def get_analog_value(self, address, function_code, dialog = None):
        if not self.server_running:
            return
        if function_code == 6:
            registers = self.server.data_bank.get_holding_registers( address - 1, number = 1)
        else:
            registers = self.server.data_bank.get_holding_registers( address - 1, number = 2)
        big_endian_c = self.big_endian_c_mode.get()
        dt_c  = self.selected_dt_c.get()

        if dt_c == "Float":
            if big_endian_c:
                value = registers_to_float_be(registers)
            else:
                value = registers_to_float(registers)
            value = round(value,4)
            self.log(f"Read holding register {address}:  float value {value}")

        elif dt_c == "Swapped Float":
            registers_new = swap(registers)
            if big_endian_c:
                value = registers_to_float_be(registers_new)
            else:
                value = registers_to_float(registers_new)
            value = round(value,4)
            self.log(f"Read holding register {address} : swapped float value {value}")

        elif dt_c == "32bit unsigned Integer":
            if big_endian_c:
                value = registers_to_unsigned_integer_be(registers)
            else:
                value = registers_to_unsigned_integer(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "32bit signed Integer":
            if big_endian_c:
                value = registers_to_signed_integer_be(registers)
            else:
                value = registers_to_signed_integer(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "16bit unsigned Integer":
            if big_endian_c:
                value = registers_to_unsigned_16bit_be(registers)
            else:
                value = registers_to_unsigned_16bit(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        elif dt_c == "16bit signed Integer":
            if big_endian_c:
                value = registers_to_signed_16bit_be(registers)
            else:
                value = registers_to_signed_16bit(registers)
            self.log(f"Read holding register {address}:  Integer value {value}")

        self.display_text_c.delete('1.0',tk.END)
        self.display_text_c.insert('1.0', f"{value}")

    def break_loop_a(self, analog_dialog_o):
        analog_dialog_o.destroy()

    def dialog_closed(self):
        """Clear the dialog reference when it is manually closed."""
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None

    def close_simulator(self):
        """Stop all servers and close the main window."""
        if messagebox.askokcancel("Quit", "Do you want to stop all servers and close the simulator?"):
            self.stop_server()
            self.master.destroy()

    def stop_server(self):
        """Stop the IEC 104 server."""
        self.server_running = False
        if self.current_dialog:
            self.current_dialog.destroy()
            self.current_dialog = None  # Clear the reference

        if self.server:
            self.server.stop()
            self.log("Server stopped .")
            print("Server stopped.")
            self.report_button_csv['state'] = "normal"
            self.report_button_text['state'] = "normal"

        self.all_points.clear()

    def reset_server(self):
        self.stop_server()
        self.server_running = True
        self.server = None  
        self.all_points = {}

class Mbus_Slave_Multiple:
    def __init__(self, master):
        self.master = master
        self.current_dialog = None
        self.server_running = True

        self.master.title(" Modbus Slave Simulator ")
        self.master.geometry("700x450+20+20")

        self.label = tk.Label(master, text=" Modbus Multiple Device Simulator ", font=("Arial", 14))
        self.label.pack(pady=10)

        # Frame to contain the log area and its scrollbar
        self.log_frame = tk.Frame(self.master)
        self.log_frame.pack(padx=10, pady=10, fill=tk.BOTH)

        # Scrollbar for the log area
        self.log_scrollbar = tk.Scrollbar(self.log_frame, orient=tk.VERTICAL)
        self.log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Log text area with the scrollbar
        self.log_text = tk.Text(self.log_frame,wrap='word',state='disabled',height=15,yscrollcommand=self.log_scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH)

        # Configure the scrollbar to scroll the Text widget
        self.log_scrollbar.config(command=self.log_text.yview)

        self.start_servers_button = tk.Button(self.master, text="Start Servers", command=self.process_signals_all_at_once)
        self.start_servers_button.pack(pady=5)

        self.status_label = tk.Label(master, text="", fg="green", font=("Arial", 10))
        self.status_label.pack(pady=10)

        self.master.protocol("WM_DELETE_WINDOW", self.close_simulator)

    def log(self, message):
        """Logs a message to the log text area."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.master.update()

    def process_signals_all_at_once(self):
        thread = threading.Thread(target=self.run_bat_file)
        thread.start()

    def run_bat_file(self):
        """Run a .bat file selected by the user."""
        try:
            # Create a hidden root window for the file dialog
            bat_file_path = filedialog.askopenfilename(
            title="Select BAT File",
            filetypes=[("Batch Files", "*.bat"), ("All Files", "*.*")])

            # If the user cancels the dialog, exit the function            
            if not bat_file_path:
                self.log("No file selected.")
                messagebox.showinfo("Info", "No file selected. Operation canceled.")
                return

            # Log the file path for debugging
            self.log(f"Selected BAT file: {bat_file_path}")

            # Execute the selected BAT file
            result = subprocess.run([bat_file_path], capture_output=True, text=True, shell=True)

            # Check for success or failure
            if result.returncode == 0:
                self.log(f"Successfully executed the BAT file: {bat_file_path}")
            else:
                self.log(f"Error while executing the BAT file: {bat_file_path}")
                self.log(f"Error Output:\n{result.stderr}")
                messagebox.showerror("Execution Error", f"Error while running BAT file:\n{result.stderr}")

        except subprocess.SubprocessError as e:
            self.log(f"Error running BAT file: {e}")
            messagebox.showerror("Error", f"An error occurred while running the BAT file:\n{e}")
        except Exception as e:
            self.log(f"Unexpected error: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")

    def close_simulator(self):
        """Stop all servers and close the main window."""
        if messagebox.askokcancel("Quit", "Do you want to stop all servers and close the simulator?"):
            self.master.destroy()

root = tk.Tk()
root.geometry("700x720+20+20")
root.title("Protocol Simulator")

def run_simulator():
    option  = my_combo.get()
    for widget in root.winfo_children():
        if widget not in [my_frame]:  # Don't destroy the combo frame
            widget.destroy()
    if option == "104 Slave - Single Device":
        IEC104SlaveSingle(root)
    elif option == "104 Slave - Multiple Device":
        IEC104SlaveMultiple(root)
    elif option == "104 Master" :
        IEC104client(root)
    elif option == "Modbus Master":
        Mbus_Master_Simulator(root)
    elif option == "Modbus Slave - Single Device" :
        Mbus_Slave_Single(root)
    elif option == "Modbus Slave - Multiple Device":
        Mbus_Slave_Multiple(root)


simulator = ["104 Slave - Single Device", "104 Slave - Multiple Device", "104 Master", "Modbus Master", "Modbus Slave - Single Device", "Modbus Slave - Multiple Device"]

my_frame = tk.Frame(root)
my_frame.pack(padx=20, pady=10)

my_label = tk.Label(my_frame, text="Select Simulator :", font=("Arial", 10))
my_label.pack(side=tk.LEFT)

my_combo = ttk.Combobox(my_frame, values=simulator, state="readonly", justify="center", width = 35)
my_combo.current(0)
my_combo.pack(side=tk.LEFT, padx=(10, 0))

start_button  = tk.Button(my_frame, text = " START ", command = run_simulator)
start_button.pack(side = tk.RIGHT, padx = 20)

root.mainloop()