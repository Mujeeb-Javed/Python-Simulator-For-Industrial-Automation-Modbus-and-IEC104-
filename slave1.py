import time
import struct
import random
import pandas as pd
from pyModbusTCP.server import ModbusServer, DataBank

port = 502
# Helper functions to convert float to/from registers
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

file_path = '1.xlsx'
xls = pd.ExcelFile(file_path)

def main():
    # Load and preprocess the Excel file
    df = pd.read_excel(xls)
    df = df[df['IP Address'] != '-']

    if 'UpdatedValue' in df.columns :
        df = df.drop(columns=['UpdatedValue'])
        df['UpdatedValue'] = df['Value']
    else:
        df['UpdatedValue'] = df['Value']

    grouped = df.groupby('IP Address')

    # Iterate through each IP group and set up servers
    for ip_address, group in grouped:
        print(f"Setting up server for IP: {ip_address}")
        server = ModbusServer(ip_address, port, no_block=True )
        server.start()
        print(f"Modbus Server started at IP : {ip_address}")

        # Update point values in a loop
        print(f"Starting continuous value updates for IP: {ip_address}")
        while True:
            for _, row in group.iterrows():
                address = int(row['Index'])
                function_code = row['Function Code']
                value = df.loc[row.name, 'UpdatedValue'] if pd.notna(df.loc[row.name, 'UpdatedValue']) else 0
                mtype = row['Type']
                endian = row['Endian']
                
                if function_code == 2: # Binary Input signal
                    value_DI = [bool(int(value))]
                    server.data_bank.set_discrete_inputs(address - 1, value_DI)
                    print(f"Updated Discrete input {address} with value {value_DI}")
                    value = not bool(value)
                    df.loc[row.name, 'UpdatedValue'] = int(value)
                    time.sleep(1)

                elif function_code ==  5: # Binary Output signal
                    value = server.data_bank.get_coils(address - 1)
                    value_b = bool(value)
                    print(f" Read coil status at {address} value : {value}")
                    time.sleep(2)
                    df.loc[row.name, 'UpdatedValue'] = int(value)

                elif function_code == 3: # Analog Input signal
                    if mtype == 'Float':
                        valuen = float(value)
                        if endian == 'Big':
                            registers = float_to_registers_be(valuen)
                        elif endian == 'Little':
                            registers = float_to_registers(valuen)
                        else:
                            print("Invalid Endian Format")
                        server.data_bank.set_holding_registers(address - 1, registers)
                        print(f"Updated Holding input {address} with float value {value}")

                    elif mtype == 'Swapped Float':
                        valuen = float(value)
                        if endian == 'Big':
                            registers = float_to_registers_be(valuen)
                        elif endian == 'Little':
                            registers = float_to_registers(valuen)
                        else:
                            print("Invalid Endian Format")
                        registers_new = swap(registers)
                        server.data_bank.set_holding_registers(address - 1, registers_new)
                        print(f"Updated Holding input {address} with Swapped float value {value}")

                    elif mtype == '32bit unsigned Integer':
                        valuen = int(value)
                        if endian == 'Big':
                            try:
                                registers = unsigned_integer_to_register_be(valuen)
                            except ValueError as e:
                                print(f"Error {str(e)}")
                        elif endian == 'Little':
                            try:
                                registers = unsigned_integer_to_register(valuen)
                            except ValueError as e:
                                print(f"Error {str(e)}")
                        else:
                            print("Invalid Endian Format")
                        server.data_bank.set_holding_registers(address - 1, registers)
                        print(f"Updated Holding input {address} with Integer value {value}")

                    elif mtype == '32bit signed Integer':
                        valuen = int(value)
                        if endian == 'Big':
                            try:
                                registers = signed_integer_to_register_be(valuen)
                            except ValueError as e:
                                print(f"Error {str(e)}")
                        elif endian == 'Little':
                            try:
                                registers = signed_integer_to_register(valuen)
                            except ValueError as e:
                                print(f"Error {str(e)}")
                        else:
                            print("Invalid Endian Format")
                        server.data_bank.set_holding_registers(address - 1, registers)
                        print(f"Updated Holding input {address} with Integer value {value}")

                    elif mtype == '16bit unsigned Integer':
                        valuen = int(value)
                        if endian == 'Big':
                            try:
                                registers = unsigned_16bit_to_register_be(valuen)
                            except ValueError as e:
                                print(f"Error {str(e)}")
                        elif endian == 'Little':
                            try:
                                registers = unsigned_16bit_to_register(valuen)
                            except ValueError as e:
                                print(f"Error {str(e)}")
                        else:
                            print("Invalid Endian Format")
                        server.data_bank.set_holding_registers(address - 1, registers)
                        print(f"Updated Holding input {address} with Integer value {value}")

                    elif mtype == '16bit signed Integer':
                        valuen = int(value)
                        if endian == 'Big':
                            try:
                                registers = signed_16bit_to_register_be(valuen)
                            except ValueError as e:
                                print(f"Error {str(e)}")
                        elif endian == 'Little':
                            try:
                                registers = signed_16bit_to_register(valuen)
                            except ValueError as e:
                                print(f"Error {str(e)}")
                        else:
                            print("Invalid Endian Format")
                        server.data_bank.set_holding_registers(address - 1, registers)
                        print(f"Updated Holding input {address} with Integer value {value}")

                    time.sleep(1)
                    value = random.randint(10,100)
                    df.loc[row.name, 'UpdatedValue'] = value

                elif function_code == 6: # Analog Output signal
                    for i in range (3):
                        registers = server.data_bank.get_holding_registers(address -1, number = 1)

                        if mtype == 'Float':
                            if endian == 'Big':
                                value = registers_to_float_be(registers)
                            else :
                                value = registers_to_float(registers)
                            value = round(value,4)
                            print(f"Read holding register {address}, float value: {value}")
                            
                        elif mtype == 'Swapped Float':
                            registers_new  = swap(registers)
                            if endian == 'Big':
                                value = registers_to_float_be(registers_new)
                            else:
                                value = registers_to_float(registers_new)
                            value = round(value,4)
                            print(f"Read holding register {address}, Swapped float value: {value}")

                        elif mtype == '32bit unsigned Integer':
                            if endian == 'Big':
                                value = registers_to_unsigned_integer_be(registers)
                            else:
                                value = registers_to_unsigned_integer(registers)
                            print(f"Read holding register {address}, Integer value: {value}")

                        elif mtype == '32bit signed Integer':
                            if endian == 'Big':
                                value = registers_to_signed_integer_be(registers)
                            else:
                                value = registers_to_signed_integer(registers)
                            print(f"Read holding register {address}, Integer value: {value}")

                        elif mtype == '16bit unsigned Integer':
                            if endian == 'Big':
                                value = registers_to_unsigned_16bit_be(registers)
                            else:
                                value = registers_to_unsigned_16bit(registers)
                            print(f"Read holding register {address}, Integer value: {value}")

                        elif mtype == '16bit signed Integer':
                            if endian == 'Big':
                                value = registers_to_signed_16bit_be(registers)
                            else:
                                value = registers_to_signed_16bit(registers)
                            print(f"Read holding register {address}, Integer value: {value}")
                        
                        time.sleep(1)
                        df.loc[row.name, 'UpdatedValue'] = value
                    
                elif function_code == 16: # Analog Output signal
                    registers = server.data_bank.get_holding_registers(address -1, number = 2)

                    if mtype == 'Float':
                        if endian == 'Big':
                            value = registers_to_float_be(registers)
                        else :
                            value = registers_to_float(registers)
                        value = round(value,4)
                        print(f"Read holding register {address}, float value: {value}")
                        
                    elif mtype == 'Swapped Float':
                        registers_new  = swap(registers)
                        if endian == 'Big':
                            value = registers_to_float_be(registers_new)
                        else:
                            value = registers_to_float(registers_new)
                        value = round(value,4)
                        print(f"Read holding register {address}, Swapped float value: {value}")

                    elif mtype == '32bit unsigned Integer':
                        if endian == 'Big':
                            value = registers_to_unsigned_integer_be(registers)
                        else:
                            value = registers_to_unsigned_integer(registers)
                        print(f"Read holding register {address}, Integer value: {value}")

                    elif mtype == '32bit signed Integer':
                        if endian == 'Big':
                            value = registers_to_signed_integer_be(registers)
                        else:
                            value = registers_to_signed_integer(registers)
                        print(f"Read holding register {address}, Integer value: {value}")

                    elif mtype == '16bit unsigned Integer':
                        if endian == 'Big':
                            value = registers_to_unsigned_16bit_be(registers)
                        else:
                            value = registers_to_unsigned_16bit(registers)
                        print(f"Read holding register {address}, Integer value: {value}")

                    elif mtype == '16bit signed Integer':
                        if endian == 'Big':
                            value = registers_to_signed_16bit_be(registers)
                        else:
                            value = registers_to_signed_16bit(registers)
                        print(f"Read holding register {address}, Integer value: {value}")
                    
                    time.sleep(1)
                    df.loc[row.name, 'UpdatedValue'] = value
                else:
                    print(" Invalid Function code ")

            # Save the updated DataFrame back to Excel
            df.to_excel(file_path, index=False)
            print("Saved updated values to Excel.")
            time.sleep(5)  # Sleep for 4 seconds before the next update loop

if __name__ == "__main__":
    main()
