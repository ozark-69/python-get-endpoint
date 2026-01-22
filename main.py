import os
import json
import csv
from netmiko import (
    ConnectHandler,
    NetMikoAuthenticationException,
    NetMikoTimeoutException,
)
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

os.environ["NET_TEXTFSM"] = "ntc_templates"


def connect_to_device(creds):
    hostname = creds["hostname"]
    ip = creds["ip"]

    creds = {
        "device_type": creds["device_type"],
        "ip": creds["ip"],
        "username": creds["username"],
        "password": creds["password"],
        "secret": creds["password"],
        "fast_cli": False,
    }

    try:
        device = ConnectHandler(**creds)
        device.enable()
        return device

    except NetMikoTimeoutException:
        with open("connect_error.csv", "a") as file:
            file.write(f"{hostname};{ip};Device Unreachable/SSH not enabled")
        return None

    except NetMikoAuthenticationException:
        with open("connect_error.csv", "a") as file:
            file.write(f"{hostname};{ip};Authentication failure")
        return None
    except Exception as e:
        print(f"error: {e}")


# ADJUST CREDS
def load_devices(file="creds-dc.csv"):
    devices = []
    try:
        with open(file, "r") as f:
            reader = csv.reader(f, delimiter=";")

            for row in reader:
                hostname, ip, device_type, username, password = row

                devices.append(
                    {
                        "hostname": hostname,
                        "ip": ip,
                        "device_type": device_type,
                        "username": username,
                        "password": password,
                    }
                )

        return devices

    except FileNotFoundError:
        return []


def process_to_device(device):
    conn = connect_to_device(device)
    if conn:
        print(f"connected to {device.get('hostname', '')}")
        try:
            raw_show_int = conn.send_command(
                "show interface", use_textfsm=True, read_timeout=300, delay_factor=4
            )
            data = []
            for intf in raw_show_int:
                interface = intf.get("interface", "")  # type: ignore
                link_state = intf.get("link_status", intf.get("oper_state", ""))
                description = intf.get("description", "")
                ip = intf.get("ip_address", "")
                vlan = intf.get("vlan_id", "")

                if "admin_state" in intf:
                    admin_state = intf.get("admin_state", "")
                else:
                    # IOS: protocol_status is closest to admin state
                    admin_state = intf.get("protocol_status", "")

                data.append(
                    {
                        "interface": interface,
                        "link_state": link_state,
                        "admin_state": admin_state,
                        "description": description,
                        "vlan" : vlan,
                        "ip": ip,
                    }
                )
            return data
            

        except Exception as e:
            print(f"error {e}")
    else:
        print("Error")

def main():
    devices = load_devices()
    data = {}
    for device in devices:
        show_intf = process_to_device(device)
        if not show_intf:
            print("Skipping...")
            continue
        data[device.get("hostname")] = show_intf
    
    filename = "inventory_interfaces.json"
    try:
        with open(filename, "w", encoding="utf-8") as json_file:
            json.dump(data, json_file, indent=4)
        print(f"\n[SUCCESS] Data berhasil disimpan ke {filename}")
    except Exception as e:
        print(f"Gagal menyimpan file JSON: {e}")


if __name__ == "__main__":
    main()
