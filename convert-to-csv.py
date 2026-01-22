import json
import csv

def convert_json_to_csv(input_filename, output_filename):
    try:
        # 1. Load data dari file JSON
        with open(input_filename, 'r') as f:
            data = json.load(f)

        # 2. Persiapkan header untuk CSV
        # Kita tambahkan 'hostname' agar tahu interface ini milik perangkat mana
        header = ["hostname", "interface", "link_state", "admin_state", "vlan", "ip", "description"]

        # 3. Tulis ke file CSV
        with open(output_filename, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=header, delimiter=';')
            writer.writeheader()

            for hostname, interfaces in data.items():
                for intf in interfaces:
                    # Menambahkan info hostname ke dalam dictionary interface
                    intf['hostname'] = hostname
                    
                    # Menulis baris ke CSV (hanya mengambil kunci yang ada di header)
                    writer.writerow({k: intf.get(k, "") for k in header})

        print(f"Berhasil! Data telah dikonversi ke: {output_filename}")

    except FileNotFoundError:
        print(f"Error: File {input_filename} tidak ditemukan.")
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

if __name__ == "__main__":
    # Sesuaikan dengan nama file hasil script sebelumnya
    convert_json_to_csv("inventory_interfaces.json", "inventory_interfaces.csv")