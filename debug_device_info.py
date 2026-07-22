
from pymobiledevice3.lockdown import create_using_usbmux
from pymobiledevice3.usbmux import list_devices

print("Buscando dispositivos...")
devices = list_devices()
if not devices:
    print("No se encontraron dispositivos.")
else:
    device = devices[0]
    print(f"Dispositivo encontrado: {device.serial}")
    try:
        lockdown = create_using_usbmux(serial=device.serial)
        
        print("\n--- ROOT DOMAIN VALUES ---")
        all_values = lockdown.get_value(key=None) # Get everything
        for k, v in all_values.items():
            if "Capacity" in k or "Color" in k:
                print(f"{k}: {v}")
                
        print("\n--- DISK USAGE DOMAIN ---")
        try:
            disk_values = lockdown.get_value(domain="com.apple.disk_usage")
            if disk_values:
                for k, v in disk_values.items():
                    print(f"{k}: {v}")
            else:
                print("com.apple.disk_usage devolvió None")
        except Exception as e:
            print(f"Error leyendo com.apple.disk_usage: {e}")

    except Exception as e:
        print(f"Error conectando: {e}")
