
try:
    from pymobiledevice3.lockdown import create_using_usbmux
    print("\ncreate_using_usbmux found in lockdown!")
except ImportError:
    print("\ncreate_using_usbmux NOT found in lockdown.")

try:
    from pymobiledevice3.usbmux import create_using_usbmux
    print("\ncreate_using_usbmux found in usbmux!")
except ImportError:
    print("\ncreate_using_usbmux NOT found in usbmux.")
