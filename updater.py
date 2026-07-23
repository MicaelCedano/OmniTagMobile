# -*- coding: utf-8 -*-
"""
OmniTag Updater - Helper Silencioso de Actualizaciones
Autor: Micael Cedano
"""
import sys
import os
import time
import subprocess
import argparse

def main():
    parser = argparse.ArgumentParser(description="OmniTag Auto-Updater")
    parser.add_argument("--new", required=True, help="Ruta al nuevo ejecutable descargado")
    parser.add_argument("--target", required=True, help="Ruta al ejecutable actual a reemplazar")
    args = parser.parse_args()

    new_exe = os.path.abspath(args.new)
    target_exe = os.path.abspath(args.target)
    target_name = os.path.basename(target_exe)

    # 1. Esperar a que OmniTagMobile.exe cierre completamente
    for _ in range(30):
        time.sleep(0.4)
        try:
            res = subprocess.run(
                f'tasklist /FI "IMAGENAME eq {target_name}" 2>nul',
                shell=True, capture_output=True, text=True
            )
            if target_name.lower() not in res.stdout.lower():
                break
        except Exception: pass

    time.sleep(0.8)

    # 2. Reemplazar ejecutable objetivo
    success = False
    for attempt in range(15):
        try:
            if os.path.exists(target_exe):
                try: os.remove(target_exe)
                except Exception: pass
            
            with open(new_exe, "rb") as f_in, open(target_exe, "wb") as f_out:
                f_out.write(f_in.read())
            success = True
            break
        except Exception as e:
            time.sleep(0.8)

    if not success:
        try:
            import shutil
            shutil.copy2(new_exe, target_exe)
            success = True
        except Exception: pass

    # 3. Eliminar instalador temporal
    try:
        if os.path.exists(new_exe):
            os.remove(new_exe)
    except Exception: pass

    # 4. Volver a iniciar la aplicación principal
    try:
        subprocess.Popen([target_exe])
    except Exception as e:
        print(f"Error iniciando aplicación: {e}")

if __name__ == "__main__":
    main()
