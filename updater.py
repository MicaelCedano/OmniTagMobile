# -*- coding: utf-8 -*-
"""
OmniTag Updater - Helper Silencioso de Actualizaciones (v2 - Batch Helper Pattern)
Autor: Micael Cedano / Kiwi

Usa el patrón batch helper + VBS validado en MCTools:
1. Espera a que el proceso principal cierre
2. Reemplaza el .exe con move /Y (atómico, sin locks)
3. Relanza la app
4. Se auto-elimina

Uso:
    updater.exe --new <ruta_nuevo_exe> --target <ruta_exe_actual>
"""
import sys
import os
import time
import subprocess
import tempfile

def main():
    if "--new" in sys.argv and "--target" in sys.argv:
        # Modo CLI: lanzado desde omnitag_mobile.py
        new_idx = sys.argv.index("--new")
        target_idx = sys.argv.index("--target")
        new_exe = os.path.abspath(sys.argv[new_idx + 1])
        target_exe = os.path.abspath(sys.argv[target_idx + 1])

        # Crear batch helper
        exe_name = os.path.basename(target_exe)
        exe_dir = os.path.dirname(target_exe)
        exe_short = exe_name.replace('.exe', '')

        bat_path = os.path.join(exe_dir, "_update_helper.bat")

        bat_lines = [
            "@echo off",
            "chcp 65001 >nul 2>&1",
            "",
            ":wait",
            "ping 127.0.0.1 -n 2 >nul",
            'tasklist /FI "IMAGENAME eq ' + exe_name + '" 2>nul | find /I "' + exe_short + '" >nul',
            "if not errorlevel 1 goto wait",
            "",
            'move /Y "' + new_exe + '" "' + target_exe + '" >nul 2>&1',
            "",
            'start "" "' + target_exe + '"',
            "",
            '(goto) 2>nul & del "%~f0" >nul 2>&1',
        ]
        bat_content = "\r\n".join(bat_lines)

        with open(bat_path, 'w', newline='\r\n') as f:
            f.write(bat_content)

        # Crear VBS para lanzar el batch oculto (con soporte para rutas con espacios)
        vbs_path = os.path.join(exe_dir, "_update_helper.vbs")
        vbs_code = 'CreateObject("WScript.Shell").Run chr(34) & "' + bat_path + '" & chr(34), 0, False'

        with open(vbs_path, 'w', newline='\r\n') as f:
            f.write(vbs_code)


        # Lanzar VBS con wscript.exe (invisible, sin consola)
        subprocess.Popen(
            ['wscript.exe', vbs_path],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
            close_fds=True
        )

        # El batch helper se encarga del resto. Salir limpiamente.
        sys.exit(0)

    else:
        print("Uso: updater.exe --new <ruta_nuevo> --target <ruta_actual>")
        sys.exit(1)

if __name__ == "__main__":
    main()
