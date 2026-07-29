"""
Subida automática a GitHub — Reportes PLD
==========================================
Sube automáticamente todos los archivos generados al repositorio
de GitHub Pages sin necesidad de abrir el navegador.

Uso:
    python subir_a_github.py

Solo necesitas configurar las 3 variables de la sección CONFIG.
"""

import subprocess
import sys
from pathlib import Path

# ══════════════════════════════════════════════════════════════
# CONFIG — edita solo estas 3 líneas
# ══════════════════════════════════════════════════════════════

GITHUB_TOKEN = "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"          # ← pega aquí tu token
GITHUB_USUARIO = "misaelu72-design"      # ← tu usuario de GitHub
REPO_NOMBRE = "reportes-pld"            # ← nombre del repositorio

# Carpeta donde están los archivos generados
CARPETA_SALIDA = Path(r"C:\Users\Misael Urbina\OneDrive - CECOM\Documentos\PORTAL WEB PLD\COMPARATIVAS")

# ══════════════════════════════════════════════════════════════
# ARCHIVOS A SUBIR — se detectan automáticamente
# ══════════════════════════════════════════════════════════════

def obtener_archivos():
    """Devuelve lista de (ruta_local, ruta_en_repo) de todos los archivos a subir."""
    archivos = []
    script_dir = Path(__file__).resolve().parent

    # ── HTML — NO se suben automáticamente ──────────────────────
    # index.html, home.html, dashboard.html, generador.html
    # se suben manualmente desde GitHub cuando haya cambios

    # ── JSON — datos y cortes ─────────────────────────────────
    for nombre in ["datos.json", "cortes.json", "usuarios.json"]:
        ruta = CARPETA_SALIDA / nombre
        if not ruta.exists():
            ruta = script_dir / nombre
        if ruta.exists():
            archivos.append((ruta, nombre))

    # ── JSON — historial de cortes (carpeta datos/) ───────────
    carpeta_datos = CARPETA_SALIDA / "datos"
    if carpeta_datos.exists():
        for json_file in sorted(carpeta_datos.glob("*.json")):
            archivos.append((json_file, f"datos/{json_file.name}"))

    # ── Excel — reporte consolidado ───────────────────────────
    for xlsx in sorted(CARPETA_SALIDA.glob("Reporte_PLD_TODAS_SUCURSALES_*.xlsx")):
        archivos.append((xlsx, xlsx.name))

    # ── Excel — detalle de cambios ────────────────────────────
    for xlsx in sorted(CARPETA_SALIDA.glob("Detalle_Cambios_Estatus_*.xlsx")):
        archivos.append((xlsx, xlsx.name))

    # ── Excel — verificados semana ────────────────────────────
    for xlsx in sorted(CARPETA_SALIDA.glob("Verificados_Semana_*.xlsx")):
        archivos.append((xlsx, xlsx.name))

    # ── EXCLUIDO: Base_Unificada (archivo muy pesado) ─────────
    # for xlsx in sorted(CARPETA_SALIDA.glob("Base_Unificada_*.xlsx")):
    #     archivos.append((xlsx, xlsx.name))

    return archivos


# ══════════════════════════════════════════════════════════════
# LÓGICA DE SUBIDA CON GIT
# ══════════════════════════════════════════════════════════════

def run(cmd, cwd=None, check=True):
    """Ejecuta un comando y muestra la salida."""
    result = subprocess.run(
        cmd, cwd=cwd, capture_output=True, text=True, shell=True
    )
    if result.stdout.strip():
        print(f"    {result.stdout.strip()}")
    if result.returncode != 0 and check:
        if result.stderr.strip():
            print(f"    ⚠ {result.stderr.strip()}")
    return result


def main():
    print("\n" + "="*55)
    print("  SUBIDA AUTOMÁTICA A GITHUB — Reportes PLD")
    print("="*55 + "\n")

    if GITHUB_TOKEN == "TU_TOKEN_AQUI":
        print("❌ Falta configurar el token de GitHub.")
        print("   Abre este script y reemplaza TU_TOKEN_AQUI con tu token.")
        input("\nPresiona Enter para cerrar...")
        sys.exit(1)

    # Carpeta temporal para el repositorio
    repo_dir = Path.home() / "reportes_pld_git"
    repo_url = f"https://{GITHUB_TOKEN}@github.com/{GITHUB_USUARIO}/{REPO_NOMBRE}.git"

    # 1. Clonar o actualizar el repositorio
    if not repo_dir.exists():
        print("📥 Clonando repositorio por primera vez...")
        result = run(f'git clone "{repo_url}" "{repo_dir}"', check=False)
        if result.returncode != 0:
            print("❌ Error al clonar. Verifica tu token y usuario.")
            input("\nPresiona Enter para cerrar...")
            sys.exit(1)
        print("  ✅ Repositorio clonado")
    else:
        print("🔄 Actualizando repositorio local...")
        run(f'git -C "{repo_dir}" pull origin main', check=False)
        print("  ✅ Actualizado")

    # 2. Copiar archivos al repositorio local
    archivos = obtener_archivos()
    if not archivos:
        print("❌ No se encontraron archivos para subir.")
        print(f"   Verifica que CARPETA_SALIDA existe: {CARPETA_SALIDA}")
        input("\nPresiona Enter para cerrar...")
        sys.exit(1)

    print(f"\n📂 Copiando {len(archivos)} archivo(s) al repositorio...\n")
    import shutil
    for ruta_local, ruta_repo in archivos:
        destino = repo_dir / ruta_repo
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ruta_local, destino)
        print(f"  ✅ {ruta_repo}")

    # 3. Configurar git si es la primera vez
    run(f'git -C "{repo_dir}" config user.email "pld@cecom.com"', check=False)
    run(f'git -C "{repo_dir}" config user.name "Reportes PLD CECOM"', check=False)

    # 4. Agregar todos los archivos
    print("\n📤 Preparando commit...")
    run(f'git -C "{repo_dir}" add -A')

    # 5. Verificar si hay cambios
    status = run(f'git -C "{repo_dir}" status --porcelain', check=False)
    if not status.stdout.strip():
        print("  ℹ️  No hay cambios nuevos — el repositorio ya está actualizado.")
        input("\nPresiona Enter para cerrar...")
        return

    # 6. Commit con fecha automática
    from datetime import datetime
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    mensaje = f"Actualización reporte PLD — {fecha}"
    run(f'git -C "{repo_dir}" commit -m "{mensaje}"')
    print(f"  ✅ Commit: {mensaje}")

    # 7. Push a GitHub
    print("\n🚀 Subiendo a GitHub...")
    result = run(
        f'git -C "{repo_dir}" push "{repo_url}" main',
        check=False
    )
    if result.returncode == 0:
        print("  ✅ Subida exitosa")
    else:
        # Intentar con 'master' si 'main' falla
        result2 = run(
            f'git -C "{repo_dir}" push "{repo_url}" master',
            check=False
        )
        if result2.returncode == 0:
            print("  ✅ Subida exitosa")
        else:
            print(f"  ❌ Error al subir: {result.stderr.strip()}")
            input("\nPresiona Enter para cerrar...")
            sys.exit(1)

    print("\n" + "="*55)
    print("  ✅ PÁGINA ACTUALIZADA")
    print(f"  🌐 https://{GITHUB_USUARIO}.github.io/{REPO_NOMBRE}/")
    print("="*55)
    print("\nEspera 1-2 minutos y recarga la página con Ctrl+Shift+R\n")
    input("Presiona Enter para cerrar...")


if __name__ == "__main__":
    main()
