import os
import sys
import shutil
import subprocess
from pathlib import Path

def run_cmd(cmd):
    print(f"\n[EXECUTANDO] {cmd}")
    result = subprocess.run(cmd, shell=True)
    if result.returncode != 0:
        print(f"\n[ERRO] O comando falhou com código {result.returncode}")
        return False
    return True

def main():
    base_dir = Path(__file__).parent.resolve()
    os.chdir(base_dir)
    print(f"=== INICIANDO BUILD DO AGREGADOR DE PLANILHAS ===")
    print(f"Diretório do Projeto: {base_dir}")

    # 1. Limpar pastas de builds antigas
    for folder in ["build", "dist"]:
        p = base_dir / folder
        if p.exists():
            print(f"-> Apagando pasta antiga '{folder}'...")
            try:
                shutil.rmtree(p)
            except Exception as e:
                print(f"   Aviso ao apagar {folder}: {e}")

    # 2. Identificar executável do Python correto (da venv se existir, ou do sistema)
    venv_python = base_dir / "venv" / "Scripts" / "python.exe"
    if venv_python.exists():
        py_exe = f'"{venv_python}"'
        print(f"-> Usando Python da venv: {venv_python}")
    else:
        py_exe = f'"{sys.executable}"'
        print(f"-> Usando Python do sistema: {sys.executable}")

    # 3. Garantir instalação/atualização das dependências
    print("\n-> Verificando e instalando dependências...")
    install_cmd = f'{py_exe} -m pip install --upgrade pip pandas customtkinter openpyxl python-calamine pyinstaller'
    if not run_cmd(install_cmd):
        print("[ERRO] Falha ao instalar dependências.")
        input("Pressione ENTER para sair...")
        return

    # 4. Executar o PyInstaller com todas as flags de compatibilidade e hidden imports
    pyinstaller_cmd = (
        f'{py_exe} -m PyInstaller '
        f'--noconfirm '
        f'--onefile '
        f'--windowed '
        f'--hidden-import=pandas '
        f'--hidden-import=customtkinter '
        f'--hidden-import=openpyxl '
        f'--hidden-import=calamine '
        f'--hidden-import=PIL '
        f'--hidden-import=tkinter '
        f'--collect-all customtkinter '
        f'--name "AgregadorPlanilhas" '
        f'app.py'
    )

    if run_cmd(pyinstaller_cmd):
        exe_path = base_dir / "dist" / "AgregadorPlanilhas" / "AgregadorPlanilhas.exe"
        print("\n=======================================================")
        print("  SUCESSO! O EXECUTÁVEL FOI GERADO COM SUCESSO! 🚀")
        print(f"  Local do .exe: {exe_path}")
        print("=======================================================\n")
    else:
        print("\n[ERRO] Ocorreu uma falha na compilação do PyInstaller.")

    input("Pressione ENTER para fechar...")

if __name__ == "__main__":
    main()
