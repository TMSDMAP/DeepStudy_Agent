# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

project_dir = Path.cwd()

datas = [
    (str(project_dir / "app.py"), "."),
    (str(project_dir / "agents.py"), "."),
    (str(project_dir / "workflow.py"), "."),
    (str(project_dir / "state.py"), "."),
]

if (project_dir / ".env").exists():
    datas.append((str(project_dir / ".env"), "."))
if (project_dir / ".teacher_kb").exists():
    datas.append((str(project_dir / ".teacher_kb"), ".teacher_kb"))

binaries = []
hiddenimports = [
    "app",
    "workflow",
    "agents",
    "state",
]
for pkg in [
    "streamlit",
    "altair",
    "pydeck",
    "watchdog",
    "docx",
    "pptx",
]:
    try:
        pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(pkg)
    except Exception:
        continue
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hiddenimports


block_cipher = None

a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[str(project_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[str(project_dir / "hooks")],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='deepstudy_agent',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='deepstudy_agent',
)
