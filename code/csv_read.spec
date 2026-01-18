# -*- mode: python ; coding: utf-8 -*-
# Add these lines in the spec
import os
import pyvis

pyvis_path = os.path.dirname(pyvis.__file__)

a = Analysis(
    ['main_app.py'],
    pathex=[],
    binaries=[],
	datas = [
		("csv.ico", "."),
		(os.path.join(pyvis_path, "templates", "template.html"), "pyvis/templates"),
		(os.path.join(pyvis_path, "templates", "lib", "bindings", "utils.js"), "pyvis/templates/lib/bindings"),
		(os.path.join(pyvis_path, "templates", "lib", "tom-select", "tom-select.css"), "pyvis/templates/lib/tom-select"),
		# (os.path.join(pyvis_path, "templates", "lib", "tom-select", "tom-select.complete.min.js"), "pyvis/templates/lib/tom-select"),
		(os.path.join(pyvis_path, "templates", "lib", "vis-9.1.2", "vis-network.css"), "pyvis/templates/lib/vis-9.1.2"),
		# (os.path.join(pyvis_path, "templates", "lib", "vis-9.1.2", "vis-network.min.js"), "pyvis/templates/lib/vis-9.1.2"),
	],
	hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='MDTCSVEditor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='csv.ico',
)
