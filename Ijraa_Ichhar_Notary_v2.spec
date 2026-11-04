# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['desktop_app/main_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('desktop_app/docx_templates', 'docx_templates')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'torch', 'torchvision', 'paddle', 'cv2', 'scipy', 'numpy', 'pandas', 'sklearn',
        'matplotlib', 'seaborn', 'h5py', 'pyarrow', 'xarray', 'netCDF4', 'llvmlite',
        'numba', 'faiss', 'faiss_cpu', 'pygame', 'IPython', 'pytest', 'grpc', 'botocore',
        'boto3', 'setuptools', 'pip', 'distutils', 'pkg_resources', 'pydantic',
        'docx', 'lxml', 'PIL', 'Pillow', 'fitz', 'pymupdf'
    ],
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
    name='Ijraa_Ichhar_Notary_v2',
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
)
