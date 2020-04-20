# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['petrel2.py'],
             pathex=['C:\\Users\\админ\\Desktop\\excel2las\\debug'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=['socketserver', 'math', 'browser', '_sqlite3', '_multiprocessing', '_asyncio', 'webbrowser', 'multiprocessing', 'sqlite3', 'asyncore', 'asynchat','asynsio', 'FixTk', 'tcl', 'tk', '_tkinter', 'tkinter', 'Tkinter'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='petrel2',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=['vcruntime140.dll'],
          runtime_tmpdir=None,
          console=False )
