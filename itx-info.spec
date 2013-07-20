# -*- mode: python -*-
a = Analysis(['itx-info.py'],
             pathex=['C:\\Users\\459662\\Desktop\\ITXInfo'],
             hiddenimports=[],
             hookspath=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name=os.path.join('dist', 'itx-info.exe'),
          debug=False,
          strip=None,
          upx=True,
          console=False )
app = BUNDLE(exe,
             name=os.path.join('dist', 'itx-info.exe.app'))
