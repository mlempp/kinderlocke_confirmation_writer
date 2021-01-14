# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['BestellDrucker.py'],
             pathex=['C:\\Users\\Martin\\Documents\\_Arbeit\\_side_projects\\5_kinderlocke_ConfirmationWriter\\BestellDrucker_V.0.2.2b'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

a.datas += [('highlight1_logo.png','C:\\Users\\Martin\\Documents\\_Arbeit\\_side_projects\\5_kinderlocke_ConfirmationWriter\\BestellDrucker_V.0.2.2\\highlight1_logo.png', 'Data')]
a.datas += [('Rechnung_Vorlage_V_0.2.docx','C:\\Users\\Martin\\Documents\\_Arbeit\\_side_projects\\5_kinderlocke_ConfirmationWriter\\BestellDrucker_V.0.2.2\\Rechnung_Vorlage_V_0.2.docx', 'Data')]

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='BestellDrucker',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='highlight1_logo.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='BestellDrucker')
