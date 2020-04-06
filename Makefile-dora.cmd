pyinstaller -F -w -i dora-128x128.ico dora.py
move dist\dora.exe .
rd/s/q __pycache__ build dist
del dora.spec
pause