pyinstaller -F -w -i main-128x128.ico main.py
move dist\main.exe .
rd/s/q __pycache__ build dist
del main.spec
pause