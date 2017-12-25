FOR %G IN (extract split fillin merge) DO (
	pyinstaller --nowindow %G.py
	pyinstaller %G.spec)
