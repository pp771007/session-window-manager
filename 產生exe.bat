@echo off
setlocal enabledelayedexpansion

rem ���Ĥ@��.pyw�ɮ�
for %%f in (*.pyw) do (
    set "py_file=%%f"
	echo �����pyhton�ɮ�: "!py_file!"
    goto :py_found
)

:py_found

rem �ˬd�O�_�s�b.ico�ɮ�
for %%i in (*.ico) do (
    set "custom_icon=%%i"
    goto :icon_found
)

:icon_found

rem �ϥ�pyinstaller�ͦ��i�����ɡA���w�ۭq�ϥ�
if defined custom_icon (
	echo �����ϥ���: "!custom_icon!"
    pyinstaller --onefile --noconsole --uac-admin --add-data "!custom_icon!;." --icon=!custom_icon! "!py_file!"
) else (
    rem �ϥ�pyinstaller�ͦ��i�����ɡA�����w�ϥ�
    pyinstaller --onefile --noconsole --uac-admin "!py_file!"
)

rem �M��ͦ����i������
for %%i in (dist\*.exe) do (
    move "%%i" .
)

rem �R��build��Ƨ��Bdist��Ƨ��M.spec�ɮ�
rd /s /q build
rd /s /q dist
del *.spec

endlocal
