@echo off
setlocal

cd /d C:\xampp\htdocs\projeto-aniversario
if not exist logs mkdir logs

>> logs\agendador_diretoria.txt echo.
>> logs\agendador_diretoria.txt echo [%date% %time%] Iniciando rascunho automatico diretoria
C:\Users\Suporte\anaconda3\python.exe C:\xampp\htdocs\projeto-aniversario\automacao\enviar_aniversarios.py --profile diretoria >> logs\agendador_diretoria.txt 2>&1
>> logs\agendador_diretoria.txt echo [%date% %time%] Finalizado com codigo %errorlevel%
