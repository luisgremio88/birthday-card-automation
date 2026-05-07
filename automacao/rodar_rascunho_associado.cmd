@echo off
setlocal

cd /d C:\xampp\htdocs\projeto-aniversario
if not exist logs mkdir logs

>> logs\agendador_associado.txt echo.
>> logs\agendador_associado.txt echo [%date% %time%] Iniciando rascunho automatico associado
C:\Users\Suporte\anaconda3\python.exe C:\xampp\htdocs\projeto-aniversario\automacao\enviar_aniversarios.py --profile associado >> logs\agendador_associado.txt 2>&1
>> logs\agendador_associado.txt echo [%date% %time%] Finalizado com codigo %errorlevel%
