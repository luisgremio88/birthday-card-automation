@echo off
setlocal

cd /d C:\xampp\htdocs\projeto-aniversario
if not exist logs mkdir logs

echo Teste de configuracao - Associado
C:\Users\Suporte\anaconda3\python.exe C:\xampp\htdocs\projeto-aniversario\automacao\testar_configuracao.py --profile associado
echo.
echo Teste de configuracao - Diretoria
C:\Users\Suporte\anaconda3\python.exe C:\xampp\htdocs\projeto-aniversario\automacao\testar_configuracao.py --profile diretoria
echo.
pause
