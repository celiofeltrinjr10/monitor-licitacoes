@echo off
cd /d "%~dp0"

:MENU
cls
echo =====================================================
echo   MONITOR DE LICITACOES - S.A. Paulista
echo =====================================================
echo.
echo   1. Buscar licitacoes (ultimos 9 dias)
echo   2. Buscar licitacoes (definir periodo)
echo   3. Simulacao sem salvar (dry-run)
echo   4. Abrir app web local (Streamlit)
echo   5. Publicar atualizacoes no GitHub
echo   6. Abrir app online no navegador
echo   7. Ver planilha de resultados
echo   0. Sair
echo.
set /p OPCAO=   Escolha uma opcao: 

if "%OPCAO%"=="1" goto BUSCA_PADRAO
if "%OPCAO%"=="2" goto BUSCA_CUSTOM
if "%OPCAO%"=="3" goto DRY_RUN
if "%OPCAO%"=="4" goto APP_LOCAL
if "%OPCAO%"=="5" goto GIT_PUSH
if "%OPCAO%"=="6" goto APP_ONLINE
if "%OPCAO%"=="7" goto PLANILHA
if "%OPCAO%"=="0" goto FIM
echo   Opcao invalida.
timeout /t 2 >nul
goto MENU


:BUSCA_PADRAO
cls
echo Iniciando busca (ultimos 9 dias)...
echo.
python pncp_scraper.py
echo.
echo Concluido. Pressione qualquer tecla para voltar ao menu.
pause >nul
goto MENU


:BUSCA_CUSTOM
cls
set /p DIAS=   Quantos dias atras buscar? 
echo Iniciando busca (ultimos %DIAS% dias)...
echo.
python pncp_scraper.py --dias %DIAS%
echo.
echo Concluido. Pressione qualquer tecla para voltar ao menu.
pause >nul
goto MENU


:DRY_RUN
cls
echo Modo simulacao - nenhum dado sera salvo na planilha.
echo.
set /p DIAS_DR=   Quantos dias atras? (Enter = 9): 
if "%DIAS_DR%"=="" set DIAS_DR=9
python pncp_scraper.py --dias %DIAS_DR% --dry-run
echo.
echo Pressione qualquer tecla para voltar ao menu.
pause >nul
goto MENU


:APP_LOCAL
cls
echo Iniciando Streamlit localmente...
echo Abra http://localhost:8501 no navegador.
echo Para parar, pressione Ctrl+C nesta janela.
echo.
streamlit run app.py
goto MENU


:GIT_PUSH
cls
echo ---- Publicando atualizacoes no GitHub ----
echo.

REM Localiza o git nos caminhos mais comuns de instalacao
set GIT_EXE=git
if exist "C:\Program Files\Git\cmd\git.exe" set GIT_EXE="C:\Program Files\Git\cmd\git.exe"
if exist "C:\Program Files (x86)\Git\cmd\git.exe" set GIT_EXE="C:\Program Files (x86)\Git\cmd\git.exe"
if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" set GIT_EXE="%LOCALAPPDATA%\Programs\Git\cmd\git.exe"
if exist "%USERPROFILE%\AppData\Local\GitHubDesktop\app-3.4.3\resources\app\git\cmd\git.exe" set GIT_EXE="%USERPROFILE%\AppData\Local\GitHubDesktop\app-3.4.3\resources\app\git\cmd\git.exe"

%GIT_EXE% add .
echo.
set /p MSG=   Descricao da mudanca (ex: atualiza palavras-chave):
%GIT_EXE% commit -m "%MSG%"
echo.
%GIT_EXE% push --force origin main
echo.
echo Pronto! O app online sera atualizado em ~1 minuto.
echo Pressione qualquer tecla para voltar ao menu.
pause >nul
goto MENU


:APP_ONLINE
start https://celiofeltrinjr10-monitor-licitacoes-app.streamlit.app
goto MENU


:PLANILHA
start controle_licitacoes_obras.xlsx
goto MENU


:FIM
exit
