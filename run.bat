@echo off
chcp 65001 > nul
echo.
echo  ================================================
echo     DOLE 2028 — Application de Gestion FSCF
echo  ================================================
echo.
echo  Installation des dependances en cours...
pip install -r requirements.txt --quiet
echo  OK.
echo.
echo  Demarrage de l'application...
echo  Ouvrez votre navigateur sur : http://localhost:5000
echo  Connexion admin : admin@dole2028.fr / admin2028
echo.
python app.py
pause
