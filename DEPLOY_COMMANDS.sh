#!/bin/bash
# Script de comandos para conectar ao GitHub e fazer deploy

# IMPORTANTE: Substitua SEU-USUARIO pelo seu nome de usuário do GitHub!

# 1. Adicionar repositório remoto
git remote add origin https://github.com/SEU-USUARIO/projeto-analise-emprestimo.git

# 2. Renomear branch para main
git branch -M main

# 3. Fazer push para o GitHub
git push -u origin main

# Pronto! Agora acesse https://share.streamlit.io para fazer o deploy
