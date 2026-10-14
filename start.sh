#!/bin/sh
# Inicia o app em container (Hugging Face Spaces, Render, Koyeb...).
# STREAMLIT_SECRETS recebe o conteúdo inteiro do secrets.toml do Streamlit
# Cloud e é gravado onde o Streamlit procura os secrets.
set -e

if [ -n "$STREAMLIT_SECRETS" ]; then
    mkdir -p "$HOME/.streamlit"
    printf '%s\n' "$STREAMLIT_SECRETS" > "$HOME/.streamlit/secrets.toml"
    chmod 600 "$HOME/.streamlit/secrets.toml"
fi

exec streamlit run app.py \
    --server.address=0.0.0.0 \
    --server.port="${PORT:-8501}" \
    --server.headless=true
