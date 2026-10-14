FROM python:3.11-slim-bookworm

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

# O Hugging Face Spaces roda o container com o usuário de id 1000.
RUN useradd -m -u 1000 user
USER user
ENV HOME=/home/user \
    PATH=/home/user/.local/bin:$PATH \
    PYTHONUNBUFFERED=1
WORKDIR /home/user/app

COPY --chown=user requirements.txt ./requirements.txt
RUN pip install --no-cache-dir --user -r requirements.txt

COPY --chown=user . .
RUN chmod +x start.sh

EXPOSE 8501

CMD ["./start.sh"]
