# Hospedagem no Hugging Face Spaces

O app roda num Space Docker gratuito do Hugging Face (2 vCPU, 16 GB de RAM).
O Space gratuito só dorme depois de 48 h sem nenhum acesso e acorda sozinho
quando alguém abre o link. O workflow `manter-acordado.yml` acessa o app a
cada 6 h, então na prática ele não dorme.

## Como o deploy funciona

- `Dockerfile` instala o LibreOffice (exportação em PDF) e as dependências.
- `start.sh` grava o conteúdo do secret `STREAMLIT_SECRETS` em
  `~/.streamlit/secrets.toml` e inicia o Streamlit.
- `.github/workflows/deploy-huggingface.yml` envia o código para o Space a
  cada push na `main` (sem `aeronaves.db`, `.xlsm` e pastas de dev).
- `deploy/huggingface/README.md` vira o README do Space (configuração Docker,
  porta 8501).

## Configuração (uma vez)

1. Crie uma conta em https://huggingface.co/join.
2. Crie o Space em https://huggingface.co/new-space:
   - Space name: `ficha-pesagem-gol`
   - SDK: **Docker** → template **Blank**
   - Hardware: **CPU basic (free)**
   - Visibility: **Public** (um Space privado exige login no Hugging Face
     para abrir o app; o acesso continua protegido pelo login do próprio app).
3. No Space, abra **Settings → Variables and secrets → New secret**:
   - Name: `STREAMLIT_SECRETS`
   - Value: cole o conteúdo inteiro dos Secrets do Streamlit Cloud
     (`[connections.postgresql]`, `url = ...`, `INITIAL_ADMIN_PASSWORD`...).
4. Crie um token em https://huggingface.co/settings/tokens → **Create new
   token** → tipo **Write** → copie o valor.
5. No GitHub do repositório, **Settings → Secrets and variables → Actions**:
   - aba **Secrets** → New repository secret: `HF_TOKEN` = o token do passo 4
   - aba **Variables** → New repository variable:
     `HF_SPACE` = `<seu-usuario-hf>/ficha-pesagem-gol`
   - aba **Variables** → New repository variable:
     `APP_URL` = `https://<seu-usuario-hf>-ficha-pesagem-gol.hf.space`
6. No GitHub, **Actions → Deploy no Hugging Face Spaces → Run workflow**.
   O primeiro build leva alguns minutos (instala o LibreOffice).
7. Use e compartilhe o link direto `APP_URL` (o endereço `.hf.space`), e não
   a página `huggingface.co/spaces/...`: dentro da página do Hugging Face o
   upload de arquivos (assinaturas) pode ser bloqueado pelo navegador.

Depois que tudo funcionar no novo endereço, o app no Streamlit Cloud pode ser
apagado em https://share.streamlit.io.

## Manter o Supabase ativo

O Supabase gratuito pausa o projeto depois de 7 dias sem uso, e aí o app
mostra "Não foi possível conectar ao banco de dados". Para evitar, cadastre no
GitHub (**Settings → Secrets and variables → Actions → Secrets**) o secret
`SUPABASE_DB_URL` com a mesma `url` de `[connections.postgresql]`. O workflow
`manter-acordado.yml` faz uma consulta no banco a cada 6 h.

## Observações

- O GitHub pausa workflows agendados depois de 60 dias sem nenhum commit no
  repositório. Se isso acontecer, o GitHub avisa por e-mail; basta reativar
  em **Actions → Manter o app acordado → Enable workflow**.
- O mesmo `Dockerfile` serve para Render ou Koyeb: lá, crie a variável de
  ambiente `STREAMLIT_SECRETS` com o mesmo conteúdo.
