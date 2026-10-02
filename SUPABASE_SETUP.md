# Configuração do banco Supabase

O aplicativo usa PostgreSQL no Supabase para armazenar fichas, usuários, aprovações, correções e assinaturas. O arquivo `aeronaves.db` não é usado pelo aplicativo depois desta migração.

## Antes de publicar

1. No painel do Supabase, abra o projeto e copie a URI de conexão PostgreSQL. Use a URI do pooler de sessão ou a conexão direta, conforme a disponibilidade de rede do Streamlit Community Cloud.
2. No Streamlit Community Cloud, abra o aplicativo e acesse **Settings → Secrets**.
3. Adicione a URI no formato TOML abaixo. Substitua os valores de exemplo e use o prefixo `postgresql+psycopg://`:

   ```toml
   INITIAL_ADMIN_PASSWORD = "DEFINA_UMA_SENHA_FORTE"

   [connections.postgresql]
   url = "postgresql+psycopg://USUARIO:SENHA@HOST:PORTA/postgres?sslmode=require"
   ```

   Se o painel fornecer uma URI começando com `postgresql://`, altere apenas o prefixo para `postgresql+psycopg://`. Não coloque a senha no código, no Git ou no chat.
4. Salve os Secrets. O Streamlit reinicia o app; na primeira inicialização, as tabelas são criadas no PostgreSQL automaticamente.
5. Entre com o usuário inicial `engenharia` e a senha definida em `INITIAL_ADMIN_PASSWORD`. Em seguida, crie novamente os outros usuários e carregue suas assinaturas.

## Dados anteriores do SQLite

O app não importa automaticamente `aeronaves.db`, pois esse arquivo no repositório pode não conter os dados mais recentes do site. A migração começa com tabelas vazias no PostgreSQL. Recrie os logins, as assinaturas e as fichas necessárias no app. Depois disso, novos deploys de código não dependem do SQLite local.
