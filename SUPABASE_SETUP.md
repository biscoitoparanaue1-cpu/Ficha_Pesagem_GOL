# Configuração do banco Supabase

O aplicativo usa PostgreSQL no Supabase para armazenar fichas, usuários, aprovações, correções e assinaturas. O arquivo `aeronaves.db` não é usado pelo aplicativo depois desta migração.

## Antes de publicar

1. No painel do Supabase, abra o projeto e copie a URI de conexão PostgreSQL. Use a URI do pooler de sessão ou a conexão direta, conforme a disponibilidade de rede do Streamlit Community Cloud.
2. No Streamlit Community Cloud, abra o aplicativo e acesse **Settings → Secrets**.
3. Adicione a URI completa copiada do Supabase no formato TOML abaixo. Use o prefixo `postgresql+psycopg://`:

   ```toml
   INITIAL_ADMIN_PASSWORD = "DEFINA_UMA_SENHA_FORTE"

   [connections.postgresql]
   url = "postgresql+psycopg://postgres:SENHA_CODIFICADA@db.ytfkffyrcrycvpqtxmud.supabase.co:5432/postgres?sslmode=require"
   ```

   Substitua `SENHA_CODIFICADA` pela senha real do banco, codificando caracteres especiais para uso em URL. Não copie esse marcador literalmente nem coloque a senha no código, no Git ou no chat. Se o painel fornecer uma URI começando com `postgresql://`, altere apenas o prefixo para `postgresql+psycopg://`.
4. Salve os Secrets. O Streamlit reinicia o app; na primeira inicialização, as tabelas são criadas no PostgreSQL automaticamente.
5. Entre com o usuário inicial `engenharia` e a senha definida em `INITIAL_ADMIN_PASSWORD`. Em seguida, crie novamente os outros usuários e carregue suas assinaturas.

## Importar dados do SQLite para o Supabase

O app não importa `aeronaves.db` automaticamente. Para executar uma importação única, use `scripts/migrate_sqlite_to_postgres.py` em um computador que tenha acesso de rede ao Supabase e ao arquivo SQLite escolhido. Não envie `aeronaves.db` ao Git: ele pode conter dados operacionais, hashes de senha e assinaturas.

1. Redefina a senha do banco no Supabase se ela já foi compartilhada. No ambiente que executará a migração, instale as dependências do projeto com `pip install -r requirements.txt`.
2. No terminal, informe a URI sem exibi-la e valide a conexão e o arquivo de origem:

   ```bash
   read -rsp "URI do PostgreSQL: " SUPABASE_DATABASE_URL; printf '\n'
   export SUPABASE_DATABASE_URL
   python scripts/migrate_sqlite_to_postgres.py --source aeronaves.db
   ```

   Sem `--apply`, o script apenas valida e não grava dados.
   Se a validação falhar, não use `--apply`: confira o detalhe do erro. Erros
   como `could not translate host name` indicam host incorreto; `timeout` ou
   `connection refused` geralmente indicam rede, porta ou pooler; `password
   authentication failed` indica credencial incorreta. Não compartilhe a URI
   nem a senha. O detalhe impresso pelo script omite credenciais.
3. Se a validação confirmar a origem e a conexão, execute a importação:

   ```bash
   python scripts/migrate_sqlite_to_postgres.py --source aeronaves.db --apply
   unset SUPABASE_DATABASE_URL
   ```

   A operação é transacional, importa fichas e os dados relacionados, preserva registros conflitantes nas tabelas auxiliares e cancela se o Supabase já tiver fichas, evitando duplicar ou misturar históricos. Em caso de erro, a transação é revertida. Depois, reinicie o app para atualizar o histórico em cache.

4. No Streamlit Community Cloud, confirme em **Settings → Secrets** que
   `[connections.postgresql].url` aponta para o mesmo banco/projeto Supabase
   usado na importação. A variável `SUPABASE_DATABASE_URL` definida no
   terminal de migração não configura os Secrets do app publicado. Salve os
   Secrets e reinicie o app; se a consulta continuar vazia, confira novamente
   o host/projeto configurado sem compartilhar a URI.

Se não for possível conectar, execute os comandos em uma rede que alcance o host PostgreSQL configurado. Não cole a URI nem a senha no chat.
