# InspecFlow

Sistema de gestão de inspeções. Frontend em Next.js + banco/autenticação/armazenamento no Supabase.

## Variáveis de ambiente (configurar na Vercel)

- `NEXT_PUBLIC_SUPABASE_URL` — URL do projeto Supabase (ex: https://xxxx.supabase.co)
- `NEXT_PUBLIC_SUPABASE_ANON_KEY` — chave "anon public" do projeto Supabase

## Depois do primeiro deploy

1. Em Supabase → Authentication → Users → **Add user**, crie um usuário (e-mail + senha) para você, o ADM.
2. Copie o **UID** desse usuário e rode no SQL Editor:
   ```sql
   insert into admins (user_id) values ('COLE-O-UID-AQUI');
   ```
3. Para cada empresa: crie outro usuário em Authentication → Users, copie o UID, e cole no campo
   "ID do usuário (Supabase Auth)" ao editar a empresa dentro do próprio sistema (tela Empresas).
4. Pronto — cada empresa loga com o e-mail/senha que você criou para ela.
