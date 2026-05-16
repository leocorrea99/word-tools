-- ============================================================
-- Word Tools — Setup do banco de licenças no Supabase
-- Execute este script no SQL Editor do Supabase
-- ============================================================

CREATE TABLE IF NOT EXISTS licenses (
  id           UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  email        TEXT        NOT NULL UNIQUE,
  license_key  TEXT        NOT NULL UNIQUE DEFAULT replace(gen_random_uuid()::text, '-', ''),
  plan         TEXT        NOT NULL DEFAULT 'free',  -- 'free' | 'pro'
  active       BOOLEAN     NOT NULL DEFAULT true,
  created_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
  last_check   TIMESTAMPTZ
);

-- Índice para buscas por license_key (feitas pelo app)
CREATE INDEX IF NOT EXISTS idx_licenses_key ON licenses (license_key);

-- Row Level Security
ALTER TABLE licenses ENABLE ROW LEVEL SECURITY;

-- Qualquer pessoa pode se cadastrar (LP)
CREATE POLICY "insert_anon" ON licenses
  FOR INSERT WITH CHECK (true);

-- App pode ler para validar a chave
CREATE POLICY "select_anon" ON licenses
  FOR SELECT USING (true);

-- Apenas o admin (service role) pode atualizar/deletar
-- (você faz isso pelo dashboard do Supabase)
