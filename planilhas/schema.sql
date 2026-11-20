CREATE TABLE IF NOT EXISTS planos(id INTEGER PRIMARY KEY, nome TEXT NOT NULL, descricao TEXT NOT NULL, preco INTEGER NOT NULL, prazo TEXT NOT NULL, destaque INTEGER NOT NULL DEFAULT 0, ordem INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS leads(
 id INTEGER PRIMARY KEY AUTOINCREMENT, protocolo TEXT UNIQUE, criado_em TEXT NOT NULL,
 nome TEXT NOT NULL, cpf TEXT NOT NULL, cnpj TEXT, nascimento TEXT NOT NULL, email TEXT NOT NULL,
 telefone TEXT NOT NULL, endereco TEXT NOT NULL, cidade TEXT NOT NULL, uf TEXT NOT NULL,
 plano_id INTEGER REFERENCES planos(id), descricao TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'Novo');
CREATE INDEX IF NOT EXISTS ix_leads_cpf ON leads(cpf);
CREATE INDEX IF NOT EXISTS ix_leads_data ON leads(criado_em);
CREATE VIEW IF NOT EXISTS vw_planilha AS
 SELECT l.protocolo, l.criado_em, l.nome, l.cpf, l.cnpj, l.nascimento, l.email, l.telefone, l.endereco, l.cidade, l.uf,
  COALESCE(p.nome,'Ainda não sei') AS plano, p.preco, l.descricao, l.status
 FROM leads l LEFT JOIN planos p ON p.id=l.plano_id ORDER BY l.id DESC;
INSERT OR IGNORE INTO planos VALUES
 (1,'Plano Básico','Retiramos os apontamentos do Serasa, Boa Vista, SPC, Cepront e Protestos.',700,'30 a 40 dias',0,1),
 (2,'Plano Intermediário','Retiramos os apontamentos, aumentamos o Score, aumentamos o Rating Bancário para nível A e atualizamos todos os dados no Serasa.',2200,'60 a 70 dias',0,2),
 (3,'Plano Completo','Todos os planos acima e retiramos também seu nome do SRC (Registrato ou BACEN).',3500,'60 a 90 dias',1,3),
 (4,'Retirada da lista negra (Registrato)','Removemos seu nome do Registrato e das principais bases de restrição.',2000,'Consulte o prazo',0,4),
 (5,'Aumento de score','Score com Rating Bancário e atualização de renda. Para quem tem nome limpo.',1500,'Consulte o prazo',0,5);