CREATE DATABASE IF NOT EXISTS sistema_medicacao
  CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE sistema_medicacao;

CREATE TABLE IF NOT EXISTS pacientes (
  id INT AUTO_INCREMENT PRIMARY KEY,
  nome VARCHAR(180) NOT NULL UNIQUE,
  ativo TINYINT(1) NOT NULL DEFAULT 1,
  criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS medicamentos (
  id INT AUTO_INCREMENT PRIMARY KEY,
  paciente_id INT NOT NULL,
  quantidade_texto VARCHAR(100) DEFAULT '',
  nome VARCHAR(255) NOT NULL,
  horario VARCHAR(40) NOT NULL DEFAULT 'NÃO INFORMADO',
  origem VARCHAR(255),
  conferido TINYINT(1) NOT NULL DEFAULT 0,
  criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT fk_med_paciente FOREIGN KEY (paciente_id) REFERENCES pacientes(id),
  UNIQUE KEY uq_med (paciente_id, quantidade_texto, nome, horario)
);

CREATE TABLE IF NOT EXISTS importacoes (
  id INT AUTO_INCREMENT PRIMARY KEY,
  arquivo VARCHAR(255) NOT NULL,
  tipo VARCHAR(40) NOT NULL,
  total_registros INT DEFAULT 0,
  importado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS documentos_origem (
  id INT AUTO_INCREMENT PRIMARY KEY,
  arquivo VARCHAR(255) NOT NULL,
  aba VARCHAR(180) NOT NULL,
  conteudo_json LONGTEXT NOT NULL,
  importado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- A checagem diária e o histórico serão adicionados na próxima etapa,
-- depois da validação das listas e dos horários com o responsável.
