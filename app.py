
import os
import json

from flask import (
    Flask, render_template, request,
    redirect, url_for, flash
)
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from database import conectar
from config import Config
from datetime import datetime


app = Flask(__name__)
app.config.from_object(Config)

UPLOAD_DIR = os.path.join(app.root_path, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

HORARIOS = {"MANHÃ", "ALMOÇO", "JANTAR", "NOITE", "S/N"}
STATUS_VALIDOS = {"Em uso", "Suspenso", "Encerrado"}


# ==================================================
# IMPORTAÇÃO DA PLANILHA DE CHECAGEM
# ==================================================

def importar_checagem(caminho, nome_arquivo):
    wb = load_workbook(caminho, data_only=True, read_only=True)
    cn = conectar()
    cur = cn.cursor()
    inseridos = 0

    for aba in wb.worksheets:
        nome = str(aba.cell(1, 5).value or aba.title).strip()

        if not nome:
            nome = aba.title

        cur.execute(
            """
            INSERT INTO pacientes (nome)
            VALUES (%s)
            ON DUPLICATE KEY UPDATE id=LAST_INSERT_ID(id)
            """,
            (nome,)
        )

        paciente_id = cur.lastrowid
        horario = "NÃO INFORMADO"

        for linha in range(3, aba.max_row + 1):
            qtd = aba.cell(linha, 1).value
            medicamento = aba.cell(linha, 2).value

            if not medicamento:
                continue

            texto = str(medicamento).strip()
            normalizado = texto.upper().strip()

            if normalizado in HORARIOS:
                horario = normalizado
                continue

            if normalizado in {
                "MEDICAÇÕES", "MEDICACOES", "QT", "OBS"
            }:
                continue

            cur.execute(
                """
                INSERT INTO medicamentos
                (
                    paciente_id,
                    quantidade_texto,
                    nome,
                    horario,
                    origem,
                    conferido
                )
                VALUES (%s,%s,%s,%s,%s,0)
                ON DUPLICATE KEY UPDATE id=LAST_INSERT_ID(id)
                """,
                (
                    paciente_id,
                    str(qtd).strip() if qtd is not None else "",
                    texto,
                    horario,
                    nome_arquivo
                )
            )

            inseridos += 1

    cur.execute(
        """
        INSERT INTO importacoes
        (arquivo, tipo, total_registros)
        VALUES (%s, 'checagem', %s)
        """,
        (nome_arquivo, inseridos)
    )

    cn.commit()
    cur.close()
    cn.close()
    wb.close()

    return inseridos


# ==================================================
# ARQUIVAMENTO DAS PLANILHAS DE RECEITUÁRIO
# ==================================================

def arquivar_receitas(caminho, nome_arquivo):
    wb = load_workbook(caminho, data_only=True, read_only=True)

    cn = conectar()
    cur = cn.cursor()
    total = 0

    for aba in wb.worksheets:
        dados = []

        for row in aba.iter_rows():
            valores = [c.value for c in row]

            if any(v is not None for v in valores):
                dados.append(valores)

        cur.execute(
            """
            INSERT INTO documentos_origem
            (arquivo, aba, conteudo_json)
            VALUES (%s,%s,%s)
            """,
            (
                nome_arquivo,
                aba.title,
                json.dumps(
                    dados,
                    ensure_ascii=False,
                    default=str
                )
            )
        )

        total += 1

    cur.execute(
        """
        INSERT INTO importacoes
        (arquivo, tipo, total_registros)
        VALUES (%s, 'receitas_arquivo', %s)
        """,
        (nome_arquivo, total)
    )

    cn.commit()
    cur.close()
    cn.close()
    wb.close()

    return total


# ==================================================
# PAINEL INICIAL
# ==================================================

@app.route("/")
def inicio():
    cn = conectar()
    cur = cn.cursor(dictionary=True)

    cur.execute("SELECT COUNT(*) total FROM pacientes")
    pacientes = cur.fetchone()["total"]

    cur.execute("SELECT COUNT(*) total FROM medicamentos")
    medicamentos = cur.fetchone()["total"]

    cur.execute(
        "SELECT COUNT(*) total FROM medicamentos WHERE conferido=0"
    )
    pendentes = cur.fetchone()["total"]

    cur.execute(
        "SELECT id,nome FROM pacientes ORDER BY nome"
    )
    lista = cur.fetchall()

    cur.close()
    cn.close()

    return render_template(
        "index.html",
        pacientes=pacientes,
        medicamentos=medicamentos,
        pendentes=pendentes,
        lista=lista
    )


# ==================================================
# LISTAGEM DE PACIENTES
# ==================================================

@app.route("/pacientes")
def pacientes():
    cn = conectar()
    cur = cn.cursor(dictionary=True)

    cur.execute(
        """
        SELECT
            p.id,
            p.nome,
            COUNT(m.id) total
        FROM pacientes p
        LEFT JOIN medicamentos m
            ON m.paciente_id = p.id
        GROUP BY p.id, p.nome
        ORDER BY p.nome
        """
    )

    dados = cur.fetchall()

    cur.close()
    cn.close()

    return render_template(
        "pacientes.html",
        pacientes=dados
    )


# ==================================================
# CADASTRAR PACIENTE
# ==================================================

@app.route("/paciente/novo", methods=["GET", "POST"])
def novo_paciente():

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()

        if not nome:
            flash("Informe o nome do paciente.")
            return redirect(url_for("novo_paciente"))

        cpf = request.form.get("cpf", "").strip() or None
        rg = request.form.get("rg", "").strip() or None
        data_nascimento = (
            request.form.get("data_nascimento", "").strip() or None
        )
        telefone = request.form.get("telefone", "").strip() or None
        endereco = request.form.get("endereco", "").strip() or None
        observacoes = (
            request.form.get("observacoes", "").strip() or None
        )

        cn = conectar()
        cur = cn.cursor(dictionary=True)

        cur.execute(
            "SELECT id FROM pacientes WHERE nome=%s",
            (nome,)
        )

        existente = cur.fetchone()

        if existente:
            cur.close()
            cn.close()
            flash("Já existe um paciente cadastrado com esse nome.")
            return redirect(url_for("novo_paciente"))

        cur.execute(
            """
            INSERT INTO pacientes
            (
                nome, cpf, rg, data_nascimento,
                telefone, endereco, observacoes
            )
            VALUES (%s,%s,%s,%s,%s,%s,%s)
            """,
            (
                nome,
                cpf,
                rg,
                data_nascimento,
                telefone,
                endereco,
                observacoes
            )
        )

        paciente_id = cur.lastrowid

        cn.commit()
        cur.close()
        cn.close()

        flash("Paciente cadastrado com sucesso!")

        return redirect(
            url_for("paciente", paciente_id=paciente_id)
        )

    return render_template(
        "paciente_form.html",
        paciente=None,
        titulo="Novo paciente"
    )


# ==================================================
# EDITAR DADOS DO PACIENTE
# ==================================================

@app.route(
    "/paciente/<int:paciente_id>/editar",
    methods=["GET", "POST"]
)
def editar_paciente(paciente_id):

    cn = conectar()
    cur = cn.cursor(dictionary=True)

    cur.execute(
        "SELECT * FROM pacientes WHERE id=%s",
        (paciente_id,)
    )

    p = cur.fetchone()

    if not p:
        cur.close()
        cn.close()
        return "Paciente não encontrado.", 404

    if request.method == "POST":

        nome = request.form.get("nome", "").strip()

        if not nome:
            cur.close()
            cn.close()
            flash("O nome do paciente é obrigatório.")
            return redirect(
                url_for("editar_paciente", paciente_id=paciente_id)
            )

        cpf = request.form.get("cpf", "").strip() or None
        rg = request.form.get("rg", "").strip() or None
        data_nascimento = (
            request.form.get("data_nascimento", "").strip() or None
        )
        telefone = request.form.get("telefone", "").strip() or None
        endereco = request.form.get("endereco", "").strip() or None
        observacoes = (
            request.form.get("observacoes", "").strip() or None
        )

        cur.execute(
            """
            SELECT id FROM pacientes
            WHERE nome=%s AND id<>%s
            """,
            (nome, paciente_id)
        )

        duplicado = cur.fetchone()

        if duplicado:
            cur.close()
            cn.close()
            flash("Já existe outro paciente com esse nome.")
            return redirect(
                url_for("editar_paciente", paciente_id=paciente_id)
            )

        cur.execute(
            """
            UPDATE pacientes
            SET
                nome=%s,
                cpf=%s,
                rg=%s,
                data_nascimento=%s,
                telefone=%s,
                endereco=%s,
                observacoes=%s
            WHERE id=%s
            """,
            (
                nome,
                cpf,
                rg,
                data_nascimento,
                telefone,
                endereco,
                observacoes,
                paciente_id
            )
        )

        cn.commit()
        cur.close()
        cn.close()

        flash("Dados do paciente atualizados!")

        return redirect(
            url_for("paciente", paciente_id=paciente_id)
        )

    cur.close()
    cn.close()

    return render_template(
        "paciente_form.html",
        paciente=p,
        titulo="Editar paciente"
    )


# ==================================================
# FICHA INDIVIDUAL DO PACIENTE
# ==================================================

@app.route("/paciente/<int:paciente_id>")
def paciente(paciente_id):

    cn = conectar()
    cur = cn.cursor(dictionary=True)

    cur.execute(
        "SELECT * FROM pacientes WHERE id=%s",
        (paciente_id,)
    )

    p = cur.fetchone()

    if not p:
        cur.close()
        cn.close()
        return "Paciente não encontrado.", 404

    cur.execute(
        """
        SELECT *
        FROM medicamentos
        WHERE paciente_id=%s
        ORDER BY
            FIELD(
                horario,
                'MANHÃ',
                'ALMOÇO',
                'JANTAR',
                'NOITE'
            ),
            nome
        """,
        (paciente_id,)
    )

    meds = cur.fetchall()

    # Buscar receitas pelo vínculo direto com o paciente
    cur.execute(
        """
        SELECT
            id,
            arquivo,
            aba,
            importado_em
        FROM documentos_origem
        WHERE paciente_id = %s
        ORDER BY aba
        """,
        (paciente_id,)
    )

    receitas = cur.fetchall()

    cur.close()
    cn.close()

    return render_template(
        "paciente.html",
        paciente=p,
        medicamentos=meds,
        receitas=receitas
    )


# ==================================================
# ALTERAR STATUS DO MEDICAMENTO
# ==================================================

@app.route(
    "/medicamento/<int:medicamento_id>/status",
    methods=["POST"]
)
def alterar_status_medicamento(medicamento_id):

    novo_status = request.form.get("status", "").strip()
    motivo = request.form.get("motivo", "").strip() or None

    try:
        paciente_id = int(request.form.get("paciente_id", ""))
    except (TypeError, ValueError):
        flash("Identificação do paciente inválida.")
        return redirect(url_for("pacientes"))

    if novo_status not in STATUS_VALIDOS:
        flash("Status informado inválido.")
        return redirect(
            url_for("paciente", paciente_id=paciente_id)
        )

    cn = conectar()
    cur = cn.cursor(dictionary=True)

    try:
        cn.start_transaction()

        cur.execute(
            """
            SELECT id, paciente_id, status
            FROM medicamentos
            WHERE id=%s
            FOR UPDATE
            """,
            (medicamento_id,)
        )

        medicamento = cur.fetchone()

        if not medicamento:
            cn.rollback()
            flash("Medicamento não encontrado.")
            return redirect(
                url_for("paciente", paciente_id=paciente_id)
            )

        if medicamento["paciente_id"] != paciente_id:
            cn.rollback()
            flash("Medicamento não pertence a este paciente.")
            return redirect(url_for("pacientes"))

        status_anterior = medicamento["status"]

        if status_anterior == novo_status:
            cn.rollback()
            flash("O medicamento já possui esse status.")
            return redirect(
                url_for("paciente", paciente_id=paciente_id)
            )

        cur.execute(
            """
            INSERT INTO historico_status_medicamentos
            (
                medicamento_id,
                status_anterior,
                status_novo,
                motivo
            )
            VALUES (%s,%s,%s,%s)
            """,
            (
                medicamento_id,
                status_anterior,
                novo_status,
                motivo
            )
        )

        cur.execute(
            """
            UPDATE medicamentos
            SET
                status=%s,
                status_alterado_em=%s,
                motivo_status=%s
            WHERE id=%s
            """,
            (
                novo_status,
                datetime.now(),
                motivo,
                medicamento_id
            )
        )

        cn.commit()

        flash("Status do medicamento atualizado com sucesso!")

    except Exception as erro:
        cn.rollback()
        flash(
            "Não foi possível atualizar o status. "
            "Detalhe: " + str(erro)
        )

    finally:
        cur.close()
        cn.close()

    return redirect(
        url_for("paciente", paciente_id=paciente_id)
    )


# ==================================================
# ALTERAR CONFERÊNCIA DO MEDICAMENTO
# ==================================================

@app.route(
    "/medicamento/<int:medicamento_id>/conferencia",
    methods=["POST"]
)
def alterar_conferencia_medicamento(medicamento_id):

    try:
        paciente_id = int(request.form.get("paciente_id", ""))
        novo_estado = int(request.form.get("conferido", ""))

        if novo_estado not in (0, 1):
            raise ValueError

    except (TypeError, ValueError):
        flash("Dados da conferência inválidos.")
        return redirect(url_for("pacientes"))

    cn = conectar()
    cur = cn.cursor(dictionary=True)

    try:
        cur.execute(
            """
            SELECT id
            FROM medicamentos
            WHERE id=%s AND paciente_id=%s
            """,
            (medicamento_id, paciente_id)
        )

        medicamento = cur.fetchone()

        if not medicamento:
            flash("Medicamento não encontrado para este paciente.")
        else:
            cur.execute(
                """
                UPDATE medicamentos
                SET conferido=%s
                WHERE id=%s AND paciente_id=%s
                """,
                (novo_estado, medicamento_id, paciente_id)
            )

            cn.commit()

            if novo_estado == 1:
                flash("Medicamento marcado como conferido!")
            else:
                flash("Medicamento voltou para pendente.")

    except Exception as erro:
        cn.rollback()
        flash("Erro ao atualizar conferência: " + str(erro))

    finally:
        cur.close()
        cn.close()

    return redirect(
        url_for("paciente", paciente_id=paciente_id)
    )


# ==================================================
# HISTÓRICO DE ALTERAÇÕES DE MEDICAMENTOS
# ==================================================

@app.route("/medicamento/<int:medicamento_id>/historico")
def historico_medicamento(medicamento_id):

    cn = conectar()
    cur = cn.cursor(dictionary=True)

    cur.execute(
        """
        SELECT
            m.id,
            m.nome,
            m.paciente_id
        FROM medicamentos m
        WHERE m.id=%s
        """,
        (medicamento_id,)
    )

    medicamento = cur.fetchone()

    if not medicamento:
        cur.close()
        cn.close()
        return "Medicamento não encontrado.", 404

    cur.execute(
        """
        SELECT
            status_anterior,
            status_novo,
            motivo,
            alterado_em
        FROM historico_status_medicamentos
        WHERE medicamento_id=%s
        ORDER BY alterado_em DESC
        """,
        (medicamento_id,)
    )

    historico = cur.fetchall()

    cur.close()
    cn.close()

    return render_template(
        "historico_medicamento.html",
        medicamento=medicamento,
        historico=historico
    )


# ==================================================
# RELATÓRIOS DE RECEITAS
# ==================================================

@app.route("/relatorios/receitas")
def relatorios_receitas():

    cn = conectar()
    cur = cn.cursor(dictionary=True)

    cur.execute(
        """
        SELECT
            id,
            arquivo,
            aba,
            importado_em
        FROM documentos_origem
        ORDER BY arquivo, aba
        """
    )

    receitas = cur.fetchall()

    cur.close()
    cn.close()

    return render_template(
        "relatorios_receitas.html",
        receitas=receitas
    )


# ==================================================
# VISUALIZAR RECEITA INDIVIDUAL
# ==================================================

@app.route("/relatorios/receita/<int:receita_id>")
def visualizar_receita(receita_id):

    cn = conectar()
    cur = cn.cursor(dictionary=True)

    cur.execute(
        """
        SELECT
            id,
            arquivo,
            aba,
            conteudo_json,
            importado_em
        FROM documentos_origem
        WHERE id=%s
        """,
        (receita_id,)
    )

    receita = cur.fetchone()

    cur.close()
    cn.close()

    if not receita:
        return "Receita não encontrada.", 404

    try:
        linhas = json.loads(receita["conteudo_json"])
    except (TypeError, ValueError, json.JSONDecodeError):
        return "Não foi possível interpretar o conteúdo da receita.", 500

    return render_template(
        "visualizar_receita.html",
        receita=receita,
        linhas=linhas
    )


# ==================================================
# IMPORTAÇÃO DE ARQUIVOS EXCEL
# ==================================================

@app.route("/importar", methods=["GET", "POST"])
def importar():

    if request.method == "POST":

        arquivo = request.files.get("arquivo")
        tipo = request.form.get("tipo")

        if not arquivo or not arquivo.filename.lower().endswith(".xlsx"):
            flash("Selecione um arquivo Excel .xlsx.")
            return redirect(url_for("importar"))

        nome = secure_filename(arquivo.filename)
        caminho = os.path.join(UPLOAD_DIR, nome)

        arquivo.save(caminho)

        try:

            if tipo == "checagem":

                total = importar_checagem(caminho, nome)

                flash(
                    f"Importação concluída: {total} linhas de "
                    "medicamentos processadas. "
                    "Todos ficaram aguardando conferência."
                )

            elif tipo == "receitas":

                total = arquivar_receitas(caminho, nome)

                flash(
                    f"Arquivo de receitas arquivado: {total} "
                    "abas preservadas para conferência."
                )

            else:
                flash("Tipo de importação inválido.")

        except Exception as e:

            flash(
                "Não foi possível importar. Verifique o banco "
                "de dados e o formato do arquivo. Detalhe: "
                + str(e)
            )

        return redirect(url_for("importar"))

    return render_template("importar.html")


# ==================================================
# EXECUÇÃO
# ==================================================

if __name__ == "__main__":
    app.run(debug=True)
