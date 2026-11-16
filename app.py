from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from database import conectar
from config import Config
from datetime import datetime
import os, json

app = Flask(__name__)
app.config.from_object(Config)
UPLOAD_DIR = os.path.join(app.root_path, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
HORARIOS = {"MANHÃ", "ALMOÇO", "JANTAR", "NOITE", "S/N"}

def importar_checagem(caminho, nome_arquivo):
    wb = load_workbook(caminho, data_only=True, read_only=True)
    cn = conectar()
    cur = cn.cursor()
    inseridos = 0
    for aba in wb.worksheets:
        nome = str(aba.cell(1, 5).value or aba.title).strip()
        if not nome:
            nome = aba.title
        cur.execute("INSERT INTO pacientes (nome) VALUES (%s) ON DUPLICATE KEY UPDATE id=LAST_INSERT_ID(id)", (nome,))
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
            if normalizado in {"MEDICAÇÕES", "MEDICACOES", "QT", "OBS"}:
                continue
            # Guarda a descrição original; não presume dose nem frequência além do que está escrito.
            cur.execute("""
                INSERT INTO medicamentos
                (paciente_id, quantidade_texto, nome, horario, origem, conferido)
                VALUES (%s,%s,%s,%s,%s,0)
                ON DUPLICATE KEY UPDATE id=LAST_INSERT_ID(id)
            """, (paciente_id, str(qtd).strip() if qtd is not None else "",
                  texto, horario, nome_arquivo))
            inseridos += 1
    cur.execute("INSERT INTO importacoes (arquivo, tipo, total_registros) VALUES (%s,'checagem',%s)",
                (nome_arquivo, inseridos))
    cn.commit()
    cur.close(); cn.close()
    return inseridos

def arquivar_receitas(caminho, nome_arquivo):
    wb = load_workbook(caminho, data_only=True, read_only=True)
    cn = conectar(); cur = cn.cursor()
    total = 0
    for aba in wb.worksheets:
        dados = []
        for row in aba.iter_rows():
            valores = [c.value for c in row]
            if any(v is not None for v in valores):
                dados.append(valores)
        cur.execute("""
            INSERT INTO documentos_origem (arquivo, aba, conteudo_json)
            VALUES (%s,%s,%s)
        """, (nome_arquivo, aba.title, json.dumps(dados, ensure_ascii=False, default=str)))
        total += 1
    cur.execute("INSERT INTO importacoes (arquivo, tipo, total_registros) VALUES (%s,'receitas_arquivo',%s)",
                (nome_arquivo, total))
    cn.commit(); cur.close(); cn.close()
    return total

@app.route("/")
def inicio():
    cn = conectar(); cur = cn.cursor(dictionary=True)
    cur.execute("SELECT COUNT(*) total FROM pacientes"); pacientes = cur.fetchone()["total"]
    cur.execute("SELECT COUNT(*) total FROM medicamentos"); medicamentos = cur.fetchone()["total"]
    cur.execute("SELECT COUNT(*) total FROM medicamentos WHERE conferido=0"); pendentes = cur.fetchone()["total"]
    cur.execute("SELECT id,nome FROM pacientes ORDER BY nome")
    lista = cur.fetchall()
    cur.close(); cn.close()
    return render_template("index.html", pacientes=pacientes, medicamentos=medicamentos,
                           pendentes=pendentes, lista=lista)

@app.route("/pacientes")
def pacientes():
    cn=conectar(); cur=cn.cursor(dictionary=True)
    cur.execute("""SELECT p.id,p.nome,COUNT(m.id) total FROM pacientes p
                   LEFT JOIN medicamentos m ON m.paciente_id=p.id
                   GROUP BY p.id,p.nome ORDER BY p.nome""")
    dados=cur.fetchall(); cur.close(); cn.close()
    return render_template("pacientes.html", pacientes=dados)

@app.route("/paciente/<int:paciente_id>")
def paciente(paciente_id):
    cn=conectar(); cur=cn.cursor(dictionary=True)
    cur.execute("SELECT * FROM pacientes WHERE id=%s",(paciente_id,))
    p=cur.fetchone()
    cur.execute("SELECT * FROM medicamentos WHERE paciente_id=%s ORDER BY FIELD(horario,'MANHÃ','ALMOÇO','JANTAR','NOITE'),nome",(paciente_id,))
    meds=cur.fetchall(); cur.close(); cn.close()
    if not p: return "Paciente não encontrado",404
    return render_template("paciente.html", paciente=p, medicamentos=meds)

@app.route("/importar", methods=["GET","POST"])
def importar():
    if request.method=="POST":
        arquivo=request.files.get("arquivo")
        tipo=request.form.get("tipo")
        if not arquivo or not arquivo.filename.lower().endswith(".xlsx"):
            flash("Selecione um arquivo Excel .xlsx.")
            return redirect(url_for("importar"))
        nome=secure_filename(arquivo.filename)
        caminho=os.path.join(UPLOAD_DIR, nome)
        arquivo.save(caminho)
        try:
            if tipo=="checagem":
                total=importar_checagem(caminho,nome)
                flash(f"Importação concluída: {total} linhas de medicamentos processadas. Todos ficaram aguardando conferência.")
            elif tipo=="receitas":
                total=arquivar_receitas(caminho,nome)
                flash(f"Arquivo de receitas arquivado: {total} abas preservadas para conferência.")
            else:
                flash("Tipo de importação inválido.")
        except Exception as e:
            flash("Não foi possível importar. Verifique o banco de dados e o formato do arquivo. Detalhe: "+str(e))
        return redirect(url_for("importar"))
    return render_template("importar.html")

if __name__=="__main__":
    app.run(debug=True)
