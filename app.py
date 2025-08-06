from flask import Flask, render_template, request, redirect, url_for, session, send_file
import sqlite3
import pandas as pd
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'chave_secreta_segura'
DB = 'acesso.db'

def criar_banco():
    with sqlite3.connect(DB) as con:
        cur = con.cursor()
        cur.execute("""CREATE TABLE IF NOT EXISTS usuarios (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT,
                        senha TEXT,
                        tipo TEXT
                    )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS acessos (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nome TEXT,
                        documento TEXT,
                        secoes TEXT,
                        horario_entrada TEXT,
                        horario_saida TEXT
                    )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS atendimentos (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        acesso_id INTEGER,
                        secao TEXT,
                        horario_inicio TEXT,
                        horario_fim TEXT
                    )""")
        cur.execute("SELECT COUNT(*) FROM usuarios")
        if cur.fetchone()[0] == 0:
            usuarios = [('rp1', 'senha1', 'RP'), ('secao', 'senha2', 'SECAO')]
            cur.executemany("INSERT INTO usuarios (username, senha, tipo) VALUES (?, ?, ?)", usuarios)
        con.commit()

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        senha = request.form['senha']
        with sqlite3.connect(DB) as con:
            cur = con.cursor()
            cur.execute("SELECT tipo FROM usuarios WHERE username=? AND senha=?", (username, senha))
            user = cur.fetchone()
        if user:
            session['username'] = username
            session['tipo'] = user[0]
            return redirect('/rp' if user[0] == 'RP' else '/secao')
        else:
            return render_template('login.html', erro='Usuário ou senha inválidos.')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

@app.route('/rp', methods=['GET', 'POST'])
def rp():
    if session.get('tipo') != 'RP':
        return redirect('/')
    if request.method == 'POST':
        nome = request.form['nome']
        documento = request.form['documento']
        secoes = request.form['secoes']
        horario_entrada = request.form['entrada']
        with sqlite3.connect(DB) as con:
            cur = con.cursor()
            cur.execute("INSERT INTO acessos (nome, documento, secoes, horario_entrada) VALUES (?, ?, ?, ?)",
                        (nome, documento, secoes, horario_entrada))
            con.commit()
    with sqlite3.connect(DB) as con:
        cur = con.cursor()
        cur.execute("SELECT * FROM acessos WHERE horario_saida IS NULL")
        visitantes = cur.fetchall()
    return render_template('rp.html', visitantes=visitantes)

@app.route('/registrar_saida/<int:id>', methods=['POST'])
def registrar_saida(id):
    if session.get('tipo') != 'RP':
        return redirect('/')
    horario_saida = request.form['saida']
    with sqlite3.connect(DB) as con:
        cur = con.cursor()
        cur.execute("UPDATE acessos SET horario_saida=? WHERE id=?", (horario_saida, id))
        con.commit()
    return redirect('/rp')

@app.route('/secao', methods=['GET', 'POST'])
def secao():
    if session.get('tipo') != 'SECAO':
        return redirect('/')
    if request.method == 'POST':
        acesso_id = request.form['id']
        secao = request.form['secao']
        inicio = request.form['inicio']
        fim = request.form['fim']
        with sqlite3.connect(DB) as con:
            cur = con.cursor()
            cur.execute("INSERT INTO atendimentos (acesso_id, secao, horario_inicio, horario_fim) VALUES (?, ?, ?, ?)",
                        (acesso_id, secao, inicio, fim))
            con.commit()
    with sqlite3.connect(DB) as con:
        cur = con.cursor()
        cur.execute("SELECT * FROM acessos WHERE horario_saida IS NULL")
        visitantes = cur.fetchall()
    return render_template('secao.html', visitantes=visitantes)

@app.route('/relatorio')
def relatorio():
    if session.get('tipo') != 'RP':
        return redirect('/')
    with sqlite3.connect(DB) as con:
        df_visitantes = pd.read_sql_query("SELECT * FROM acessos WHERE horario_saida IS NOT NULL", con)
        df_atendimentos = pd.read_sql_query("SELECT * FROM atendimentos", con)

    relatorio_completo = []
    for _, visitante in df_visitantes.iterrows():
        atendimentos = df_atendimentos[df_atendimentos['acesso_id'] == visitante['id']]
        atendimentos_formatados = '; '.join([f"{row['secao']}: {row['horario_inicio']} - {row['horario_fim']}" for _, row in atendimentos.iterrows()])
        relatorio_completo.append({
            'Nome': visitante['nome'],
            'Documento': visitante['documento'],
            'Seções pretendidas': visitante['secoes'],
            'Entrada': visitante['horario_entrada'],
            'Atendimentos': atendimentos_formatados,
            'Saída': visitante['horario_saida']
        })

    df_final = pd.DataFrame(relatorio_completo)
    nome_arquivo = f'relatorio_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
    df_final.to_excel(nome_arquivo, index=False)
    return send_file(nome_arquivo, as_attachment=True)

@app.route('/resetar')
def resetar():
    if session.get('tipo') != 'RP':
        return redirect('/')
    with sqlite3.connect(DB) as con:
        cur = con.cursor()
        cur.execute('DELETE FROM acessos')
        cur.execute('DELETE FROM atendimentos')
        con.commit()
    return redirect('/rp')

if __name__ == '__main__':
    criar_banco()
    app.run(debug=True)
