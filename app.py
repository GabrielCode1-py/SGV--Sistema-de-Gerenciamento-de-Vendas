# ============================
# Sistema de Gerenciamento de Estoque Simplificado (SGES) - Versão Web com Flask
# ============================
from flask import Flask, render_template, request, redirect, url_for, flash
import os
import pandas as pd
from datetime import datetime
import unicodedata  # Para normalizar acentos

app = Flask(__name__)
app.secret_key = 'chave_secreta_para_flash'  # Necessário para usar flash

# ---------------------------------------------------------------------
# FUNÇÕES DO SISTEMA

# Função para normalizar strings (remover acentos e converter para minúsculas)
def normalizar_string(texto):
    # Remove acentos e converte para minúsculas
    return unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode('ascii').lower()

# Função para carregar produtos do Excel (se existir)
def carregar_produtos():
    if os.path.exists("estoque.xlsx"):
        df = pd.read_excel("estoque.xlsx")
        produtos = []
        for _, row in df.iterrows():
            produto = {
                "id": int(row["ID"]),
                "nome": row["Nome"],
                "quantidade": int(row["Quantidade"]),
                "preco_unitario": float(row["Preço Unitário"])
            }
            produtos.append(produto)
        return produtos
    return []

# Função para registrar logs com horário
def registrar_log(acao):
    with open("estoque_log.txt", "a") as arquivo_log:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        arquivo_log.write(f"[{timestamp}] {acao}\n")

# ------------------------------
# 1. Cadastro de produtos (via POST) com verificação de duplicatas
def cadastrar_produto(produtos, nome, quantidade, preco_unitario):
    nome_normalizado = normalizar_string(nome)
    # Verifica se o produto já existe (comparação case-insensitive e sem acentos)
    for produto in produtos:
        if normalizar_string(produto["nome"]) == nome_normalizado:
            return f"Erro: Produto '{nome}' já está cadastrado!"
    # Se não existe, cadastra
    id = len(produtos) + 1
    produto = {
        "id": id,
        "nome": nome,
        "quantidade": quantidade,
        "preco_unitario": preco_unitario
    }
    produtos.append(produto)
    registrar_log(f"Cadastrou produto: {nome} (ID: {id})")
    return f"Produto '{nome}' cadastrado com sucesso!"

# ------------------------------
# 2. Listagem dos produtos em estoque
def listar_produtos(produtos):
    registrar_log("Listou produtos em estoque")
    return produtos

# ------------------------------
# 3. Análise de desempenho com condicionais
def verificar_desempenho(produtos):
    desempenhos = []
    for produto in produtos:
        if produto["quantidade"] < 10:
            desempenho = "Baixo"
        elif 10 <= produto["quantidade"] <= 50:
            desempenho = "Médio"
        else:
            desempenho = "Alto"
        desempenhos.append({
            "nome": produto["nome"],
            "desempenho": desempenho
        })
    registrar_log("Verificou nível de desempenho dos produtos")
    return desempenhos

# ------------------------------
# 4. Criar matriz (corredores e prateleiras)
def gerar_matriz_estoque(produtos):
    corredores = 3
    prateleiras = 3
    matriz = []
    for corredor in range(corredores):
        corredor_data = {"numero": corredor + 1, "prateleiras": []}
        for prateleira in range(prateleiras):
            indice_produto = corredor * prateleiras + prateleira
            if indice_produto < len(produtos):
                produto = produtos[indice_produto]
                corredor_data["prateleiras"].append({
                    "numero": prateleira + 1,
                    "produto": produto["nome"],
                    "quantidade": produto["quantidade"]
                })
            else:
                corredor_data["prateleiras"].append({
                    "numero": prateleira + 1,
                    "produto": "Vazio",
                    "quantidade": None
                })
        matriz.append(corredor_data)
    registrar_log("Gerou matriz de estoque")
    return matriz

# ------------------------------
# 5. Gerar nuvem de TAGs únicas (SET)
def gerar_nuvem_tags(produtos):
    tags = set()
    for produto in produtos:
        tags.add(produto["nome"])
    registrar_log("Gerou nuvem de TAGs únicas")
    return list(tags)

# ------------------------------
# 6. Salvar dados no Excel usando pandas (com tratamento de erro)
def salvar_dados(produtos):
    try:
        df = pd.DataFrame(produtos)
        df.columns = ["ID", "Nome", "Quantidade", "Preço Unitário"]
        df.to_excel("estoque.xlsx", index=False)
        registrar_log("Salvou dados em estoque.xlsx")
        return "Dados salvos com sucesso!"
    except Exception as e:
        registrar_log(f"Erro ao salvar dados: {str(e)}")
        return f"Erro ao salvar dados: {str(e)}. Verifique se o arquivo Excel não está aberto."

# ---------------------------------------------------------------------
# ROTAS DO FLASK

produtos = carregar_produtos()  # Carrega produtos no início
registrar_log("Iniciou o sistema SGES")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/cadastrar', methods=['GET', 'POST'])
def cadastrar():
    if request.method == 'POST':
        nome = request.form.get('nome')
        quantidade = int(request.form.get('quantidade'))
        preco_unitario = float(request.form.get('preco_unitario'))
        mensagem = cadastrar_produto(produtos, nome, quantidade, preco_unitario)
        flash(mensagem)  # Exibe mensagem temporária (sucesso ou erro)
        return redirect(url_for('index'))  # Redireciona para o menu
    return render_template('cadastrar.html')

@app.route('/listar')
def listar():
    produtos_list = listar_produtos(produtos)
    return render_template('listar.html', produtos=produtos_list)

@app.route('/desempenho')
def desempenho():
    desempenhos = verificar_desempenho(produtos)
    return render_template('desempenho.html', desempenhos=desempenhos)

@app.route('/matriz')
def matriz():
    matriz_data = gerar_matriz_estoque(produtos)
    return render_template('matriz.html', matriz=matriz_data)

@app.route('/tags')
def tags():
    tags_list = gerar_nuvem_tags(produtos)
    return render_template('tags.html', tags=tags_list)

@app.route('/salvar')
def salvar():
    mensagem = salvar_dados(produtos)
    flash(mensagem)  # Exibe mensagem temporária (sucesso ou erro)
    return redirect(url_for('index'))  # Redireciona para o menu (não encerra)

if __name__ == '__main__':
    app.run(debug=True)