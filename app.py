from flask import Flask, render_template, request, redirect, url_for, flash, session
import os
import pandas as pd
from datetime import datetime, date
import unicodedata
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)
app.secret_key = 'chave_secreta_para_flash_e_session'

# ---------------------------------------------------------------------
# FUNÇÕES AUXILIARES

def normalizar_string(texto):
    """Normaliza strings removendo acentos e convertendo para minúsculas."""
    if not isinstance(texto, str):
        return ""
    return unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode('ascii').lower()

def registrar_log(acao):
    """Registra ações no log com timestamp (append)."""
    try:
        with open("vendas_log.txt", "a", encoding="utf-8") as arquivo_log:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            arquivo_log.write(f"[{timestamp}] {acao}\n")
    except Exception as e:
        print(f"Erro ao registrar log: {e}")

def validar_numero_positivo(valor):
    """Valida se um valor numérico é positivo."""
    try:
        valor_float = float(valor)
        if valor_float <= 0:
            return None
        return valor_float
    except (ValueError, TypeError):
        return None

# ---------------------------------------------------------------------
# MÓDULO: PRODUTOS (estoque.xlsx)

def carregar_produtos():
    """Carrega produtos do arquivo estoque.xlsx."""
    if not os.path.exists("estoque.xlsx"):
        registrar_log("Arquivo estoque.xlsx não encontrado. Iniciando com lista vazia.")
        return []
    try:
        df = pd.read_excel("estoque.xlsx")
        if df.empty:
            return []
        # Validação de colunas (Quantidade agora é opcional)
        colunas_esperadas = ["ID", "Nome", "Tipo", "Valor", "Controlar_Estoque"]
        if not all(col in df.columns for col in colunas_esperadas):
            registrar_log(f"Erro: Colunas inválidas em estoque.xlsx. Esperado: {colunas_esperadas}")
            return []
        produtos = []
        for _, row in df.iterrows():
            produto = {
                "id": int(row["ID"]),
                "nome": str(row["Nome"]),
                "tipo": str(row["Tipo"]),  # "unitario", "quilo", "esteira"
                "valor": float(row["Valor"]),
                "controlar_estoque": bool(row["Controlar_Estoque"]),  # True/False
                "quantidade": float(row.get("Quantidade", 0)) if row.get("Controlar_Estoque") else 0
            }
            produtos.append(produto)
        return produtos
    except Exception as e:
        registrar_log(f"Erro ao carregar produtos: {str(e)}")
        return []

def salvar_produtos(produtos):
    """Salva produtos no arquivo estoque.xlsx."""
    try:
        df = pd.DataFrame(produtos)
        df = df[["id", "nome", "tipo", "valor", "controlar_estoque", "quantidade"]].rename(
            columns={
                "id": "ID", 
                "nome": "Nome", 
                "tipo": "Tipo", 
                "valor": "Valor", 
                "controlar_estoque": "Controlar_Estoque",
                "quantidade": "Quantidade"
            }
        )
        
        if os.path.exists("estoque.xlsx"):
            wb = load_workbook("estoque.xlsx")
            if "Produtos" in wb.sheetnames:
                del wb["Produtos"]
            ws = wb.create_sheet("Produtos")
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Produtos"
        
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        wb.save("estoque.xlsx")
        registrar_log("Produtos salvos em estoque.xlsx")
    except Exception as e:
        registrar_log(f"Erro ao salvar produtos: {str(e)}")
        raise

def cadastrar_produto(produtos, nome, tipo, valor, controlar_estoque, quantidade=0):
    """Cadastra novo produto com opção de estoque."""
    # Validação de nome vazio
    if not nome or nome.strip() == "":
        return "Erro: Nome do produto não pode estar vazio!"
    
    # Validação de tipo
    tipos_validos = ["unitario", "quilo", "esteira"]
    if tipo not in tipos_validos:
        return f"Erro: Tipo inválido! Use: {', '.join(tipos_validos)}"
    
    # Validação de valor
    valor_validado = validar_numero_positivo(valor)
    if not valor_validado:
        return "Erro: Valor deve ser maior que zero!"
    
    # Validação de quantidade (só se controlar estoque)
    if controlar_estoque:
        quantidade_validada = validar_numero_positivo(quantidade)
        if not quantidade_validada:
            return "Erro: Quantidade inválida para produto com estoque!"
    else:
        quantidade_validada = 0  # Estoque ilimitado (produtos fabricados na hora)
    
    # Verificação de duplicatas
    nome_normalizado = normalizar_string(nome)
    for produto in produtos:
        if normalizar_string(produto["nome"]) == nome_normalizado:
            return f"Erro: Produto '{nome}' já está cadastrado!"
    
    # Cadastra produto
    id = len(produtos) + 1
    produto = {
        "id": id,
        "nome": nome.strip(),
        "tipo": tipo,
        "valor": valor_validado,
        "controlar_estoque": controlar_estoque,
        "quantidade": quantidade_validada
    }
    produtos.append(produto)
    tipo_estoque = "com estoque" if controlar_estoque else "fabricado na hora"
    registrar_log(f"Cadastrou produto: {nome} (ID: {id}, Tipo: {tipo}, {tipo_estoque})")
    return f"Produto '{nome}' cadastrado com sucesso!"

def remover_produto(produtos, produto_id):
    """Remove produto do estoque."""
    produto = next((p for p in produtos if p["id"] == produto_id), None)
    if not produto:
        return "Erro: Produto não encontrado!"
    produtos.remove(produto)
    salvar_produtos(produtos)
    registrar_log(f"Removeu produto: {produto['nome']} (ID: {produto_id})")
    return f"Produto '{produto['nome']}' removido com sucesso!"

def atualizar_estoque(produtos, produto_id, quantidade_vendida):
    """Atualiza quantidade em estoque após venda (só se controlar estoque)."""
    produto = next((p for p in produtos if p["id"] == produto_id), None)
    if produto and produto["controlar_estoque"]:
        produto["quantidade"] -= quantidade_vendida
        if produto["quantidade"] < 0:
            produto["quantidade"] = 0
        salvar_produtos(produtos)

# ---------------------------------------------------------------------
# MÓDULO: CLIENTES (vendas.xlsx, aba Clientes)

def carregar_clientes():
    """Carrega clientes do arquivo vendas.xlsx."""
    if not os.path.exists("vendas.xlsx"):
        return []
    try:
        df = pd.read_excel("vendas.xlsx", sheet_name="Clientes")
        if df.empty:
            return []
        clientes = []
        for _, row in df.iterrows():
            cliente = {
                "id": int(row["ID"]),
                "nome": str(row["Nome"]),
                "telefone": str(row["Telefone"]) if pd.notna(row.get("Telefone")) else "",
                "observacoes": str(row["Observacoes"]) if pd.notna(row.get("Observacoes")) else ""
            }
            clientes.append(cliente)
        return clientes
    except Exception as e:
        registrar_log(f"Erro ao carregar clientes: {str(e)}")
        return []

def salvar_clientes(clientes):
    """Salva clientes no arquivo vendas.xlsx."""
    try:
        df = pd.DataFrame(clientes)
        df = df[["id", "nome", "telefone", "observacoes"]].rename(
            columns={"id": "ID", "nome": "Nome", "telefone": "Telefone", "observacoes": "Observacoes"}
        )
        
        if os.path.exists("vendas.xlsx"):
            wb = load_workbook("vendas.xlsx")
            if "Clientes" in wb.sheetnames:
                del wb["Clientes"]
            ws = wb.create_sheet("Clientes")
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Clientes"
        
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        wb.save("vendas.xlsx")
        registrar_log("Clientes salvos em vendas.xlsx")
    except Exception as e:
        registrar_log(f"Erro ao salvar clientes: {str(e)}")
        raise

def cadastrar_cliente(clientes, nome, telefone, observacoes):
    """Cadastra novo cliente."""
    if not nome or nome.strip() == "":
        return "Erro: Nome do cliente não pode estar vazio!"
    
    nome_normalizado = normalizar_string(nome)
    for cliente in clientes:
        if normalizar_string(cliente["nome"]) == nome_normalizado:
            return f"Erro: Cliente '{nome}' já está cadastrado!"
    
    id = len(clientes) + 1
    cliente = {
        "id": id,
        "nome": nome.strip(),
        "telefone": telefone.strip(),
        "observacoes": observacoes.strip()
    }
    clientes.append(cliente)
    salvar_clientes(clientes)
    registrar_log(f"Cadastrou cliente: {nome} (ID: {id})")
    return f"Cliente '{nome}' cadastrado com sucesso!"

def remover_cliente(clientes, cliente_id):
    """Remove cliente."""
    cliente = next((c for c in clientes if c["id"] == cliente_id), None)
    if not cliente:
        return "Erro: Cliente não encontrado!"
    clientes.remove(cliente)
    salvar_clientes(clientes)
    registrar_log(f"Removeu cliente: {cliente['nome']} (ID: {cliente_id})")
    return f"Cliente '{cliente['nome']}' removido com sucesso!"

# ---------------------------------------------------------------------
# MÓDULO: CARRINHO DE VENDAS (Session-based)

def inicializar_carrinho():
    """Inicializa carrinho na session."""
    if 'carrinho' not in session:
        session['carrinho'] = []
    if 'cliente_id_carrinho' not in session:
        session['cliente_id_carrinho'] = None

def adicionar_item_carrinho(produtos, produto_id, quantidade_input):
    """Adiciona item ao carrinho com cálculos."""
    produto = next((p for p in produtos if p["id"] == produto_id), None)
    if not produto:
        return None, "Erro: Produto não encontrado!"
    
    quantidade_validada = validar_numero_positivo(quantidade_input)
    if not quantidade_validada:
        return None, "Erro: Quantidade inválida!"
    
    # Lógica de cálculo por tipo
    if produto["tipo"] == "esteira":
        quantidade_total = quantidade_validada * 30
        valor_total = round(quantidade_total * produto["valor"], 2)
        descricao = f"{int(quantidade_validada)} esteiras ({int(quantidade_total)} pães)"
    elif produto["tipo"] == "quilo":
        quantidade_total = quantidade_validada
        valor_total = round(quantidade_validada * produto["valor"], 2)
        descricao = f"{quantidade_validada} kg"
    else:
        quantidade_total = quantidade_validada
        valor_total = round(quantidade_validada * produto["valor"], 2)
        descricao = f"{int(quantidade_validada)} unidades"
    
    # Validação de estoque (só se controlar estoque)
    if produto["controlar_estoque"]:
        if produto["quantidade"] < quantidade_total:
            return None, f"Erro: Estoque insuficiente! Disponível: {produto['quantidade']}"
    
    item = {
        "produto_id": produto_id,
        "produto_nome": produto["nome"],
        "tipo_produto": produto["tipo"],
        "quantidade_input": quantidade_validada,
        "quantidade_total": quantidade_total,
        "valor_unitario": produto["valor"],
        "valor_total": valor_total,
        "descricao": descricao
    }
    return item, None

def finalizar_pedido(clientes, produtos, cliente_id, carrinho):
    """Finaliza pedido, salva vendas e atualiza estoque (se necessário)."""
    cliente = next((c for c in clientes if c["id"] == cliente_id), None)
    if not cliente:
        return "Erro: Cliente não encontrado!"
    
    if not carrinho:
        return "Erro: Carrinho vazio! Adicione itens antes de finalizar."
    
    total_pedido = sum(item["valor_total"] for item in carrinho)
    data_venda = date.today().isoformat()
    
    # Salva cada item do carrinho como venda separada
    for item in carrinho:
        venda = {
            "Cliente_ID": cliente_id,
            "Cliente_Nome": cliente["nome"],
            "Produto_Nome": item["produto_nome"],
            "Tipo_Produto": item["tipo_produto"],
            "Quantidade_Input": item["quantidade_input"],
            "Quantidade_Total": item["quantidade_total"],
            "Valor_Unitario": item["valor_unitario"],
            "Valor_Total": item["valor_total"],
            "Data": data_venda
        }
        salvar_venda_diaria(venda)
        
        # Atualiza estoque (só se produto controlar estoque)
        atualizar_estoque(produtos, item["produto_id"], item["quantidade_total"])
    
    registrar_log(f"Pedido finalizado para {cliente['nome']} - Total: R$ {total_pedido:.2f}")
    return f"Pedido finalizado para {cliente['nome']} - Total: R$ {total_pedido:.2f}"

# ---------------------------------------------------------------------
# MÓDULO: VENDAS (vendas.xlsx, aba Diario)

def carregar_vendas_diarias():
    """Carrega vendas diárias do arquivo vendas.xlsx."""
    if not os.path.exists("vendas.xlsx"):
        return []
    try:
        df = pd.read_excel("vendas.xlsx", sheet_name="Diario")
        if df.empty:
            return []
        vendas = df.to_dict('records')
        return vendas
    except Exception as e:
        registrar_log(f"Erro ao carregar vendas diárias: {str(e)}")
        return []

def salvar_venda_diaria(venda):
    """Salva venda diária no arquivo vendas.xlsx usando openpyxl."""
    try:
        df_nova = pd.DataFrame([venda])
        
        if os.path.exists("vendas.xlsx"):
            wb = load_workbook("vendas.xlsx")
            if "Diario" in wb.sheetnames:
                df_existente = pd.read_excel("vendas.xlsx", sheet_name="Diario")
                df_atualizado = pd.concat([df_existente, df_nova], ignore_index=True)
                del wb["Diario"]
            else:
                df_atualizado = df_nova
            ws = wb.create_sheet("Diario")
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Diario"
            df_atualizado = df_nova
        
        for r in dataframe_to_rows(df_atualizado, index=False, header=True):
            ws.append(r)
        
        wb.save("vendas.xlsx")
        registrar_log("Venda diária salva")
    except Exception as e:
        registrar_log(f"Erro ao salvar venda diária: {str(e)}")
        raise

# ---------------------------------------------------------------------
# MÓDULO: FECHAMENTOS

def fechamento_mensal():
    """Cria aba mensal consolidada no vendas.xlsx."""
    vendas = carregar_vendas_diarias()
    if not vendas:
        return "Nenhum dado para fechamento mensal."
    
    mes_atual = date.today().strftime("%Y_%m")
    nome_aba = f"Mes_{mes_atual}"
    
    # Verifica se aba já existe
    try:
        if os.path.exists("vendas.xlsx"):
            wb = load_workbook("vendas.xlsx")
            if nome_aba in wb.sheetnames:
                return f"Aba '{nome_aba}' já existe. Fechamento mensal já foi realizado."
            wb.close()
    except Exception as e:
        registrar_log(f"Erro ao verificar abas: {str(e)}")
    
    # Consolida vendas por cliente
    df_vendas = pd.DataFrame(vendas)
    df_consolidado = df_vendas.groupby("Cliente_Nome").agg({
        "Valor_Total": "sum"
    }).reset_index()
    df_consolidado.columns = ["Cliente", "Valor_Total"]
    
    try:
        wb = load_workbook("vendas.xlsx")
        ws = wb.create_sheet(nome_aba)
        for r in dataframe_to_rows(df_consolidado, index=False, header=True):
            ws.append(r)
        wb.save("vendas.xlsx")
        registrar_log(f"Fechamento mensal criado: {nome_aba}")
        return f"Fechamento mensal criado: {nome_aba}"
    except Exception as e:
        registrar_log(f"Erro no fechamento mensal: {str(e)}")
        return f"Erro no fechamento mensal: {str(e)}"

# ---------------------------------------------------------------------
# ROTAS DO FLASK

produtos = carregar_produtos()
clientes = carregar_clientes()
registrar_log("Sistema iniciado")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/cadastrar_produto', methods=['GET', 'POST'])
def cadastrar_produto_route():
    if request.method == 'POST':
        nome = request.form.get('nome')
        tipo = request.form.get('tipo')
        valor = request.form.get('valor')
        controlar_estoque = request.form.get('controlar_estoque') == 'sim'
        quantidade = request.form.get('quantidade', 0) if controlar_estoque else 0
        mensagem = cadastrar_produto(produtos, nome, tipo, valor, controlar_estoque, quantidade)
        flash(mensagem)
        return redirect(url_for('index'))
    return render_template('cadastrar_produto.html')

@app.route('/remover_produto/<int:produto_id>', methods=['POST'])
def remover_produto_route(produto_id):
    mensagem = remover_produto(produtos, produto_id)
    flash(mensagem)
    return redirect(url_for('listar_produtos_route'))

@app.route('/produtos')
def listar_produtos_route():
    return render_template('produtos.html', produtos=produtos)

@app.route('/cadastrar_cliente', methods=['GET', 'POST'])
def cadastrar_cliente_route():
    if request.method == 'POST':
        nome = request.form.get('nome')
        telefone = request.form.get('telefone', '')
        observacoes = request.form.get('observacoes', '')
        mensagem = cadastrar_cliente(clientes, nome, telefone, observacoes)
        flash(mensagem)
        return redirect(url_for('index'))
    return render_template('cadastrar_cliente.html')

@app.route('/remover_cliente/<int:cliente_id>', methods=['POST'])
def remover_cliente_route(cliente_id):
    mensagem = remover_cliente(clientes, cliente_id)
    flash(mensagem)
    return redirect(url_for('listar_clientes_route'))

@app.route('/clientes')
def listar_clientes_route():
    return render_template('clientes.html', clientes=clientes)

@app.route('/cliente/<int:cliente_id>')
def cliente_detalhes(cliente_id):
    cliente = next((c for c in clientes if c["id"] == cliente_id), None)
    if not cliente:
        flash("Cliente não encontrado!")
        return redirect(url_for('listar_clientes_route'))
    
    # Inicializa carrinho
    inicializar_carrinho()
    session['cliente_id_carrinho'] = cliente_id
    
    # Busca vendas históricas do cliente
    vendas_cliente = []
    vendas = carregar_vendas_diarias()
    for venda in vendas:
        if venda["Cliente_ID"] == cliente_id:
            vendas_cliente.append(venda)
    
    total_carrinho = sum(item["valor_total"] for item in session.get('carrinho', []))
    
    return render_template('cliente_detalhes.html', 
                         cliente=cliente, 
                         produtos=produtos, 
                         carrinho=session.get('carrinho', []),
                         total_carrinho=total_carrinho,
                         vendas_cliente=vendas_cliente)

@app.route('/adicionar_carrinho', methods=['POST'])
def adicionar_carrinho():
    """Adiciona item ao carrinho."""
    inicializar_carrinho()
    cliente_id = session.get('cliente_id_carrinho')
    if not cliente_id:
        flash("Erro: Selecione um cliente primeiro!")
        return redirect(url_for('listar_clientes_route'))
    
    produto_id = int(request.form.get('produto_id'))
    quantidade = request.form.get('quantidade')
    
    item, erro = adicionar_item_carrinho(produtos, produto_id, quantidade)
    if erro:
        flash(erro)
    else:
        session['carrinho'].append(item)
        session.modified = True
        flash(f"Item adicionado: {item['descricao']} - R$ {item['valor_total']:.2f}")
    
    return redirect(url_for('cliente_detalhes', cliente_id=cliente_id))

@app.route('/finalizar_pedido', methods=['POST'])
def finalizar_pedido_route():
    """Finaliza pedido e limpa carrinho."""
    cliente_id = session.get('cliente_id_carrinho')
    carrinho = session.get('carrinho', [])
    
    if not cliente_id:
        flash("Erro: Nenhum cliente selecionado!")
        return redirect(url_for('listar_clientes_route'))
    
    mensagem = finalizar_pedido(clientes, produtos, cliente_id, carrinho)
    flash(mensagem)
    
    # Limpa carrinho
    session['carrinho'] = []
    session['cliente_id_carrinho'] = None
    session.modified = True
    
    return redirect(url_for('index'))

@app.route('/limpar_carrinho', methods=['POST'])
def limpar_carrinho():
    """Limpa carrinho sem finalizar."""
    cliente_id = session.get('cliente_id_carrinho')
    session['carrinho'] = []
    session.modified = True
    flash("Carrinho limpo!")
    if cliente_id:
        return redirect(url_for('cliente_detalhes', cliente_id=cliente_id))
    return redirect(url_for('index'))

@app.route('/relatorios')
def relatorios():
    vendas_diarias = carregar_vendas_diarias()
    total_geral = sum(v["Valor_Total"] for v in vendas_diarias)
    
    # Agrupa por cliente
    vendas_por_cliente = {}
    for venda in vendas_diarias:
        cliente = venda["Cliente_Nome"]
        if cliente not in vendas_por_cliente:
            vendas_por_cliente[cliente] = {"vendas": [], "total": 0}
        vendas_por_cliente[cliente]["vendas"].append(venda)
        vendas_por_cliente[cliente]["total"] += venda["Valor_Total"]
    
    return render_template('relatorios.html', 
                         vendas_diarias=vendas_diarias, 
                         total_geral=total_geral,
                         vendas_por_cliente=vendas_por_cliente)

@app.route('/fechamento_mensal')
def fechamento_mensal_route():
    mensagem = fechamento_mensal()
    flash(mensagem)
    return redirect(url_for('index'))

@app.route('/listar')
def listar():
    return render_template('listar.html', produtos=produtos)

@app.route('/salvar')
def salvar():
    try:
        salvar_produtos(produtos)
        salvar_clientes(clientes)
        registrar_log("Dados salvos")
        return render_template('salvar.html')
    except Exception as e:
        flash(f"Erro ao salvar: {str(e)}")
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=False, host='127.0.0.1', port=5000)