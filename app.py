import streamlit as st
import json
import os
from datetime import datetime
from io import BytesIO
from docx import Document
import unicodedata
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sys

# Adicione ESTAS DUAS LINHAS para configurar a porta
# Isso cria uma variável de ambiente que o Streamlit irá ler
os.environ["STREAMLIT_SERVER_PORT"] = "8502" # Tentar a porta 8502 (ou outra, se preferir)

# --- Configurações Iniciais ---
# Verifica se o aplicativo está rodando como um executável PyInstaller
if getattr(sys, 'frozen', False):
    # Se sim, o caminho base é o diretório temporário onde o PyInstaller extrai os arquivos
    base_path = sys._MEIPASS
else:
    # Se não, está rodando em ambiente Python normal, o caminho base é o diretório atual do script
    base_path = os.path.abspath(".")

# Ajusta DATA_DIR para usar o base_path
DATA_DIR = os.path.join(base_path, "data")
USERS_FILE = os.path.join(DATA_DIR, "users.json")
PARECERES_FILE = os.path.join(DATA_DIR, "pareceres.json")
SCHOOL_NAME = "ESCOLA MUNICIPAL DE EDUCAÇÃO FUNDAMENTAL ELESBÃO BARBOSA DE CARVALHO"
COORDENADOR_NAME = "NOME DO COORDENADOR AQUI"

# --- Funções de Utilitário ---
def load_data(file_path, default_value):
    """Carrega dados de um arquivo JSON, retornando um valor padrão se não existir."""
    if not os.path.exists(file_path):
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(default_value, f, ensure_ascii=False, indent=4)
    with open(file_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_data(data, file_path):
    """Salva dados em um arquivo JSON."""
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def sanitize_student_name_for_filename(name):
    """Remove caracteres especiais e espaços do nome para usar em nomes de arquivo."""
    sanitized = ''.join(c for c in name if c.isalnum() or c.isspace())
    return sanitized.strip().replace(' ', '_')

def create_parecer_docx(data):
    """Cria um documento Word com o parecer do aluno a partir de um template."""
    template_path = os.path.join(DATA_DIR, "parecer_template.docx")
    try:
        doc = Document(template_path)
    except FileNotFoundError:
        return st.error(f"Erro: Arquivo de template não encontrado em {template_path}. Certifique-se de que o arquivo 'parecer_template.docx' está na pasta 'data'.")

    replacements = {
        'Nº ': data['numero_aluno'],
        'Período: ': data['periodo'],
        'Turma: ': data['turma'],
        'Turno: ': data['turno'],
        'Ano Letivo: ': data['ano_letivo'],
        'Semestre: ': data['semestre'],
        'Nome do Aluno (a):': data['nome_aluno'],
        'Filiação: ': f"Filiação: {data['filiacao_mae']} e {data['filiacao_pai']}",
        'Endereço: ': data['endereco'],
        'UF: ': data['uf'],
        'Data de Nascimento: ': data['data_nascimento'],
        'Naturalidade: ': data['naturalidade'],
        'Professor': data['nome_professor'],
        'Coordenador': COORDENADOR_NAME,
        'Maravilha, AL': f"Maravilha, AL, {data['data_parecer']}",
    }

    # Substituir no documento inteiro
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, value)

    # Inserir o corpo do parecer
    for paragraph in doc.paragraphs:
        if 'PARECER DESCRITIVO:' in paragraph.text:
            p = paragraph.insert_paragraph_before(data['texto_parecer'])
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            break

    # Salvar em BytesIO para download
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- Funções de Autenticação ---
def login():
    st.session_state['logged_in'] = False
    with st.form("login_form"):
        username = st.text_input("Usuário")
        password = st.text_input("Senha", type="password")
        col1, col2 = st.columns(2)
        with col1:
            if st.form_submit_button("Entrar"):
                users = load_data(USERS_FILE, {})
                if username in users and users[username]['password'] == password:
                    st.session_state['logged_in'] = True
                    st.session_state['username'] = username
                    st.success("Login bem-sucedido!")
                    st.experimental_rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
        with col2:
            st.write("---")
            if st.form_submit_button("Criar Conta"):
                st.session_state['show_create_account'] = True

def create_account():
    st.session_state['logged_in'] = False
    st.session_state['show_create_account'] = True
    with st.form("create_account_form"):
        st.write("### Criar Nova Conta")
        new_username = st.text_input("Novo Usuário", key="new_user")
        new_password = st.text_input("Nova Senha", type="password", key="new_pass")
        confirm_password = st.text_input("Confirmar Senha", type="password", key="confirm_pass")
        if st.form_submit_button("Confirmar Cadastro"):
            if not new_username or not new_password or not confirm_password:
                st.error("Preencha todos os campos.")
            elif new_password != confirm_password:
                st.error("As senhas não coincidem.")
            else:
                users = load_data(USERS_FILE, {})
                if new_username in users:
                    st.error("Usuário já existe. Escolha outro nome.")
                else:
                    users[new_username] = {'password': new_password, 'is_admin': False}
                    save_data(users, USERS_FILE)
                    st.success("Conta criada com sucesso! Faça login para continuar.")
                    st.session_state['show_create_account'] = False
                    st.experimental_rerun()
        if st.form_submit_button("Voltar para Login"):
            st.session_state['show_create_account'] = False
            st.experimental_rerun()

# --- Conteúdo do Aplicativo ---
def app_content():
    st.sidebar.title(f"Bem-vindo(a), {st.session_state['username']}")
    st.sidebar.button("Logout", on_click=lambda: st.session_state.clear())
    st.title("Gerador de Parecer Descritivo Individual")
    
    # Adicionando o seletor de tipo de parecer
    parecer_type = st.radio("Selecione o tipo de parecer a ser gerado:",
                            ["Parecer Completo", "Apenas o Texto"])

    # Carregar dados
    pareceres_db = load_data(PARECERES_FILE, {})
    
    # Coletar dados do aluno
    st.header("1. Dados do Aluno")
    student_name = st.text_input("Nome Completo do Aluno(a)")
    filiacao_mae = st.text_input("Filiação (Mãe)")
    filiacao_pai = st.text_input("Filiação (Pai)")
    data_nascimento = st.text_input("Data de Nascimento (dd/mm/aaaa)")
    naturalidade = st.text_input("Naturalidade")
    endereco = st.text_input("Endereço")
    uf = st.text_input("UF")
    numero_aluno = st.text_input("Número do Aluno(a) na Lista de Chamada")
    
    st.header("2. Dados da Turma")
    periodo = st.selectbox("Período/Ano", ["1º", "2º", "3º", "4º", "5º", "6º", "7º", "8º", "9º", "EJA"])
    turma = st.text_input("Turma")
    turno = st.selectbox("Turno", ["Manhã", "Tarde", "Noturno"])
    
    st.header("3. Dados do Parecer")
    nome_professor = st.text_input("Nome Completo do(a) Professor(a)")
    
    # Opção para alunos que deixaram de frequentar - NOVO CAMPO DE SELEÇÃO
    status_aluno = st.selectbox(
        "Status do Aluno(a)",
        ["Frequente", "Deixou de Frequentar"] # NOVO CAMPO
    )

    if status_aluno == "Frequente":
        st.subheader("Habilidades Desenvolvidas")
        foco_parecer = st.radio("Foco do Parecer", ["Geral", "Comportamental", "Específico"])
    
        if foco_parecer == "Geral":
            with st.container(border=True):
                st.write("#### Áreas de Desempenho")
                leitura = st.selectbox("Leitura e Escrita", ["N/A", "Muito Bom", "Bom", "Precisa Melhorar"])
                matematica = st.selectbox("Matemática", ["N/A", "Muito Bom", "Bom", "Precisa Melhorar"])
                comportamento = st.selectbox("Comportamento", ["N/A", "Muito Bom", "Bom", "Precisa Melhorar"])
                participacao = st.selectbox("Participação", ["N/A", "Muito Bom", "Bom", "Precisa Melhorar"])
        
        elif foco_parecer == "Comportamental":
            with st.container(border=True):
                st.write("#### Foco Comportamental")
                comportamento = st.selectbox("Comportamento", ["N/A", "Muito Bom", "Bom", "Precisa Melhorar"])
                participacao = st.selectbox("Participação", ["N/A", "Muito Bom", "Bom", "Precisa Melhorar"])
                leitura = matematica = "N/A"
        
        elif foco_parecer == "Específico":
            with st.container(border=True):
                st.write("#### Foco Específico (Personalizado)")
                parecer_personalizado = st.text_area("Digite o texto do parecer:")
                leitura = matematica = comportamento = participacao = "N/A"
    
    # Botão de Geração do Parecer
    if st.button("Gerar Parecer"):
        # --- Lógica de Geração do Parecer ---
        if not student_name or not nome_professor:
            st.error("Por favor, preencha o nome do aluno e do professor.")
        else:
            parecer_text = ""
            
            # NOVO TRECHO DE CÓDIGO - TRATA ALUNOS QUE DEIXARAM DE FREQUENTAR
            if status_aluno == "Deixou de Frequentar":
                parecer_text = (
                    f"Constatou-se que o(a) aluno(a) {student_name} deixou de frequentar a escola. "
                    "Devido ao curto período de tempo de sua presença em sala de aula, não foi possível "
                    "estabelecer uma relação de aprendizado sólida."
                )
            
            elif status_aluno == "Frequente":
                if foco_parecer == "Geral":
                    parecer_text = f"O(a) aluno(a) {student_name} demonstra uma postura colaborativa em sala de aula, apresentando um bom desempenho geral. No que diz respeito à leitura e escrita, demonstra {leitura}. Na área de Matemática, seu desempenho é {matematica}. Seu comportamento é {comportamento} e sua participação em atividades em grupo é {participacao}."
                elif foco_parecer == "Comportamental":
                    parecer_text = f"O(a) aluno(a) {student_name} demonstra uma postura {comportamento} em sala de aula. Sua participação nas atividades é {participacao}, contribuindo positivamente para o ambiente escolar."
                elif foco_parecer == "Específico":
                    parecer_text = parecer_personalizado
            
            data_parecer = datetime.now().strftime("%d de %B de %Y").replace(
                "January", "janeiro"
            ).replace("February", "fevereiro").replace("March", "março").replace(
                "April", "abril"
            ).replace("May", "maio").replace("June", "junho").replace(
                "July", "julho"
            ).replace("August", "agosto").replace("September", "setembro").replace(
                "October", "outubro"
            ).replace("November", "novembro").replace("December", "dezembro")

            parecer_data = {
                'nome_aluno': student_name,
                'filiacao_mae': filiacao_mae,
                'filiacao_pai': filiacao_pai,
                'data_nascimento': data_nascimento,
                'naturalidade': naturalidade,
                'endereco': endereco,
                'uf': uf,
                'numero_aluno': numero_aluno,
                'periodo': periodo,
                'turma': turma,
                'turno': turno,
                'ano_letivo': datetime.now().year,
                'semestre': '2025.1', # Ajuste conforme o ano
                'nome_professor': nome_professor,
                'texto_parecer': parecer_text,
                'data_parecer': data_parecer
            }

            if 'parecer_personalizado' in locals() and parecer_personalizado:
                parecer_data['opcao'] = 'Específico'
            elif status_aluno == 'Deixou de Frequentar':
                parecer_data['opcao'] = 'Deixou de Frequentar'
            else:
                parecer_data['characteristics_levels'] = {
                    "leitura": leitura,
                    "matematica": matematica,
                    "comportamento": comportamento,
                    "participacao": participacao
                }

            # Gerar o arquivo .docx
            docx_data_bytes = None
            if parecer_type == "Parecer Completo":
                doc_bytes_io = create_parecer_docx(parecer_data)
                if doc_bytes_io:
                    docx_data_bytes = doc_bytes_io.getvalue()
                    parecer_data['docx_data'] = docx_data_bytes.hex()
                    
            st.success("Parecer gerado com sucesso!")
            st.write("---")
            st.subheader("Pré-visualização do Parecer")
            st.write(parecer_data['texto_parecer'])
            
            if docx_data_bytes:
                sanitized_name = sanitize_student_name_for_filename(student_name)
                file_name = f"parecer_descritivo_{sanitized_name}.docx"
                st.download_button(
                    label="Baixar Parecer (DOCX)",
                    data=docx_data_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            # Salvar no banco de dados
            if st.session_state['username'] not in pareceres_db:
                pareceres_db[st.session_state['username']] = {}
            if student_name not in pareceres_db[st.session_state['username']]:
                pareceres_db[st.session_state['username']][student_name] = []
            
            parecer_data_to_save = parecer_data.copy()
            if docx_data_bytes:
                parecer_data_to_save['docx_data'] = docx_data_bytes.hex()
            
            pareceres_db[st.session_state['username']][student_name].append(parecer_data_to_save)
            save_data(pareceres_db, PARECERES_FILE)


def admin_dashboard():
    # --- Lógica do Painel de Admin ---
    st.title("Painel do Administrador")
    st.sidebar.title(f"Bem-vindo, {st.session_state['username']}")
    st.sidebar.button("Logout", on_click=lambda: st.session_state.clear())
    
    st.header("Gerenciar Usuários")
    users_db = load_data(USERS_FILE, {})
    user_list = [user for user in users_db if user != st.session_state['username']]

    if user_list:
        st.write("### Usuários Existentes")
        for user in user_list:
            is_admin = users_db[user].get('is_admin', False)
            col1, col2, col3 = st.columns([0.4, 0.2, 0.4])
            with col1:
                st.write(f"- {user}")
            with col2:
                if st.button("Tornar Admin", key=f"make_admin_{user}"):
                    users_db[user]['is_admin'] = True
                    save_data(users_db, USERS_FILE)
                    st.success(f"Usuário {user} agora é administrador.")
                    st.experimental_rerun()
            with col3:
                if st.button("Remover", key=f"remove_user_{user}"):
                    del users_db[user]
                    save_data(users_db, USERS_FILE)
                    st.success(f"Usuário {user} removido com sucesso.")
                    st.experimental_rerun()
            st.markdown("---")
    else:
        st.info("Nenhum outro usuário cadastrado.")

    st.header("Gerenciar Pareceres")
    pareceres_db = load_data(PARECERES_FILE, {})
    
    if pareceres_db:
        st.write("### Pareceres Salvos por Usuário")
        for username_parecer, user_data in pareceres_db.items():
            st.subheader(f"Pareceres de {username_parecer}")
            for student_name_display, pareceres in user_data.items():
                st.markdown(f"**Aluno(a):** {student_name_display}")
                for i, parecer_info in enumerate(pareceres):
                    with st.expander(f"Parecer {i+1} - {parecer_info.get('data_parecer', 'Data Indisponível')}"):
                        st.markdown(f"**Professor(a):** {parecer_info['nome_professor']}")
                        st.markdown(f"**Texto:** {parecer_info['texto_parecer']}")
                        
                        if 'characteristics_levels' in parecer_info and parecer_info['characteristics_levels']:
                            st.write(f"**Leitura e Escrita:** {parecer_info['characteristics_levels'].get('leitura', 'N/A')}")
                            st.write(f"**Matemática:** {parecer_info['characteristics_levels'].get('matematica', 'N/A')}")
                            st.write(f"**Comportamento:** {parecer_info['characteristics_levels'].get('comportamento', 'N/A')}")
                            st.write(f"**Participação:** {parecer_info['characteristics_levels'].get('participacao', 'N/A')}")
                        elif 'opcao' in parecer_info:
                             st.write(f"**Opção Geral:** {parecer_info['opcao']}")

                        if 'docx_data' in parecer_info and parecer_info['docx_data']:
                            try:
                                docx_data_bytes = bytes.fromhex(parecer_info['docx_data'])
                                file_name = f"parecer_{sanitize_student_name_for_filename(student_name_display)}_{parecer_info['data_parecer'].replace(' ', '_').replace(':', '')}_{i}.docx"
                                st.download_button(
                                    label=f"Baixar Parecer {i+1} (DOCX)",
                                    data=docx_data_bytes,
                                    file_name=file_name,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key=f"admin_download_docx_{student_name_display}_{i}"
                                )
                            except ValueError:
                                st.error(f"Erro ao carregar DOCX para o parecer {i+1}. Dados corrompidos.")
                        else:
                            st.info(f"DOCX não disponível para o parecer {i+1}.")
                        st.markdown("---")
    else:
        st.info("Nenhum parecer salvo ainda.")

    st.markdown("---")
    st.info("Sistema desenvolvido com Streamlit e Python.")


# --- Lógica principal ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'show_create_account' not in st.session_state:
    st.session_state['show_create_account'] = False

users_db = load_data(USERS_FILE, {})
current_user = st.session_state.get('username')

if st.session_state['logged_in']:
    if users_db.get(current_user, {}).get('is_admin', False):
        admin_dashboard()
    else:
        app_content()
else:
    if st.session_state['show_create_account']:
        create_account()
    else:
        login()
