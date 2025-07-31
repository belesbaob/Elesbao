import streamlit as st
import json
import os
from datetime import datetime
from io import BytesIO
from docx import Document
import unicodedata
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sys # <--- Adicione esta linha (Import sys)

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
DATA_DIR = os.path.join(base_path, "data") # <--- AQUI ESTÁ A MUDANÇA PRINCIPAL
USERS_FILE = os.path.join(DATA_DIR, "users.json")
PARECERES_FILE = os.path.join(DATA_DIR, "pareceres.json")
SCHOOL_NAME = "ESCOLA MUNICIPAL DE EDUCAÇÃO FUNDAMENTAL ELESBÃO BARBOSA DE CARVALHO"
COORDENADOR_NAME = "NOME DO COORDENADOR AQUI"

# Lista de Nomes de Alunos (mantida a mesma)
STUDENT_NAMES = [
    "Andressa Viturino dos Santos", "Carlos Eduardo Conceição da Silva", "Damiana Honorato dos Santos",
    "Ivonete Nunes da Silva", "João Pedro da Silva", "José Alexandre Lemo da Silva",
    "Joselma Conceição da Silva", "José Gonzaga de Melo", "Maria Angelica Honorato de Oliveira",
    "Maria das Dores Rodrigues Pereira", "Nelson Pereira da Silva", "Nivaldilson Vitorino da Silva",
    "Roberlange Pereira da Silva", "Rosania Valério da Silva", "Rosileide Ferreira da Silva",
    "Thayse Ferreira dos Santos", "Alessandra Lopes da Silva", "Adenilson Lourenço dos Santos",
    "Antônio Ronaldo Firmino", "Cicero da Silva", "Cristiano da Conceição Nogueira",
    "Daiana Conceição da Silva", "Daniela da Conceição Nogueira", "Edimilson Lima de Queiroz",
    "Eraldo Bernardino Gomes", "Ivanilda Ferreira dos Anjos", "João Batista Aureliano da Silva",
    "Joana D’arc Gouveia da Silva", "José Eraldo Silva", "Juliana Conceição da Silva",
    "Leonardo da Silva", "Maria de Lourdes Emerinda da Silva", "Maria Paula Viturino dos Santos",
    "Maria Pereira da Silva", "Maria Rosilma da Conceição", "Quiteria Viturino Santos Barros",
    "Vandilma dos Santos Silva", "Wellison Gomes Ferreira", "Ana Maria da Silva Aquino",
    "Bartolomeu Ferreira Vanderlei", "Genilson Alves Araújo", "Nilda Barbosa Silva Santos",
    "Vilma Maria de França", "Adalva Gomes dos Santos", "Adriano Viturino dos Santos",
    "Aristeu Vitorino dos Santos", "Aurene de Melo Gonzaga", "Carlos Pereira da Silva",
    "Cristóvão Pereira dos Santos", "Davino Conceição dos Santos", "Edineuza Teles da Silva",
    "Estelita da Silva", "Francisco Lourenço da Silva", "Francisco Viturino dos Santos",
    "Inês Juliana dos Santos", "Jânio Gonzaga de Melo", "José Adelmo Soares da Silva",
    "José Edijario Soares Lemos", "José França da Silva", "José Honélio dos Santos",
    "José Nilton Avelino", "José Ronaldo dos Santos", "José Vitorino Júnior",
    "Joselia Honorato de Melo", "Jucilene de Melo Santos", "Manoel Alves Araújo",
    "Marcio José Rodrigues Limeira", "Maria Aparecida dos Anjos Viturino",
    "Maria Cleide da Silva", "Maria Gomes dos Santos", "Maria Jaqueline dos Santos Silva",
    "Rosilânia da Silva Honorato", "Sebastião Rodrigues dos Santos", "Valdimira Gomes"
]

os.makedirs(DATA_DIR, exist_ok=True)

# --- Funções de Ajuda ---
def load_data(filepath):
    if not os.path.exists(filepath) or os.path.getsize(filepath) == 0:
        return []
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data, filepath):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def sanitize_student_name_for_filename(name):
    """
    Sanitiza o nome do aluno para ser usado como parte de um nome de arquivo.
    Removes acentos, converte para minúsculas e substitui espaços por underscores.
    """
    name = unicodedata.normalize('NFKD', name).encode('ascii', 'ignore').decode('utf-8')
    name = name.lower()
    name = name.replace(" ", "_")
    name = "".join(c for c in name if c.isalnum() or c == '_')
    return name

# --- FUNÇÃO PARA GERAR O TEXTO DETALHADO DO PARECER (Com texto mais conciso) ---
def generate_detailed_parecer_text(characteristics_levels, student_name):
    """
    Gera um texto detalhado e contínuo para o parecer, integrando todas as características
    e buscando aproximar-se de 100-150 palavras.
    """
    comportamento = characteristics_levels["comportamento"]
    participacao = characteristics_levels["participacao"]
    leitura_escrita = characteristics_levels["leitura_escrita"]
    matematica = characteristics_levels["matematica"]

    parecer_parts = []

    # Introdução geral (Aprox. 20 palavras)
    parecer_parts.append(f"Este parecer detalha o desenvolvimento de {student_name} no período letivo, abordando seu progresso acadêmico e social.")

    # Detalhes sobre Comportamento (Aprox. 20-30 palavras cada)
    if comportamento == "Ótimo":
        parecer_parts.append("Demonstra comportamento exemplar, contribuindo ativamente para um ambiente de aprendizado positivo e harmonioso. Sua postura disciplinada inspira os colegas.")
    elif comportamento == "Bom":
        parecer_parts.append("Comportamento consistentemente bom, respeitando normas e mantendo conduta adequada. Contribui positivamente para o bom andamento das atividades.")
    elif comportamento == "Regular":
        parecer_parts.append("Comportamento geralmente adequado, mas com momentos de distração, necessitando de lembretes para manter o foco. Há espaço para aprimoramento na autorregulação.")
    else: # Ruim
        parecer_parts.append("Comportamento desafiador em sala de aula, com dificuldades em seguir regras e focar, gerando interrupções. Intervenções específicas são cruciais.")

    # Detalhes sobre Participação (Aprox. 20-30 palavras cada)
    if participacao == "Ótimo":
        parecer_parts.append("Participação notavelmente ativa e pertinente, com grande interesse pelos conteúdos. Realiza perguntas perspicazes e oferece contribuições valiosas.")
    elif participacao == "Bom":
        parecer_parts.append("Participa de forma consistente, mostrando interesse e esforço em contribuir. Busca interagir e se envolver para aprofundar seu entendimento.")
    elif participacao == "Regular":
        parecer_parts.append("A participação é pontual e ocasional. Demonstra potencial, mas por vezes reticência em se expressar. Incentivos adicionais podem estimular maior engajamento.")
    else: # Ruim
        parecer_parts.append("Participação mínima em sala de aula, com pouca iniciativa para interagir. Essa passividade limita o aproveitamento e a consolidação do aprendizado.")

    # Detalhes sobre Leitura e Escrita (Aprox. 20-30 palavras cada)
    if leitura_escrita == "Ótimo":
        parecer_parts.append("Excelente capacidade de leitura e escrita, com compreensão aprofundada de textos complexos. Produz textos coesos, coerentes e bem estruturados, com vocabulário rico.")
    elif leitura_escrita == "Bom":
        parecer_parts.append("Boa habilidade em leitura e escrita, compreendendo a maioria dos textos e expressando-se claramente. Produz redações com ideias bem definidas, com refinamentos pontuais.")
    elif leitura_escrita == "Regular":
        parecer_parts.append("Nível regular de leitura e escrita, com dificuldades na compreensão de nuances ou elaboração de frases complexas. Práticas direcionadas são necessárias para maior autonomia.")
    else: # Ruim
        parecer_parts.append("Significativas dificuldades em leitura e escrita, impactando a compreensão e produção de ideias. Esforços contínuos e intervenções pedagógicas específicas são cruciais.")

    # Detalhes sobre Matemática (Aprox. 20-30 palavras cada)
    if matematica == "Ótimo":
        parecer_parts.append("Excelente domínio das operações e conceitos matemáticos. Resolve problemas com autonomia, aplicando diferentes estratégias e justificando raciocínios logicamente. Aptidão evidente.")
    elif matematica == "Bom":
        parecer_parts.append("Boa capacidade em operações matemáticas básicas e compreensão dos conceitos. Aplica o conhecimento em diversas situações, mostrando solidez em sua base.")
    elif matematica == "Regular":
        parecer_parts.append("Realiza operações matemáticas básicas com alguma dificuldade, demandando tempo e suporte. Revisão de conceitos e prática constante são recomendadas para fortalecer proficiência.")
    else: # Ruim
        parecer_parts.append("Grande dificuldade em operações e conceitos matemáticos básicos. Domínio numérico e lógico limitado, exigindo plano de intervenção intensivo para construir base sólida.")

    # Nova Lógica de Reprovação (1. Aluno só reprova se Leitura/Escrita E Matemática forem Ruim)
    reprovado = False
    if leitura_escrita == "Ruim" and matematica == "Ruim":
        reprovado = True
    
    ressalvas = False
    for level in characteristics_levels.values():
        if level == "Regular":
            ressalvas = True
            break # Se houver qualquer "Regular", já marca como ressalvas

    # Conclusão do parecer (Aprox. 25-35 palavras cada)
    if reprovado:
        conclusao = "Diante do exposto, e considerando os desafios persistentes em leitura e matemática, o aluno não atingiu os critérios necessários para aprovação neste período letivo."
    elif ressalvas:
        conclusao = "Aluno(a) aprovado(a) com ressalvas, sendo crucial que o acompanhamento pedagógico e as intervenções específicas sejam mantidas. O foco deve ser nas áreas que demandam maior desenvolvimento para progresso consistente."
    else:
        conclusao = "Conclui-se que o aluno(a) foi aprovado(a). Seu desempenho e desenvolvimento geral indicam que atingiu os objetivos propostos para o período, demonstrando preparo para avançar para a próxima etapa."

    parecer_parts.append(conclusao)

    # Junta todas as partes em um texto único
    full_text = "\n\n".join(parecer_parts)
    # Substitui o nome do aluno em todas as partes da descrição gerada
    full_text = full_text.replace("{student_name}", student_name)
    return full_text


# --- FUNÇÃO GERAR_DOCX_PARECER ---
def gerar_docx_parecer(student_name, characteristics_levels, teacher_name):
    """
    Gera o documento DOCX do parecer preenchendo o template específico do aluno.
    """
    sanitized_name = sanitize_student_name_for_filename(student_name)
    student_template_file = os.path.join(DATA_DIR, f"template_{sanitized_name}.docx")

    if not os.path.exists(student_template_file):
        st.error(f"Erro: O template específico para '{student_name}' não foi encontrado.")
        st.info(f"Por favor, crie o arquivo '{os.path.basename(student_template_file)}' na pasta 'data/' com as informações pessoais do aluno e os placeholders necessários.")
        return None

    try:
        document = Document(student_template_file)

        full_parecer_text = generate_detailed_parecer_text(characteristics_levels, student_name)
        
        current_date = datetime.now()
        dia = current_date.strftime("%d")
        mes_extenso = current_date.strftime("%B").replace("January", "Janeiro").replace("February", "Fevereiro").replace("March", "Março").replace("April", "Abril").replace("May", "Maio").replace("June", "Junho").replace("July", "Julho").replace("August", "Agosto").replace("September", "Setembro").replace("October", "Outubro").replace("November", "Novembro").replace("December", "Dezembro")
        ano = current_date.strftime("%Y")

        replacements = {
            "{{NOME_ALUNO}}": student_name,
            "{{PARECER_GERADO}}": full_parecer_text,
            "{{NOME_PROFESSOR}}": teacher_name,
            "{{NOME_COORDENADOR}}": COORDENADOR_NAME, # Novo placeholder para coordenador
            "{{DATA_PARECER}}": current_date.strftime("%d/%m/%Y"),
            "{{DIA_PARECER}}": dia,
            "{{MES_PARECER}}": mes_extenso,
            "{{ANO_CORRENTE}}": ano,
            "{{SEMESTRE}}": "2025.1"
        }

        # Função para substituir placeholders em parágrafos, aplicar justificação e definir fonte
        def replace_and_format_paragraph(paragraph, old_text, new_text):
            full_paragraph_text = "".join(run.text for run in paragraph.runs)
            
            if old_text in full_paragraph_text:
                # Limpa os runs existentes no parágrafo
                for i in range(len(paragraph.runs) - 1, -1, -1):
                    paragraph._element.remove(paragraph.runs[i]._element)
                
                # Adiciona o novo texto como um único run
                new_run = paragraph.add_run(new_text)

                # Aplica justificação e define a fonte Times New Roman se for o placeholder do parecer gerado
                if old_text == "{{PARECER_GERADO}}":
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    new_run.font.name = "Times New Roman" # Define a fonte Times New Roman

        # Iterar sobre todos os parágrafos do documento e substituir os placeholders
        for paragraph in document.paragraphs:
            for key, value in replacements.items():
                replace_and_format_paragraph(paragraph, key, value)

        # Iterar sobre as tabelas (se houver) e substituir os placeholders
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for key, value in replacements.items():
                            replace_and_format_paragraph(paragraph, key, value)

        buffer = BytesIO()
        document.save(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Ocorreu um erro ao preencher o DOCX: {e}")
        st.info("Verifique se o template DOCX está correto e se o nome dos placeholders está exato.")
        st.info("Certifique-se de que a biblioteca 'python-docx' está instalada (`pip install python-docx`).")
        return None


# --- Carregar Usuários (ou criar se não existirem) ---
def initialize_users():
    users = {}
    if os.path.exists(USERS_FILE) and os.path.getsize(USERS_FILE) == 0:
        return users # Retorna vazio se o arquivo não existe ou está vazio

    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            users = json.load(f)
    except json.JSONDecodeError:
        st.error("Erro ao ler o arquivo de usuários. Ele pode estar corrompido ou vazio. Recriando usuários padrão.")
        users = {} # Força a recriação se houver erro de leitura

    if not users:
        users = {
            "professor1": {"password": "p1", "role": "professor"},
            "professor2": {"password": "p2", "role": "professor"},
            "professor3": {"password": "p3", "role": "professor"},
            "admin": {"password": "adminpass", "role": "admin"}
        }
        save_data(users, USERS_FILE)
    return users

users = initialize_users()

# Carregar pareceres existentes
pareceres_salvos = load_data(PARECERES_FILE)

# --- Layout do Streamlit ---
st.set_page_config(
    page_title="Sistema de Pareceres de Alunos",
    layout="centered",
    initial_sidebar_state="auto"
)

st.title("Sistema de Pareceres de Alunos")
st.markdown("---")

# --- Tela de Login ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.role = None

if not st.session_state.logged_in:
    st.header("Login")
    username = st.text_input("Usuário", key="login_user")
    password = st.text_input("Senha", type="password", key="login_pass")

    if st.button("Entrar", key="login_button"):
        if username in users and users.get(username, {}).get("password") == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = users.get(username, {}).get("role")
            st.success(f"Bem-vindo(a), {username}!")
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")
else:
    st.sidebar.write(f"Usuário logado: **{st.session_state.username}**")
    st.sidebar.write(f"Papel: **{st.session_state.role.capitalize()}**")
    if st.sidebar.button("Sair", key="logout_button"):
        st.session_state.logged_in = False
        st.session_state.username = None
        st.session_state.role = None
        st.rerun()

    if st.session_state.role == "professor":
        st.header("Gerar Parecer de Aluno")

        selected_student = st.selectbox(
            "Selecione o Aluno:",
            [""] + sorted(STUDENT_NAMES),
            key="student_name_select"
        )

        st.subheader("Avaliação das Características:")
        levels_options = ["Bom", "Ótimo", "Regular", "Ruim"]

        comportamento_level = st.selectbox("Comportamento:", levels_options, key="comportamento_level")
        participacao_level = st.selectbox("Participação em sala de aula:", levels_options, key="participacao_level")
        leitura_escrita_level = st.selectbox("Capacidade de leitura e escrita:", levels_options, key="leitura_escrita_level")
        matematica_level = st.selectbox("Operações matemáticas básicas:", levels_options, key="matematica_level")

        if st.button("Gerar e Salvar Parecer em DOCX", key="generate_save_docx_button"):
            if selected_student:
                characteristics_levels = {
                    "comportamento": comportamento_level,
                    "participacao": participacao_level,
                    "leitura_escrita": leitura_escrita_level,
                    "matematica": matematica_level
                }
                
                docx_buffer = gerar_docx_parecer(selected_student, characteristics_levels, st.session_state.username)
                
                if docx_buffer:
                    docx_bytes = docx_buffer.getvalue()
                    file_name = f"parecer_{sanitize_student_name_for_filename(selected_student)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

                    pareceres_salvos.append({
                        "student_name": selected_student,
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "professor": st.session_state.username,
                        "characteristics_levels": characteristics_levels,
                        "docx_data": docx_bytes.hex()
                    })
                    save_data(pareceres_salvos, PARECERES_FILE)

                    st.download_button(
                        label="Baixar Parecer Gerado (DOCX)",
                        data=docx_bytes,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    st.success(f"Parecer para **{selected_student}** gerado e salvo com sucesso!")
            else:
                st.warning("Por favor, selecione o nome do aluno para gerar o parecer.")

    elif st.session_state.role == "admin":
        st.header("Visualizar e Baixar Pareceres")

        if not pareceres_salvos:
            st.info("Nenhum parecer salvo ainda.")
        else:
            alunos_com_pareceres = sorted(list(set(p.get('student_name') for p in pareceres_salvos if 'student_name' in p)))

            if not alunos_com_pareceres:
                st.info("Nenhum parecer salvo ainda com nome de aluno.")
            else:
                selected_student_admin = st.selectbox(
                    "Selecione um aluno para visualizar os pareceres:",
                    [""] + alunos_com_pareceres,
                    key="admin_student_select"
                )

                pareceres_a_exibir = []
                if not selected_student_admin:
                    pareceres_a_exibir = pareceres_salvos
                else:
                    pareceres_a_exibir = [p for p in pareceres_salvos if p.get('student_name') == selected_student_admin]

                if not pareceres_a_exibir:
                    st.info(f"Nenhum parecer encontrado para {selected_student_admin}.")
                else:
                    if not selected_student_admin:
                        st.subheader("Todos os Pareceres Salvos:")
                    else:
                        st.subheader(f"Pareceres para {selected_student_admin}:")

                    for i, parecer_info in enumerate(pareceres_a_exibir):
                        student_name_display = parecer_info.get('student_name', 'Aluno Desconhecido')
                        st.markdown(f"**Parecer {i+1} para {student_name_display}**")
                        st.write(f"**Data:** {parecer_info['data']}")
                        st.write(f"**Professor:** {parecer_info['professor']}")
                        
                        if 'characteristics_levels' in parecer_info:
                            st.write(f"**Níveis Avaliados:**")
                            st.write(f"  - Comportamento: {parecer_info['characteristics_levels'].get('comportamento', 'N/A')}")
                            st.write(f"  - Participação: {parecer_info['characteristics_levels'].get('participacao', 'N/A')}")
                            st.write(f"  - Leitura/Escrita: {parecer_info['characteristics_levels'].get('leitura_escrita', 'N/A')}")
                            st.write(f"  - Matemática: {parecer_info['characteristics_levels'].get('matematica', 'N/A')}")
                        elif 'opcao' in parecer_info:
                             st.write(f"**Opção Geral (Legado):** {parecer_info['opcao']}")

                        if 'docx_data' in parecer_info and parecer_info['docx_data']:
                            try:
                                docx_data_bytes = bytes.fromhex(parecer_info['docx_data'])
                                file_name = f"parecer_{sanitize_student_name_for_filename(student_name_display)}_{parecer_info['data'].replace(' ', '_').replace(':', '')}_{i}.docx"
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

    st.markdown("---")
    st.info("Sistema desenvolvido com Streamlit e Python.")