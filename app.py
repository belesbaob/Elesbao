import streamlit as st
import json
import os
from datetime import datetime
from io import BytesIO
from docx import Document
import unicodedata
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sys

# Verifica se o aplicativo está rodando como um executável PyInstaller
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

# --- Configurações de Diretórios e Arquivos ---
DATA_DIR = os.path.join(base_path, "data")
TEMPLATES_DIR = os.path.join(base_path, "templates") # Novo diretório para os modelos
USERS_FILE = os.path.join(DATA_DIR, "users.json")
PARECERES_FILE = os.path.join(DATA_DIR, "pareceres.json")
SCHOOL_NAME = "ESCOLA MUNICIPAL DE EDUCAÇÃO FUNDAMENTAL ELESBÃO BARBOSA DE CARVALHO"

# --- Funções Auxiliares ---

def sanitize_student_name_for_filename(name):
    """
    Remove acentos, espaços e caracteres especiais para criar um nome de arquivo seguro.
    Ex: "João da Silva" -> "joao_da_silva"
    """
    s = str(name).strip().lower()
    s = unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode("utf-8")
    s = s.replace(' ', '_')
    s = ''.join(c for c in s if c.isalnum() or c == '_')
    return s

def load_data(file_path):
    if not os.path.exists(file_path):
        return {}
    with open(file_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_data(data, file_path):
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def get_student_names(pareceres_data):
    return sorted(pareceres_data.keys())

def generate_parecer_docx(parecer_info, template_path):
    try:
        if not os.path.exists(template_path):
            st.error(f"Erro: O arquivo de modelo não foi encontrado em: {template_path}")
            return None

        doc = Document(template_path)

        # Substituições no documento
        substitutions = {
            'NOME_ALUNO': parecer_info['nome'],
            'NOME_MAE': parecer_info['filiacao_mae'],
            'NOME_PAI': parecer_info['filiacao_pai'],
            'ENDERECO': parecer_info['endereco'],
            'NATURALIDADE': parecer_info['naturalidade'],
            'NASCIMENTO': parecer_info['data_nascimento'],
            'PERIODO': parecer_info['periodo'],
            'TURNO': parecer_info['turno'],
            'PARECER_TEXTO': parecer_info['parecer_texto'],
            'DATA_PARECER': parecer_info['data'],
            'OBSERVACAO': parecer_info['observacao']
        }

        for paragraph in doc.paragraphs:
            for key, value in substitutions.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(value))
            # Ajustar alinhamento após a substituição
            if paragraph.text.strip():
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Salva o documento em um objeto BytesIO na memória
        doc_buffer = BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)
        return doc_buffer

    except Exception as e:
        st.error(f"Erro ao gerar o documento: {e}")
        return None

# --- Main App ---

st.set_page_config(page_title="Gerador de Parecer Escolar", layout="centered")

st.title("Sistema de Gestão de Pareceres")

# Tente carregar os dados
pareceres = load_data(PARECERES_FILE)

# Sidebar
st.sidebar.title("Navegação")
page = st.sidebar.radio("Ir para:", ["Formulário", "Relatórios Salvos"])

# --- Formulário de Geração ---
if page == "Formulário":
    st.header("Gerar Novo Parecer")

    st.markdown("---")
    st.markdown("#### Informações do Aluno")

    student_name = st.text_input("Nome do Aluno(a)")
    filiacao_mae = st.text_input("Filiação (Mãe)")
    filiacao_pai = st.text_input("Filiação (Pai)")
    endereco = st.text_input("Endereço")
    naturalidade = st.text_input("Naturalidade")
    data_nascimento = st.date_input("Data de Nascimento")
    periodo = st.text_input("Período")
    turno = st.text_input("Turno")

    st.markdown("---")
    st.markdown("#### Conteúdo do Parecer")

    parecer_texto = st.text_area("Digite o texto do parecer aqui:", height=250)
    observacao = st.text_area("Observações (opcional):", height=100)

    is_nao_frequentou = st.checkbox("Aluno(a) deixou de frequentar?", value=False, help="Marque esta opção se o aluno não frequentou a escola.")

    if st.button("Gerar e Salvar Parecer"):
        if not student_name or not parecer_texto:
            st.error("Por favor, preencha o nome do aluno e o texto do parecer.")
        else:
            if student_name not in pareceres:
                pareceres[student_name] = []

            # Determine qual modelo usar
            if is_nao_frequentou:
                template_file = "template_nao_frequentou.docx" # Crie um modelo específico para este caso
            else:
                sanitized_name = sanitize_student_name_for_filename(student_name)
                template_file = f"template_{sanitized_name}.docx"

            template_path = os.path.join(TEMPLATES_DIR, template_file)

            if not os.path.exists(template_path):
                st.warning(f"Atenção: Modelo para '{student_name}' não encontrado. Verifique se o arquivo '{template_file}' existe na pasta 'templates'.")
                # Se não houver template específico, pode usar um genérico ou continuar
                template_path = os.path.join(TEMPLATES_DIR, "parecer_template.docx") # Modelo genérico
                if not os.path.exists(template_path):
                    st.error("Nenhum modelo específico ou genérico encontrado. A geração do documento foi cancelada.")
                    st.stop()

            # Estrutura do parecer
            parecer_data = {
                "id": len(pareceres[student_name]) + 1,
                "nome": student_name,
                "filiacao_mae": filiacao_mae,
                "filiacao_pai": filiacao_pai,
                "endereco": endereco,
                "naturalidade": naturalidade,
                "data_nascimento": data_nascimento.strftime("%d/%m/%Y"),
                "periodo": periodo,
                "turno": turno,
                "parecer_texto": parecer_texto,
                "observacao": observacao,
                "data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "nao_frequentou": is_nao_frequentou
            }

            # Geração do DOCX
            docx_bytes = generate_parecer_docx(parecer_data, template_path)
            if docx_bytes:
                parecer_data['docx_data'] = docx_bytes.hex()
                pareceres[student_name].append(parecer_data)
                save_data(pareceres, PARECERES_FILE)
                st.success("Parecer gerado e salvo com sucesso!")

                # Botão de download
                sanitized_filename = sanitize_student_name_for_filename(student_name)
                download_filename = f"parecer_{sanitized_filename}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                st.download_button(
                    label="Baixar Parecer (DOCX)",
                    data=docx_bytes,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

# --- Relatórios Salvos ---
elif page == "Relatórios Salvos":
    st.header("Relatórios Salvos")

    student_list = get_student_names(pareceres)
    if not student_list:
        st.info("Nenhum parecer salvo ainda.")
    else:
        selected_student = st.selectbox("Selecione o Aluno(a):", [""] + student_list)

        if selected_student:
            student_pareceres = pareceres.get(selected_student, [])
            st.subheader(f"Pareceres de {selected_student}")
            if not student_pareceres:
                st.info("Nenhum parecer encontrado para este aluno.")
            else:
                for i, parecer_info in enumerate(student_pareceres):
                    st.markdown(f"**Parecer {i+1}** - Salvo em: {parecer_info['data']}")
                    st.write(f"**Nome:** {parecer_info['nome']}")
                    st.write(f"**Parecer:** {parecer_info['parecer_texto']}")
                    
                    if parecer_info['nao_frequentou']:
                        st.info("Este parecer é para um aluno que deixou de frequentar.")

                    if 'docx_data' in parecer_info and parecer_info['docx_data']:
                        try:
                            docx_data_bytes = bytes.fromhex(parecer_info['docx_data'])
                            file_name = f"parecer_{sanitize_student_name_for_filename(selected_student)}_{parecer_info['data'].replace(' ', '_').replace(':', '')}_{i}.docx"
                            st.download_button(
                                label=f"Baixar Parecer {i+1} (DOCX)",
                                data=docx_data_bytes,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"admin_download_docx_{selected_student}_{i}"
                            )
                        except ValueError:
                            st.error(f"Erro ao carregar DOCX para o parecer {i+1}. Dados corrompidos.")
                    else:
                        st.info(f"DOCX não disponível para o parecer {i+1}.")
                    st.markdown("---")
