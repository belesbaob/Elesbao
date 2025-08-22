import streamlit as st
import json
import os
from datetime import datetime
from io import BytesIO
from docx import Document
import unicodedata
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sys

# --- Configurações Iniciais ---
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

DATA_DIR = os.path.join(base_path, "data")
USERS_FILE = os.path.join(DATA_DIR, "users.json")
PARECERES_FILE = os.path.join(DATA_DIR, "pareceres.json")
SCHOOL_NAME = "ESCOLA MUNICIPAL DE EDUCAÇÃO FUNDAMENTAL ELESBÃO BARBOSA DE CARVALHO"
COORDENADOR_NAME = "NOME DO COORDENADOR AQUI"
STUDENT_NAMES = [
    "Andressa Viturino dos Santos", "Carlos Eduardo Conceição da Silva", "Damiana Honorato dos Santos",
    "Ivonete Nunes da Silva", "João Pedro da Silva", "José Alexandre Lemo da Silva",
    "Joselma Conceição da Silva", "José Gonzaga de Melo", "Maria Angelica Honorato de Oliveira",
    "Maria das Dores Rodrigues Pereira", "Nelson Pereira da Silva", "Nivaldilson Vitorino da Silva",
    "Roberlange Pereira da Silva", "Rosania Valério da Silva", "Rosileide Ferreira da Silva",
    "Thayse Ferreira dos Santos", "Alessandra Lopes da Silva", "Adenilson Lourenço dos Santos",
    "Antônio Ronaldo Firmino", "Adriano Viturino dos Santos", "Aristeu Vitorino dos Santos",
    "Aurene de Melo Gonzaga", "Bartolomeu Ferreira Vanderlei", "Adalva Gomes dos Santos",
    "Ana Maria da Silva Aquino"
]

# --- Funções Auxiliares ---
def sanitize_student_name_for_filename(name):
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

def get_parecer_template(template_type="normal"):
    # Caminho para o template padrão
    template_path = "parecer_template.docx"

    # Se a opção de não frequentar for selecionada, use um template diferente (opcional)
    # Por exemplo, "parecer_nao_frequentou.docx"
    # Você precisaria criar este arquivo e colocá-lo na mesma pasta do app.py
    # if template_type == "nao_frequentou":
    #     template_path = "parecer_nao_frequentou.docx"
    
    if not os.path.exists(template_path):
        st.error(f"Erro: O arquivo de modelo '{template_path}' não foi encontrado.")
        return None
    return Document(template_path)

def generate_parecer_docx(parecer_info):
    try:
        doc = get_parecer_template(template_type="normal")

        if not doc:
            return None

        substitutions = {
            'NOME_ESCOLA': SCHOOL_NAME,
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
    
    student_name = st.selectbox("Selecione o(a) aluno(a)", options=[""] + sorted(STUDENT_NAMES))
    filiacao_mae = st.text_input("Filiação (Mãe)")
    filiacao_pai = st.text_input("Filiação (Pai)")
    endereco = st.text_input("Endereço")
    naturalidade = st.text_input("Naturalidade")
    data_nascimento = st.date_input("Data de Nascimento")
    periodo = st.text_input("Período")
    turno = st.text_input("Turno")

    st.markdown("---")
    st.markdown("#### Conteúdo do Parecer")

    # Novo campo: Aluno que deixou de frequentar
    is_nao_frequentou = st.checkbox(
        "Aluno(a) deixou de frequentar?",
        help="Marque esta opção se o aluno não frequentou a escola."
    )

    parecer_texto_default = ""
    if is_nao_frequentou:
        parecer_texto_default = "Durante o período letivo, o(a) aluno(a) não frequentou a escola, não apresentando os critérios mínimos para avaliação."

    parecer_texto = st.text_area(
        "Digite o texto do parecer aqui:",
        value=parecer_texto_default,
        disabled=is_nao_frequentou,
        height=250
    )
    
    observacao = st.text_area("Observações (opcional):", height=100)

    if st.button("Gerar e Salvar Parecer"):
        if not student_name or (not parecer_texto and not is_nao_frequentou):
            st.error("Por favor, selecione um aluno e, se não for um aluno que não frequentou, preencha o texto do parecer.")
        else:
            if student_name not in pareceres:
                pareceres[student_name] = []

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
            docx_bytes = generate_parecer_docx(parecer_data)
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

    student_list = sorted(pareceres.keys())
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
                    
                    if 'nao_frequentou' in parecer_info and parecer_info['nao_frequentou']:
                        st.info("Este parecer é para um aluno que deixou de frequentar.")
                    else:
                        st.write(f"**Parecer:** {parecer_info['parecer_texto']}")
                        if 'observacao' in parecer_info and parecer_info['observacao']:
                            st.write(f"**Observações:** {parecer_info['observacao']}")

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
