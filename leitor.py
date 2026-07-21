import streamlit as st
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_bytes
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import io
import datetime
from docx import Document

# Configuração do Tesseract
## pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Dicionário de conversão de Estados para Siglas
ESTADOS_BR = {
    "ACRE": "AC", "ALAGOAS": "AL", "AMAPÁ": "AP", "AMAZONAS": "AM", "BAHIA": "BA",
    "CEARÁ": "CE", "DISTRITO FEDERAL": "DF", "ESPÍRITO SANTO": "ES", "GOIÁS": "GO",
    "MARANHÃO": "MA", "MATO GROSSO": "MT", "MATO GROSSO DO SUL": "MS", "MINAS GERAIS": "MG",
    "PARÁ": "PA", "PARAÍBA": "PB", "PARANÁ": "PR", "PERNAMBUCO": "PE", "PIAUÍ": "PI",
    "RIO DE JANEIRO": "RJ", "RIO GRANDE DO NORTE": "RN", "RIO GRANDE DO SUL": "RS",
    "RONDÔNIA": "RO", "RORAIMA": "RR", "SANTA CATARINA": "SC", "SÃO PAULO": "SP",
    "SERGIPE": "SE", "TOCANTINS": "TO"
}

# ==========================================
# Interface do App (Streamlit) - BARRA LATERAL INTELIGENTE
# ==========================================
st.set_page_config(page_title="Gerador de Iniciais", layout="wide")

st.sidebar.title("🔐 Configurações")

# Tenta buscar do cofre de senhas (secrets.toml). Se não achar, deixa vazio ("").
usuario_salvo = st.secrets.get("COBRARE_USER", "")
senha_salva = st.secrets.get("COBRARE_PASSWORD", "")

# Os campos aparecem na tela, mas já vêm preenchidos se a senha estiver no cofre!
usuario_cobrare = st.sidebar.text_input("Usuário do COBRARE", value=usuario_salvo)
senha_cobrare = st.sidebar.text_input("Senha do COBRARE", type="password", value=senha_salva)

# ==========================================
# Módulo 1: Extração e Mineração de Confissão
# ==========================================
def extrair_texto_hibrido(arquivo_bytes):
    texto_completo = ""
    try:
        doc = fitz.open(stream=arquivo_bytes, filetype="pdf")
        for num_pagina in range(len(doc)):
            pagina = doc.load_page(num_pagina)
            texto_nativo = pagina.get_text()
            
            if len(texto_nativo.strip()) < 50:
                imagens = convert_from_bytes(arquivo_bytes, first_page=num_pagina+1, last_page=num_pagina+1)
                for imagem in imagens:
                    texto_ocr = pytesseract.image_to_string(imagem, lang='por')
                    texto_completo += texto_ocr + "\n"
            else:
                texto_completo += texto_nativo + "\n"
        doc.close()
    except Exception as e:
        return f"Erro ao processar o arquivo: {e}"
    return texto_completo

def minerar_dados_confissao(texto_bruto):
    dados = {
        "credor": "Não encontrado",
        "devedor": "Não encontrado",
        "cpf_devedor": None,
        "cidade_comarca": "Não encontrada"
    }
    texto_limpo = texto_bruto.replace('\n', ' ').replace('  ', ' ')
    
    padrao_credor = r"de um lado,\s*(.*?),\s*(?:representada neste ato|doravante denominada 1ª)"
    busca_credor = re.search(padrao_credor, texto_limpo, re.IGNORECASE)
    if busca_credor: dados["credor"] = busca_credor.group(1).strip()
        
    padrao_devedor = r"de outro lado,\s*(.*?),\s*doravante denominado"
    busca_devedor = re.search(padrao_devedor, texto_limpo, re.IGNORECASE)
    if busca_devedor:
        dados["devedor"] = busca_devedor.group(1).strip()
        padrao_cpf = r"CPF/MF sob o nº\s*([\d\.\-]+)"
        busca_cpf = re.search(padrao_cpf, dados["devedor"])
        if busca_cpf: dados["cpf_devedor"] = busca_cpf.group(1).strip()
        
        padrao_cidade = r"([A-Za-zÀ-ÿ\s]+)\s*-\s*[A-Z]{2}(?:\s*,|$)"
        busca_cidade = re.search(padrao_cidade, dados["devedor"])
        if busca_cidade: dados["cidade_comarca"] = busca_cidade.group(1).strip().upper()
            
    return dados

# ==========================================
# Módulo 2: Automação Web (COBRARE)
# ==========================================
# ==========================================
# Módulo 2: Automação Web (COBRARE)
# ==========================================
def buscar_endereco_cobrare(pesquisa_devedor):
    chrome_options = Options()
    
    # --- NOVAS CONFIGURAÇÕES PARA O SERVIDOR NA NUVEM ---
    chrome_options.add_argument("--headless=new") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")

    # ADICIONE ESTAS DUAS LINHAS:
    chrome_options.add_argument("--window-size=1920,1080") # Força uma tela Full HD
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36") # Finge ser um PC normal
    # ----------------------------------------------------
    # ----------------------------------------------------
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 15)
    
    try:
        driver.get("https://cobrare.atenta.pro/controleDivida/menu")
        
        wait.until(EC.presence_of_element_located((By.ID, "login"))).send_keys(usuario_cobrare)
        driver.find_element(By.ID, "password").send_keys(senha_cobrare)
        driver.find_element(By.ID, "password").submit()
        
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-toggle='dropdown' and contains(., 'Movimentação')]"))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/controleDivida/negociacao')]"))).click()
        
        campo_busca = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@aria-controls='negociacao']")))
        
        # 1. Limpa o campo por garantia
        campo_busca.clear()
        
        # 2. Digita o CPF
        cpf_limpo = pesquisa_devedor.replace(".", "").replace("-", "")
        campo_busca.send_keys(cpf_limpo)
        
        # 3. Força um "Enter" no teclado
        campo_busca.send_keys(Keys.ENTER)
        
        # 4. Aumenta a pausa para o sistema ter tempo de filtrar a tabela e ocultar os devedores antigos
        time.sleep(4)
        
        botao_editar = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/edit') and contains(@class, 'btn-mini')]")))
        botao_editar.click()
        
        aba_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@id, 'ui-id-') and contains(@class, 'ui-tabs-anchor')]")))
        aba_info.click()
        
        wait.until(EC.presence_of_element_located((By.ID, "txEndereco")))
        
        cidade_select = Select(driver.find_element(By.ID, "cidadeCombo"))
        estado_select = Select(driver.find_element(By.ID, "estadoCombo"))
        
        cidade_texto = cidade_select.first_selected_option.get_attribute("textContent").strip()
        estado_completo = estado_select.first_selected_option.get_attribute("textContent").strip().upper()
        
        estado_sigla = ESTADOS_BR.get(estado_completo, estado_completo) 
        
        dados_sistema = {
            "latitude": driver.find_element(By.ID, "nrLatitude").get_attribute("value"),
            "longitude": driver.find_element(By.ID, "nrLongitude").get_attribute("value"),
            "logradouro": driver.find_element(By.ID, "txEndereco").get_attribute("value"),
            "numero": driver.find_element(By.ID, "nrEndereco").get_attribute("value"),
            "complemento": driver.find_element(By.ID, "dsComplemento").get_attribute("value"),
            "bairro": driver.find_element(By.ID, "nmBairro").get_attribute("value"),
            "cep": driver.find_element(By.ID, "nrCep").get_attribute("value"),
            "cidade": cidade_texto,
            "estado": estado_sigla
        }
        return dados_sistema
    except Exception as e:
        return f"Erro na automação: {e}"
    finally:
        driver.quit()

def atualizar_endereco_devedor(texto_devedor, endereco_cobrare, cidade_comarca):
    if not isinstance(endereco_cobrare, dict):
        return texto_devedor
        
    comp = f" - {endereco_cobrare['complemento']}" if endereco_cobrare['complemento'] else ""
    cidade_estado = f"{endereco_cobrare['cidade']} - {endereco_cobrare['estado']}"
    
    lat = endereco_cobrare.get('latitude', '').strip()
    lon = endereco_cobrare.get('longitude', '').strip()
    coords = f", Latitude {lat}, Longitude {lon}" if lat and lon else ""
    
    novo_endereco = f"{endereco_cobrare['logradouro']}, nº {endereco_cobrare['numero']}{comp}, Bairro {endereco_cobrare['bairro']}, CEP {endereco_cobrare['cep']}{coords}, {cidade_estado}"
    
    padrao = r"(residente\(s\)\s+e\s+domiciliado\(s\)\s+em\s+)(.*)"
    texto_atualizado = re.sub(padrao, rf"\g<1>{novo_endereco}", texto_devedor, flags=re.IGNORECASE)
    
    return texto_atualizado

# ==========================================
# Módulo 3: Mineração Detran e Formatação
# ==========================================
def minerar_dados_detran(texto_bruto):
    dados = {
        "marca": "Não encontrada", "placa": "Não encontrada",
        "ano_modelo": "Não encontrado", "cor": "Não encontrada",
        "renavam": "Não encontrado", "chassi": "Não encontrado"
    }
    busca_placa = re.search(r"Placa:\s*([A-Z0-9]{7})", texto_bruto, re.IGNORECASE)
    if busca_placa: dados["placa"] = busca_placa.group(1).strip()
    
    busca_chassi = re.search(r"Chassi:\s*([A-Z0-9]+)", texto_bruto, re.IGNORECASE)
    if busca_chassi: dados["chassi"] = busca_chassi.group(1).strip()
    
    busca_marca = re.search(r"Marca:\s*(.+)", texto_bruto, re.IGNORECASE)
    if busca_marca: dados["marca"] = busca_marca.group(1).strip()
        
    busca_renavam = re.search(r"RENAVAM:\s*([0-9]+)", texto_bruto, re.IGNORECASE)
    if busca_renavam: dados["renavam"] = busca_renavam.group(1).strip()
        
    busca_ano = re.search(r"Fabricação/Modelo:\s*([\d]{4}\s*/\s*[\d]{4})", texto_bruto, re.IGNORECASE)
    if busca_ano: dados["ano_modelo"] = busca_ano.group(1).replace(" ", "")
        
    busca_cor = re.search(r"Cor:\s*([A-Za-zÀ-ÿ]+)", texto_bruto, re.IGNORECASE)
    if busca_cor: dados["cor"] = busca_cor.group(1).strip().lower()
    return dados

# ==========================================
# Módulo 4: Preenchimento Avançado do Word
# ==========================================
def gerar_documento_word(caminho_modelo, comarca, credor, devedor_qualificado, lista_veiculos):
    doc = Document(caminho_modelo)
    qualificacao_count = 0
    
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    hoje = datetime.date.today()
    data_formatada = f"{hoje.day} de {meses[hoje.month - 1]} de {hoje.year}"
    
    for p in doc.paragraphs:
        if '[COMARCA]' in p.text:
            p.text = p.text.replace('[COMARCA]', comarca)
            
        if '[QUALIFICAÇÃO COMPLETA]' in p.text:
            if qualificacao_count == 0:
                p.text = p.text.replace('[QUALIFICAÇÃO COMPLETA]', credor)
                qualificacao_count += 1
            else:
                p.text = p.text.replace('[QUALIFICAÇÃO COMPLETA]', devedor_qualificado)
                
        if 'Santa Cruz do Sul/RS,' in p.text:
            p.text = re.sub(r'Santa Cruz do Sul/RS,.*', f'Santa Cruz do Sul/RS, {data_formatada}.', p.text)
                
        if 'd) A expedição de' in p.text and lista_veiculos:
            p.clear() 
            p.add_run("d) Para a efetivação da penhora, a Exequente indica os seguintes veículos de propriedade do Executado: ")
            
            romanos = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]
            for i, v in enumerate(lista_veiculos):
                idx_romano = romanos[i] if i < len(romanos) else str(i + 1)
                
                run_bold = p.add_run(f"{idx_romano} - {v['marca']}")
                run_bold.bold = True
                
                separator = "; " if i < len(lista_veiculos) - 1 else "."
                p.add_run(f", de placa {v['placa']}, ano/modelo {v['ano_modelo']}, cor {v['cor']}, possui o RENAVAM {v['renavam']} e o chassi {v['chassi']}{separator} ")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# Execução Principal
# ==========================================
st.title("📄 Gerador de Iniciais Automatizado")
st.write("Faça o upload da Confissão de Dívida e das Certidões do Detran abaixo:")

confissao_file = st.file_uploader("Anexar Confissão de Dívida (PDF)", type=["pdf"])
detran_files = st.file_uploader("Anexar Certidão(ões) do Detran (PDF)", type=["pdf"], accept_multiple_files=True)

if st.button("Processar e Gerar Inicial"):
    if confissao_file and detran_files:
        st.success("Arquivos recebidos! Iniciando montagem da inicial...")
        
        with st.spinner("Extraindo e minerando Confissão de Dívida..."):
            texto_confissao = extrair_texto_hibrido(confissao_file.getvalue())
            dados_minerados = minerar_dados_confissao(texto_confissao)
            qualificacao_devedor = dados_minerados["devedor"] 
            
        if dados_minerados['cpf_devedor']:
            with st.spinner("Buscando endereço no COBRARE..."):
                endereco_cobrare = buscar_endereco_cobrare(dados_minerados['cpf_devedor'])
                
                # Se for texto (erro) e não dicionário, mostramos na tela!
                if isinstance(endereco_cobrare, str):
                    st.error(f"Detalhe do erro: {endereco_cobrare}")
                    
                qualificacao_devedor = atualizar_endereco_devedor(
                    qualificacao_devedor, 
                    endereco_cobrare, 
                    dados_minerados['cidade_comarca']
                )

        lista_veiculos = []
        for i, certidao in enumerate(detran_files):
            with st.spinner(f"Processando Certidão {i+1} ({certidao.name})..."):
                texto_detran = extrair_texto_hibrido(certidao.getvalue())
                lista_veiculos.append(minerar_dados_detran(texto_detran))
                
        with st.spinner("Montando o documento final e aplicando formatação..."):
            caminho_modelo_word = "1. Modelo inicial execução (confissão de dívida).docx"
            
            doc_final_bytes = gerar_documento_word(
                caminho_modelo_word,
                dados_minerados['cidade_comarca'],
                dados_minerados['credor'],
                qualificacao_devedor,
                lista_veiculos 
            )
            
            st.success("✅ Petição Inicial gerada com sucesso!")
            
            st.download_button(
                label="⬇️ Baixar Inicial Pronta (.docx)",
                data=doc_final_bytes,
                file_name=f"Inicial_Execucao_{dados_minerados['cpf_devedor']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
    else:
        st.warning("⚠️ Por favor, insira a Confissão de Dívida e pelo menos uma Certidão do Detran antes de continuar.")