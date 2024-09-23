from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
from bs4 import BeautifulSoup
import os
import json

# Função para calcular a pontuação dos links
def calcular_pontuacao_link(href, preview_text, palavras_empresa):
    pontuacao = 0
    for palavra in palavras_empresa:
        if palavra.lower() in href.lower():
            pontuacao += 1
        if palavra.lower() in preview_text.lower():
            pontuacao += 1
    return pontuacao

# Função para salvar o progresso em um arquivo CSV
def salvar_progresso(dados_empresa, output_file="empresas_redes_sociais.csv"):
    try:
        # Verifica se o diretório onde o arquivo será salvo existe, caso contrário, cria
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        df_resultados = pd.DataFrame(dados_empresa)
        df_resultados.to_csv(output_file, index=False)
        print(f"Progresso salvo no arquivo '{output_file}'.")
    except Exception as e:
        print(f"Erro ao salvar o progresso: {e}")

# Função para inicializar o WebDriver com SSL options
def initialize_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    return webdriver.Chrome(options=chrome_options)

# Função para extrair dados do InformeCadastral
def extrair_informe_cadastral(cnpjs, output_file='dados_empresas_com_socios.csv'):
    # Inicializa o WebDriver
    driver = initialize_driver()

    # URL do site de pesquisa
    url = "https://www.informecadastral.com.br/"

    # Lista para armazenar os dados extraídos
    dados_empresa = []

    try:
        # Loop através dos CNPJs
        for cnpj in cnpjs:
            try:
                # Acessa o site
                driver.get(url)
                
                # Localiza o campo de pesquisa por sua classe e insere o CNPJ
                search_field = driver.find_element(By.CLASS_NAME, "form-control.border-radius-right-0.border-0")
                search_field.send_keys(cnpj)
                search_field.send_keys(Keys.RETURN)
                
                # Aguarda a página carregar
                time.sleep(5)  # Ajuste o tempo de espera conforme necessário

                # Obtém o conteúdo HTML da página
                page_source = driver.page_source

                # Utiliza BeautifulSoup para parsear o HTML
                soup = BeautifulSoup(page_source, 'html.parser')
                
                # Encontra o script que contém o JSON com os dados da empresa
                script_tag = soup.find('script', {'type': 'application/ld+json'})
                
                # Verifica se o script com JSON foi encontrado
                if script_tag:
                    # Carrega o conteúdo JSON como um dicionário
                    data = json.loads(script_tag.string)
                    
                    # Extrai os dados desejados (telefone e e-mail)
                    telefone = data.get("telephone", None)
                    email = data.get("contactPoint", {}).get("email", None)
                    
                    # Extrai os nomes dos sócios da div com a classe 'col-md-6'
                    socios_divs = soup.find_all('div', class_='col-md-6')
                    socios = [div.get_text(strip=True) for div in socios_divs]
                    socios_str = ", ".join(socios)
                    
                    # Adiciona os dados à lista
                    dados_empresa.append({
                        "CNPJ": cnpj,
                        "Telefone": telefone,
                        "Email": email,
                        "Sócios": socios_str
                    })
                else:
                    print(f"Dados não encontrados para o CNPJ {cnpj}.")
            
            except Exception as e:
                print(f"Erro ao processar o CNPJ {cnpj}: {str(e)}")
                salvar_progresso(dados_empresa, output_file)  # Salva o progresso até o ponto da falha
                driver.quit()
                return  # Interrompe a execução para evitar mais erros

    except KeyboardInterrupt:
        print("\nInterrupção detectada. Salvando progresso...")
        salvar_progresso(dados_empresa, output_file)
        print("Progresso salvo. Saindo do programa.")
        return

    finally:
        # Fecha o navegador
        driver.quit()
        # Salvar o progresso final
        salvar_progresso(dados_empresa, output_file)

# Função para extrair dados de redes sociais
def extrair_dados_redes_sociais(empresas, output_file):
    # Inicializa o WebDriver
    driver = initialize_driver()
    dados_empresa = []

    try:
        # Loop através das empresas
        for empresa in empresas:
            try:
                print(f"Iniciando extração para a empresa: {empresa}")

                # Inicializa as variáveis de links e telefone
                facebook_link, instagram_link, linkedin_link, site_oficial_link = None, None, None, None

                # Pesquisa no Google para Facebook
                driver.get("https://www.google.com")
                search_field = driver.find_element(By.NAME, "q")
                search_field.send_keys(f"{empresa} facebook")
                search_field.send_keys(Keys.RETURN)
                time.sleep(10)  # Ajuste o tempo de espera conforme necessário

                # Extração de link do Facebook
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, 'html.parser')
                for link in soup.find_all('a', href=True):
                    href = link['href']
                    if 'facebook.com' in href:
                        facebook_link = href
                        print('Facebook: OK')
                        break

                # Pesquisa no Google para Instagram
                driver.get("https://www.google.com")
                search_field = driver.find_element(By.NAME, "q")
                search_field.send_keys(f"{empresa} instagram")
                search_field.send_keys(Keys.RETURN)
                time.sleep(10)

                # Extração de link do Instagram
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, 'html.parser')
                for link in soup.find_all('a', href=True):
                    href = link['href']
                    if 'instagram.com' in href:
                        instagram_link = href
                        print('Instagram: OK')
                        break

                # Pesquisa no Google para LinkedIn
                driver.get("https://www.google.com")
                search_field = driver.find_element(By.NAME, "q")
                search_field.send_keys(f"{empresa} linkedin")
                search_field.send_keys(Keys.RETURN)
                time.sleep(10)

                # Extração de link do LinkedIn
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, 'html.parser')
                for link in soup.find_all('a', href=True):
                    href = link['href']
                    if 'linkedin.com' in href:
                        linkedin_link = href
                        print('LinkedIn: OK')
                        break

                # Pesquisa no Google para Site Oficial
                driver.get("https://www.google.com")
                search_field = driver.find_element(By.NAME, "q")
                search_field.send_keys(f"{empresa} site oficial")
                search_field.send_keys(Keys.RETURN)
                time.sleep(10)

                # Extração de link do Site Oficial
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, 'html.parser')

                # Dicionário para armazenar links e suas pontuações
                links_com_pontuacao = {}
                empresa_palavras = empresa.split()
                # Ignorar links que provavelmente não são o site oficial
                palavras_ignoradas = ["cnpj", "econodata", "lista de empresas", "guia", "dados", "consulta", "banco de dados", "google", "search", "emis", "solutudo",
                                      "boxcis", "gett", "fazcomex", "carreiraseprofissoes", "glassdoor", "estadao", "reclameaqui", "gupy", "instagram", "facebook", "find", "shell",
                                      "relatorioreservado", "linkedin", "som-automotivo", "jusbrasil", "folhavitoria", "folha", "globo", "chapeco.org", "detran", "diregional", 
                                      "leismunicipais", "infojobs", "enlizt", "jus.br", "trf4"]

                # Procurando por links que sejam externos e que contenham palavras da empresa
                for link in soup.find_all('a', href=True):
                    href = link['href']
                    # Obtendo o preview ou título da página
                    preview_text = link.get_text()  # Pega o texto visível do link
                    # Filtra apenas links externos e ignora sites irrelevantes
                    if href.startswith('http') and not any(ignorada in href.lower() for ignorada in palavras_ignoradas) and not any(ignorada in preview_text.lower() for ignorada in palavras_ignoradas):
                        # Calcula a pontuação do link baseado nas palavras da empresa e no preview da página
                        pontuacao = calcular_pontuacao_link(href, preview_text, empresa_palavras)
                        # Apenas consideramos links com pontuação maior que 0
                        if pontuacao > 0:
                            links_com_pontuacao[href] = pontuacao

                # Verifica se encontramos links e escolhe o melhor
                if links_com_pontuacao:
                    # Ordena os links pela maior pontuação
                    site_oficial_link = max(links_com_pontuacao, key=links_com_pontuacao.get)
                    print(f"Link mais provável de {empresa}: {site_oficial_link}")
                else:
                    print(f"Nenhum link encontrado relacionado à {empresa}.")

                # Adiciona os dados à lista
                dados_empresa.append({
                    "Nome da empresa": empresa,
                    "Facebook": facebook_link,
                    "Instagram": instagram_link,
                    "LinkedIn": linkedin_link,
                    "Site Oficial": site_oficial_link
                })

                # Exibe a conclusão da extração para a empresa
                print(f"Extração concluída para a empresa: {empresa}")

            except Exception as e:
                print(f"Erro ao processar a empresa {empresa}: {str(e)}")
                salvar_progresso(dados_empresa, output_file)
                driver.quit()
                return  # Interrompe a execução para evitar mais erros

    except KeyboardInterrupt:
        print("\nInterrupção detectada. Salvando progresso...")
        salvar_progresso(dados_empresa, output_file)
        print("Progresso salvo. Saindo do programa.")
        return

    finally:
        # Fecha o navegador
        driver.quit()
        # Salvar o progresso final
        salvar_progresso(dados_empresa, output_file)

# Função principal para iniciar o script
def main():
    escolha_funcionalidade = input("Escolha uma funcionalidade:\n1. Extrair dados de redes sociais\n2. Extrair dados do InformeCadastral\nDigite 1 ou 2: ")

    if escolha_funcionalidade == '1':
        # Extração de redes sociais
        escolha = input("A planilha deve conter cabeçalhos: 'Nome da empresa' e 'Informações do CNPJ'\nEscolha uma opção:\n1. Inserir o link da planilha do Google Sheets\n2. Inserir o caminho da planilha baixada\nDigite 1 ou 2: ")

        if escolha == '1':
            # Exemplo de link do Google Sheets
            # O link deve ser no formato: 'https://docs.google.com/spreadsheets/d/<ID_da_planilha>/export?format=xlsx'
            google_sheets_link = input("\nOBS.: A planilha deve estar compartilhada para qualquer pessoa com o link possa acessar\nExemplo de link: https://docs.google.com/spreadsheets/d/1UwL-kTQYpUeVrVUz0J8wUzWu6UGCUxvFcFcTHtGQgqs/export?format=xlsx\nInsira o link da planilha do Google Sheets: ")
            try:
                empresas_df = pd.read_excel(google_sheets_link)
            except Exception as e:
                print(f"Erro ao acessar a planilha do Google Sheets: {e}")
                return

        elif escolha == '2':
            # Exemplo de caminho do arquivo local:
            # 'C:\\Users\\gustavo\\Downloads\\Leads Econodata.xlsx'
            file_path = input("\nExemplo de caminho do arquivo local: C:\\Users\\gustavo\\Downloads\\Leads Econodata.xlsx\nInsira o caminho da planilha baixada: ")
            try:
                empresas_df = pd.read_excel(file_path)
            except Exception as e:
                print(f"Erro ao acessar a planilha baixada: {e}")
                return

        else:
            print("Opção inválida. Execute o script novamente e escolha 1 ou 2.")
            return

        # Verifica se a coluna 'Nome da empresa' existe e extrai os nomes
        if 'Nome da empresa' in empresas_df.columns:
            empresas = empresas_df['Nome da empresa'].dropna().tolist()  # Remove valores nulos e converte para lista
        else:
            raise ValueError("A coluna 'Nome da empresa' não foi encontrada no arquivo Excel.")

        # Pergunta se o usuário deseja carregar um progresso salvo
        dados_empresa = []
        carregar_progresso = input("Deseja carregar um progresso salvo? (s/n): ")
        if carregar_progresso.lower() == 's':
            output_file = input("Insira o caminho completo do arquivo CSV com o progresso salvo (exemplo: C:\\Users\\usuario\\Downloads\\empresas_redes_sociais.csv): ")
            if os.path.exists(output_file):
                df_salvo = pd.read_csv(output_file)
                empresas_extraidas = df_salvo['Nome da empresa'].tolist()
                # Continuar a partir do último cliente
                dados_empresa = df_salvo.to_dict('records')
                # Remove empresas já extraídas da lista
                empresas = [empresa for empresa in empresas if empresa not in empresas_extraidas]
            else:
                print(f"Arquivo {output_file} não encontrado. Iniciando novo processo.")
        else:
            # Define um nome padrão para o arquivo de progresso
            output_file = 'empresas_redes_sociais.csv'

        # Executa a extração das redes sociais
        extrair_dados_redes_sociais(empresas, output_file)

    elif escolha_funcionalidade == '2':
        # Extração do InformeCadastral
        file_path = input("\nExemplo de caminho do arquivo local: C:\\Users\\gustavo\\Downloads\\Leads Econodata.xlsx\nInsira o caminho da planilha baixada contendo os CNPJs: ")
        try:
            empresas_df = pd.read_excel(file_path)
            if 'CNPJ' not in empresas_df.columns:
                raise ValueError("A coluna 'CNPJ' não foi encontrada no arquivo Excel.")
            cnpjs = empresas_df['CNPJ'].dropna().tolist()  # Extrai a lista de CNPJs
        except Exception as e:
            print(f"Erro ao acessar a planilha ou extrair CNPJs: {e}")
            return
        
        # Define um nome padrão para o arquivo de progresso
        output_file = 'dados_empresas_com_socios.csv'
        
        # Executa a extração do InformeCadastral
        extrair_informe_cadastral(cnpjs, output_file)

    else:
        print("Opção inválida. Execute o script novamente e escolha 1 ou 2.")

# Executar o script
main()
