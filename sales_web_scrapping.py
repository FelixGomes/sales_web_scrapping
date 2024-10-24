import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup
import json

class MainApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Ferramenta de Extração de Dados")
        self.geometry("600x400")
        self.create_widgets()

    def create_widgets(self):
        # Label principal
        self.label = tk.Label(self, text="Escolha uma funcionalidade:", font=("Arial", 14))
        self.label.pack(pady=10)

        # Combobox para escolher a funcionalidade
        self.function_var = tk.StringVar()
        self.function_combobox = ttk.Combobox(self, textvariable=self.function_var, font=("Arial", 12), width=50, state="readonly")
        self.function_combobox['values'] = ('Extrair dados de redes sociais', 'Extrair dados do InformeCadastral')
        self.function_combobox.current(0)
        self.function_combobox.pack(pady=10)

        # Botão para selecionar o arquivo
        self.file_path_var = tk.StringVar()
        self.file_label = tk.Label(self, text="Caminho da planilha:", font=("Arial", 12))
        self.file_label.pack(pady=10)
        self.file_entry = tk.Entry(self, textvariable=self.file_path_var, font=("Arial", 12), width=50)
        self.file_entry.pack(pady=10)
        self.browse_button = tk.Button(self, text="Selecionar arquivo", command=self.browse_file)
        self.browse_button.pack(pady=10)

        # Botão para executar a extração
        self.run_button = tk.Button(self, text="Executar", command=self.run_extraction, font=("Arial", 14), width=20, bg="lightgray")
        self.run_button.pack(pady=20)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if file_path:
            self.file_path_var.set(file_path)

    def run_extraction(self):
        # Verifica se o arquivo foi selecionado
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("Erro", "Por favor, selecione um arquivo.")
            return

        # Le o arquivo Excel
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
            return

        # Verifica qual funcionalidade foi selecionada
        selected_function = self.function_var.get()
        if selected_function == 'Extrair dados de redes sociais':
            self.extrair_dados_redes_sociais(df)
        elif selected_function == 'Extrair dados do InformeCadastral':
            self.extrair_informe_cadastral(df)

    def extrair_dados_redes_sociais(self, df):
        if 'Nome da empresa' not in df.columns:
            messagebox.showerror("Erro", "A coluna 'Nome da empresa' não foi encontrada no arquivo.")
            return

        empresas = df['Nome da empresa'].dropna().tolist()
        messagebox.showinfo("Sucesso", f"Extração de redes sociais iniciada para {len(empresas)} empresas.")
        # Executa a função de extração de redes sociais
        extrair_dados_redes_sociais(empresas, "empresas_redes_sociais.csv")

    def extrair_informe_cadastral(self, df):
        if 'CNPJ' not in df.columns:
            messagebox.showerror("Erro", "A coluna 'CNPJ' não foi encontrada no arquivo.")
            return

        cnpjs = df['CNPJ'].dropna().tolist()
        messagebox.showinfo("Sucesso", f"Extração de dados do InformeCadastral iniciada para {len(cnpjs)} CNPJs.")
        # Executa a função de extração do InformeCadastral
        extrair_informe_cadastral(cnpjs, "dados_empresas_com_socios.csv")
# Função para calcular a pontuação dos links
def calcular_pontuacao_link(href, preview_text, palavras_empresa):
    pontuacao = 0
    for palavra in palavras_empresa:
        if palavra.lower() in href.lower():
            pontuacao += 1
        if palavra.lower() in preview_text.lower():
            pontuacao += 1
    return pontuacao
    
# Funções externas de extração:
def initialize_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    return webdriver.Chrome(options=chrome_options)

def salvar_progresso(dados_empresa, output_file="empresas_redes_sociais.csv"):
    try:
        df_resultados = pd.DataFrame(dados_empresa)
        df_resultados.to_csv(output_file, index=False)
        print(f"Progresso salvo no arquivo '{output_file}'.")
    except Exception as e:
        print(f"Erro ao salvar o progresso: {e}")

def extrair_informe_cadastral(cnpjs, output_file='dados_empresas_com_socios.csv'):
    driver = initialize_driver()
    url = "https://www.informecadastral.com.br/"
    dados_empresa = []

    try:
        for cnpj in cnpjs:
            try:
                driver.get(url)
                search_field = driver.find_element(By.CLASS_NAME, "form-control.border-radius-right-0.border-0")
                search_field.send_keys(cnpj)
                search_field.send_keys(Keys.RETURN)
                time.sleep(5)
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, 'html.parser')
                script_tag = soup.find('script', {'type': 'application/ld+json'})

                if script_tag:
                    data = json.loads(script_tag.string)
                    telefone = data.get("telephone", None)
                    email = data.get("contactPoint", {}).get("email", None)
                    socios_divs = soup.find_all('div', class_='col-md-6')
                    socios = [div.get_text(strip=True) for div in socios_divs]
                    socios_str = ", ".join(socios)
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
                salvar_progresso(dados_empresa, output_file)
                driver.quit()
                return

    except KeyboardInterrupt:
        print("\nInterrupção detectada. Salvando progresso...")
        salvar_progresso(dados_empresa, output_file)
        print("Progresso salvo. Saindo do programa.")
        return

    finally:
        driver.quit()
        salvar_progresso(dados_empresa, output_file)

def extrair_dados_redes_sociais(empresas, output_file='empresas_redes_sociais.csv'):
    driver = initialize_driver()
    dados_empresa = []

    try:
        for empresa in empresas:
            try:
                print(f"Iniciando extração para a empresa: {empresa}")
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

                dados_empresa.append({
                    "Nome da empresa": empresa,
                    "Facebook": facebook_link,
                    "Instagram": instagram_link,
                    "LinkedIn": linkedin_link,
                    "Site Oficial": site_oficial_link
                })
                print(f"Extração concluída para a empresa: {empresa}")

            except Exception as e:
                print(f"Erro ao processar a empresa {empresa}: {str(e)}")
                salvar_progresso(dados_empresa, output_file)
                driver.quit()
                return

    except KeyboardInterrupt:
        print("\nInterrupção detectada. Salvando progresso...")
        salvar_progresso(dados_empresa, output_file)
        print("Progresso salvo. Saindo do programa.")
        return

    finally:
        driver.quit()
        salvar_progresso(dados_empresa, output_file)

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
