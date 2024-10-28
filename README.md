# sales-department-web-scrapping

**Download do executável:** [baixar versão 1.1](https://github.com/FelixGomes/sales_web_scrapping/releases/download/v1.1/sales_web_scrapping.exe).

Repositório de códigos utilizado para web scrapping e data extraction das planilhas com leads. Utilizado para área comercial, com objetivo de preencher uma planilha que possua o nome dos clientes (leads) com dados sobre **redes sociais, site oficial, telefone, email e nome dos sócios**. Os dados são extraídos da internet.

Redes Sociais extraídas: **Facebook**, **Instagram** e **LinkedIn**. 

Veja exemplo de uma planilha com leads antes de rodar o código:
![image](https://github.com/user-attachments/assets/f47739d2-6e02-4767-a692-7fb719adc82d)

Após rodar o código:
![image](https://github.com/user-attachments/assets/fb9511b1-b1d6-4f09-aecc-fc8a4f9694be)

## Funcionalidades:
### Extrair dados de redes sociais
- Extrai os dados de Facebook, Instagram, LinkedIn e Site Oficial da empresa
- A extração ocorre utilizando o nome da empresa para buscar os dados no google
![image](https://github.com/user-attachments/assets/28f9f054-e267-426f-bf2e-61452fe4ad2d)

### Extrair dados do informe cadastral
- Extrai os dados de telefone, email e sócios da empresa
- A extração ocorre utilizando o CNPJ informado na planilha
- Os dados são extraídos apenas do site Informe Cadastral, que utiliza dados públicos e divulgados pelas empresas
![image](https://github.com/user-attachments/assets/609d5ee0-ec92-4dbc-88d6-6bd7451317a8)

### Escreve os dados na planilha automáticamente
- Todas as funcionalidades disponíveis, ao final da execução, irão escrever as informações na planilha inserida pelo usuário

### Observações
- Certifique-se de colocar os dados de **CNPJ** e **Nome da empresa** na planilha para que a aplicação funcione
![image](https://github.com/user-attachments/assets/6ba0373a-f22b-4119-af16-30e73c34e2f0)
- Para rodar a aplicação, é necessário ter o navegador Google Chrome instalado na sua máquina
- A aplicação funciona em segundo plano, portanto ela irá abrir e fechar páginas do navegador Google Chrome
- Para que os dados sejam salvos na planilha, a aplicação deve rodar até o final






