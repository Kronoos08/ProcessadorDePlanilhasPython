# Processador-de-planilhas-Python
Um script Python robusto para cruzar, organizar e padronizar dados de múltiplas planilhas Excel, com lógica de prioridade para informações de responsáveis e formatação automática de campos.
Conteúdo Sugerido para o Arquivo README.md:
# Processador Genérico de Planilhas Excel

Este projeto contém um script Python (`process_data_files.py`) projetado para automatizar o cruzamento e a padronização de dados entre duas planilhas Excel distintas. Ele é ideal para cenários onde informações de diferentes fontes precisam ser unificadas e formatadas de maneira consistente em um único arquivo de saída.

## Funcionalidades Principais

* **Cruzamento de Dados:** Mescla dados de duas planilhas (`input_file_1.xlsx` e `input_file_2.xlsx`) com base em um identificador comum (`ID_NUMBER`).
* **Lógica de Prioridade de Responsáveis:** Preenche as informações do `RESPONSIBLE1` (Responsável 1) com dados do `PARENT2` (Responsável 2) caso as informações do `PARENT1` (Responsável 1 original) estejam ausentes.
* **Formatação Padronizada:**
    * Converte sobrenomes (`LAST_NAME`, `RESPONSIBLE1_LAST_NAME`, `RESPONSIBLE2_LAST_NAME`) para letras maiúsculas.
    * Formata nomes (`FIRST_NAME`, `RESPONSIBLE1_FIRST_NAME`, `RESPONSIBLE2_FIRST_NAME`) e campos de endereço (`ADDRESS_LINE_1`, `ADDRESS_LINE_2`, `ADDRESS_LINE_3`, `CITY`, `BIRTH_PLACE`, `BIRTH_COUNTRY`, `GENDER`, `TITLE`) para "Title Case" (primeira letra de cada palavra em maiúscula).
    * Converte endereços de e-mail (`RESPONSIBLE1_EMAIL`, `RESPONSIBLE2_EMAIL`) para letras minúsculas.
    * Padroniza números de identificação (`ID_NUMBER`) com preenchimento de zeros à esquerda.
    * Formata datas de nascimento (`BIRTH_DATE`) para `DD/MM/AAAA`.
* **Tratamento de Dados Ausentes:** Lida com valores `NaN` (Not a Number) e strings vazias para garantir um processamento robusto.
* **Mensagens de Erro Detalhadas:** Fornece feedback claro em caso de arquivos não encontrados ou colunas ausentes.

## Como Usar

### Pré-requisitos

Certifique-se de ter o Python instalado em seu sistema. Este script utiliza as bibliotecas `pandas` e `openpyxl`. Se você ainda não as tem, pode instalá-las via pip:

```bash
pip install pandas openpyxl

Estrutura dos Arquivos de Entrada
O script espera dois arquivos Excel de entrada com as seguintes estruturas de coluna (os nomes das colunas são case-sensitive):
input_file_1.xlsx (Exemplo: Dados de um sistema primário)
•	ID_NUMBER (Número de identificação único)
•	BIRTH_DATE (Data de nascimento)
•	GENDER (Gênero)
•	BIRTH_PLACE (Local de nascimento)
•	BIRTH_COUNTRY (País de nascimento)
input_file_2.xlsx (Exemplo: Dados de um sistema secundário)
•	IDENTIFIER (Identificador, será renomeado para ID_NUMBER para o merge)
•	PERSON_NAME (Nome completo da pessoa, ex: "SOBRENOME, NOME")
•	PERSON_ADDRESS (Endereço completo da pessoa, ex: "Rua X, 123 - Bairro - Complemento - Cidade - CEP")
•	PARENT1_NAME (Nome do primeiro responsável)
•	PARENT1_ADDRESS (Endereço do primeiro responsável)
•	PARENT1_EMAIL (E-mail do primeiro responsável)
•	PARENT1_PHONE (Telefone do primeiro responsável)
•	PARENT2_NAME (Nome do segundo responsável)
•	PARENT2_ADDRESS (Endereço do segundo responsável)
•	PARENT2_EMAIL (E-mail do segundo responsável)
•	PARENT2_PHONE (Telefone do segundo responsável)
Executando o Script
1.	Coloque os arquivos de entrada: Certifique-se de que input_file_1.xlsx e input_file_2.xlsx estejam no mesmo diretório que o script process_data_files.py, ou ajuste os caminhos no script.
Execute o script Python:
1.	python process_data_files.py
2.	
3.	Verifique a saída: Uma nova planilha chamada output_file.xlsx será gerada no mesmo diretório, contendo os dados cruzados e organizados.
Estrutura da Planilha de Saída (output_file.xlsx)
A planilha de saída terá as seguintes colunas na ordem especificada:
•	RECORD_NUMBER
•	TITLE
•	LAST_NAME
•	FIRST_NAME
•	BIRTH_DATE
•	GENDER
•	BIRTH_PLACE
•	BIRTH_COUNTRY
•	ID_NUMBER
•	ADDRESS_LINE_1
•	ADDRESS_LINE_2
•	ADDRESS_LINE_3
•	POSTAL_CODE
•	CITY
•	RESPONSIBLE1_ROLE (Sempre "LEGAL")
•	RESPONSIBLE1_LAST_NAME
•	RESPONSIBLE1_FIRST_NAME
•	RESPONSIBLE1_ADDRESS_LINE_1
•	RESPONSIBLE1_ADDRESS_LINE_2
•	RESPONSIBLE1_ADDRESS_LINE_3
•	RESPONSIBLE1_POSTAL_CODE
•	RESPONSIBLE1_CITY
•	RESPONSIBLE1_EMAIL
•	RESPONSIBLE1_PHONE
•	RESPONSIBLE2_ROLE (Será "LEGAL" se houver um segundo responsável, caso contrário, vazio)
•	RESPONSIBLE2_LAST_NAME
•	RESPONSIBLE2_FIRST_NAME
•	RESPONSIBLE2_ADDRESS_LINE_1
•	RESPONSIBLE2_ADDRESS_LINE_2
•	RESPONSIBLE2_ADDRESS_LINE_3
•	RESPONSIBLE2_POSTAL_CODE
•	RESPONSIBLE2_CITY
•	RESPONSIBLE2_EMAIL
•	RESPONSIBLE2_PHONE
Exemplo de Uso (com arquivos de exemplo)
Para facilitar o teste, você pode usar o script generate_sample_excel_files.py (fornecido separadamente) para criar os arquivos input_file_1.xlsx e input_file_2.xlsx com dados fictícios.
Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests para melhorias, correções de bugs ou novas funcionalidades.
Licença
Este projeto está licenciado sob a Licença MIT. Veja
