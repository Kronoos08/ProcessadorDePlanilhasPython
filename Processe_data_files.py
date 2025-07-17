import pandas as pd
import re

def process_data_files(input_file_path_1, input_file_path_2, output_file_path):
    """
    Cruza dados de duas planilhas de entrada, organiza e salva em uma terceira planilha de saída.
    Prioriza o preenchimento das colunas do 'Responsável 1' com dados do 'Responsável 2'
    se os dados do 'Responsável 1' estiverem ausentes.
    Preenche a coluna 'Responsible2_Role' com o valor "LEGAL" APENAS se houver dados para um segundo responsável.
    Converte as colunas de sobrenome ('LAST_NAME', 'RESPONSIBLE1_LAST_NAME', 'RESPONSIBLE2_LAST_NAME')
    para letras maiúsculas.
    Formata as colunas de endereço (endereço e cidade) com a primeira letra de CADA PALAVRA maiúscula (Title Case).
    Converte as colunas de e-mail ('RESPONSIBLE1_EMAIL', 'RESPONSIBLE2_EMAIL') para letras minúsculas.

    Args:
        input_file_path_1 (str): Caminho para o primeiro arquivo de planilha de entrada (ex: de um sistema).
        input_file_path_2 (str): Caminho para o segundo arquivo de planilha de entrada (ex: de outro sistema).
        output_file_path (str): Caminho para salvar a planilha de saída.
    """
    try:
        # Carrega as planilhas
        df_input_1 = pd.read_excel(input_file_path_1, parse_dates=['BIRTH_DATE'])
        df_input_2 = pd.read_excel(input_file_path_2)

        # --- Etapa 1: Preparar os dados para o merge ---
        # Renomeia a coluna de identificador na segunda planilha para corresponder à primeira
        if 'IDENTIFIER' in df_input_2.columns:
            df_input_2_prepared = df_input_2.rename(columns={'IDENTIFIER': 'ID_NUMBER'})
        else:
            print("Erro: Coluna 'IDENTIFIER' não encontrada na segunda planilha de entrada. Não é possível fazer o cruzamento por IDENTIFIER/ID_NUMBER.")
            return

        # Define as colunas obrigatórias para a primeira planilha de entrada
        required_input1_cols = ['ID_NUMBER', 'BIRTH_DATE', 'GENDER', 'BIRTH_PLACE', 'BIRTH_COUNTRY']
        if all(col in df_input_1.columns for col in required_input1_cols):
            df_input_1_prepared = df_input_1[required_input1_cols]
        else:
            print("Erro: Uma ou mais colunas obrigatórias ('ID_NUMBER', 'BIRTH_DATE', 'GENDER', 'BIRTH_PLACE', 'BIRTH_COUNTRY') não encontradas na primeira planilha de entrada.")
            return

        # Seleciona as colunas necessárias da segunda planilha de entrada para o merge
        required_input2_cols = [
            'ID_NUMBER', 'PERSON_NAME', 'PERSON_ADDRESS',
            'PARENT1_NAME', 'PARENT1_ADDRESS', 'PARENT1_EMAIL', 'PARENT1_PHONE',
            'PARENT2_NAME', 'PARENT2_ADDRESS', 'PARENT2_EMAIL', 'PARENT2_PHONE'
        ]
        if not all(col in df_input_2_prepared.columns for col in required_input2_cols):
            print("Erro: Uma ou mais colunas necessárias não encontradas na segunda planilha de entrada após renomear 'IDENTIFIER'. Verifique:")
            for col in required_input2_cols:
                if col not in df_input_2_prepared.columns:
                    print(f"- {col}")
            return
        df_input_2_prepared = df_input_2_prepared[required_input2_cols]

        # --- Etapa 2: Realizar o Merge (Junção) ---
        df_merged = pd.merge(df_input_2_prepared, df_input_1_prepared, on='ID_NUMBER', how='left')

        # --- Etapa 3: Processar os dados e popular o DataFrame final ---
        num_rows = len(df_merged)

        # Função auxiliar para capitalizar a primeira letra de cada palavra (Title Case)
        def to_title_case(text):
            if pd.isna(text) or text == '':
                return ''
            return str(text).title()

        # Função auxiliar para converter para maiúsculas de forma segura
        def to_upper_case(text):
            if pd.isna(text) or text == '':
                return ''
            return str(text).upper()
        
        # Função auxiliar para converter para minúsculas de forma segura
        def to_lower_case(text):
            if pd.isna(text) or text == '':
                return ''
            return str(text).lower()

        # Lógica para processar o endereço da PESSOA
        address_line1_list, address_line2_list, address_line3_list, city_list, postal_code_list = [], [], [], [], []
        if 'PERSON_ADDRESS' in df_merged.columns:
            for full_address in df_merged['PERSON_ADDRESS'].astype(str).fillna(''):
                parts = full_address.split('-')
                while len(parts) < 5:
                    parts.append('')
                address_line1_list.append(to_title_case(parts[0]))
                address_line2_list.append(to_title_case(parts[1]))
                address_line3_list.append(to_title_case(parts[2]))
                raw_postal_code = parts[4].strip()
                clean_postal_code = re.sub(r'\D', '', raw_postal_code)
                postal_code_list.append(clean_postal_code)
                city_list.append(to_title_case(parts[3]))
        else:
            print("Aviso: Coluna 'PERSON_ADDRESS' não encontrada. As colunas de endereço da Pessoa não serão preenchidas.")
            for _ in range(num_rows):
                address_line1_list.append('')
                address_line2_list.append('')
                address_line3_list.append('')
                city_list.append('')
                postal_code_list.append('')

        # --- Processamento dos dados do RESPONSÁVEL 1 e RESPONSÁVEL 2 com a nova lógica de prioridade e capitalização ---
        responsible1_last_name_list, responsible1_first_name_list = [], []
        responsible1_address_line1_list, responsible1_address_line2_list, responsible1_address_line3_list, responsible1_city_list, responsible1_postal_code_list = [], [], [], [], []
        responsible1_email_list, responsible1_phone_list = [], []

        responsible2_last_name_list, responsible2_first_name_list = [], []
        responsible2_address_line1_list, responsible2_address_line2_list, responsible2_address_line3_list, responsible2_city_list, responsible2_postal_code_list = [], [], [], [], []
        responsible2_email_list, responsible2_phone_list = [], []
        responsible2_role_list = []

        for index, row in df_merged.iterrows():
            # Flag para verificar se o PARENT1 original tinha dados
            parent1_was_originally_present = False

            # --- Dados do Responsável 1 (PARENT1) ---
            pai_nome = str(row.get('PARENT1_NAME', '')).strip()
            pai_endereco = str(row.get('PARENT1_ADDRESS', '')).strip()
            pai_email = str(row.get('PARENT1_EMAIL', '')).strip()
            pai_tel1 = str(row.get('PARENT1_PHONE', '')).strip()

            # Processa nome do Responsável 1 (SOBRENOME em MAIÚSCULAS, PRIMEIRO NOME em Title Case)
            current_responsible1_last_name = ''
            current_responsible1_first_name = ''
            if pai_nome and pai_nome.lower() != 'nan':
                parent1_was_originally_present = True # Define a flag se PARENT1 tinha dados
                if ' ' in pai_nome:
                    first_space_idx = pai_nome.find(' ')
                    current_responsible1_first_name = to_title_case(pai_nome[:first_space_idx])
                    current_responsible1_last_name = to_upper_case(pai_nome[first_space_idx + 1:])
                else:
                    current_responsible1_last_name = to_upper_case(pai_nome)

            responsible1_last_name_list.append(current_responsible1_last_name)
            responsible1_first_name_list.append(current_responsible1_first_name)

            # Processa endereço do Responsável 1 (em Title Case)
            current_responsible1_address_line1, current_responsible1_address_line2, current_responsible1_address_line3, current_responsible1_city, current_responsible1_postal_code = '', '', '', '', ''
            if pai_endereco and pai_endereco.lower() != 'nan':
                parent1_parts = pai_endereco.split('-')
                while len(parent1_parts) < 5:
                    parent1_parts.append('')
                current_responsible1_address_line1 = to_title_case(parent1_parts[0])
                current_responsible1_address_line2 = to_title_case(parent1_parts[1])
                current_responsible1_address_line3 = to_title_case(parent1_parts[2])
                current_responsible1_postal_code = re.sub(r'\D', '', parent1_parts[4].strip())
                current_responsible1_city = to_title_case(parent1_parts[3])

            responsible1_address_line1_list.append(current_responsible1_address_line1)
            responsible1_address_line2_list.append(current_responsible1_address_line2)
            responsible1_address_line3_list.append(current_responsible1_address_line3)
            responsible1_postal_code_list.append(current_responsible1_postal_code)
            responsible1_city_list.append(current_responsible1_city)

            # Processa email do Responsável 1 (em minúsculas)
            current_responsible1_email = to_lower_case(pai_email) if pai_email and pai_email.lower() != 'nan' else ''
            responsible1_email_list.append(current_responsible1_email)

            # Processa telefone do Responsável 1
            current_responsible1_phone = re.sub(r'\D', '', pai_tel1) if pai_tel1 and pai_tel1.lower() != 'nan' else ''
            responsible1_phone_list.append(current_responsible1_phone)

            # --- Dados do Responsável 2 (PARENT2) (para preencher Responsible1 se Responsible1 estiver vazio, ou Responsible2) ---
            mae_nome = str(row.get('PARENT2_NAME', '')).strip()
            mae_endereco = str(row.get('PARENT2_ADDRESS', '')).strip()
            mae_email = str(row.get('PARENT2_EMAIL', '')).strip()
            mae_tel1 = str(row.get('PARENT2_PHONE', '')).strip()

            # Processa nome do Responsável 2 (SOBRENOME em MAIÚSCULAS, PRIMEIRO NOME em Title Case)
            current_parent2_last_name = ''
            current_parent2_first_name = ''
            if mae_nome and mae_nome.lower() != 'nan':
                if ' ' in mae_nome:
                    first_space_idx = mae_nome.find(' ')
                    current_parent2_first_name = to_title_case(mae_nome[:first_space_idx])
                    current_parent2_last_name = to_upper_case(mae_nome[first_space_idx + 1:])
                else:
                    current_parent2_last_name = to_upper_case(mae_nome)

            # Processa endereço do Responsável 2 (em Title Case)
            current_parent2_address_line1, current_parent2_address_line2, current_parent2_address_line3, current_parent2_city, current_parent2_postal_code = '', '', '', '', ''
            if mae_endereco and mae_endereco.lower() != 'nan':
                parent2_parts = mae_endereco.split('-')
                while len(parent2_parts) < 5:
                    parent2_parts.append('')
                current_parent2_address_line1 = to_title_case(parent2_parts[0])
                current_parent2_address_line2 = to_title_case(parent2_parts[1])
                current_parent2_address_line3 = to_title_case(parent2_parts[2])
                current_parent2_postal_code = re.sub(r'\D', '', parent2_parts[4].strip())
                current_parent2_city = to_title_case(parent2_parts[3])
            
            # Processa email do Responsável 2 (em minúsculas)
            current_parent2_email = to_lower_case(mae_email) if mae_email and mae_email.lower() != 'nan' else ''

            # Processa telefone do Responsável 2
            current_parent2_phone = re.sub(r'\D', '', mae_tel1) if mae_tel1 and mae_tel1.lower() != 'nan' else ''

            # --- Lógica de Prioridade: Se Responsável 1 estiver vazio, preencher com dados do Responsável 2, senão em Responsável 2 ---
            # Nome
            if not responsible1_last_name_list[-1] and not responsible1_first_name_list[-1]:
                responsible1_last_name_list[-1] = current_parent2_last_name
                responsible1_first_name_list[-1] = current_parent2_first_name
                responsible2_last_name_list.append('')
                responsible2_first_name_list.append('')
            else:
                responsible2_last_name_list.append(current_parent2_last_name)
                responsible2_first_name_list.append(current_parent2_first_name)

            # Endereço
            if not responsible1_address_line1_list[-1] and not responsible1_address_line2_list[-1] and \
               not responsible1_address_line3_list[-1] and not responsible1_postal_code_list[-1] and not responsible1_city_list[-1]:
                responsible1_address_line1_list[-1] = current_parent2_address_line1
                responsible1_address_line2_list[-1] = current_parent2_address_line2
                responsible1_address_line3_list[-1] = current_parent2_address_line3
                responsible1_postal_code_list[-1] = current_parent2_postal_code
                responsible1_city_list[-1] = current_parent2_city
                responsible2_address_line1_list.append('')
                responsible2_address_line2_list.append('')
                responsible2_address_line3_list.append('')
                responsible2_postal_code_list.append('')
                responsible2_city_list.append('')
            else:
                responsible2_address_line1_list.append(current_parent2_address_line1)
                responsible2_address_line2_list.append(current_parent2_address_line2)
                responsible2_address_line3_list.append(current_parent2_address_line3)
                responsible2_postal_code_list.append(current_parent2_postal_code)
                responsible2_city_list.append(current_parent2_city)

            # Email
            if not responsible1_email_list[-1]:
                responsible1_email_list[-1] = current_parent2_email
                responsible2_email_list.append('')
            else:
                responsible2_email_list.append(current_parent2_email)

            # Telefone
            if not responsible1_phone_list[-1]:
                responsible1_phone_list[-1] = current_parent2_phone
                responsible2_phone_list.append('')
            else:
                responsible2_phone_list.append(current_parent2_phone)

            # Preencher Responsible2_Role:
            # Se o PARENT1 original tinha dados E o PARENT2 também tem dados, então há um segundo responsável.
            # Caso contrário (PARENT1 ausente e PARENT2 preencheu o Responsible1), o Responsible2_Role deve ser vazio.
            if parent1_was_originally_present and (mae_nome and mae_nome.lower() != 'nan'):
                responsible2_role_list.append("LEGAL")
            else:
                responsible2_role_list.append("") # Deixa vazio se não houver um segundo responsável de fato


        # Definindo a ordem EXATA das colunas no DataFrame final e preenchendo com as listas
        df_final = pd.DataFrame({
            'RECORD_NUMBER': [f"{i}.001" for i in range(1, num_rows + 1)],
            'TITLE': '',
            'LAST_NAME': '',
            'FIRST_NAME': '',
            'BIRTH_DATE': '',
            'GENDER': '',
            'BIRTH_PLACE': '',
            'BIRTH_COUNTRY': '',
            'ID_NUMBER': '',
            'ADDRESS_LINE_1': address_line1_list,
            'ADDRESS_LINE_2': address_line2_list,
            'ADDRESS_LINE_3': address_line3_list,
            'POSTAL_CODE': postal_code_list,
            'CITY': city_list,
            'RESPONSIBLE1_ROLE': 'LEGAL',
            'RESPONSIBLE1_LAST_NAME': responsible1_last_name_list,
            'RESPONSIBLE1_FIRST_NAME': responsible1_first_name_list,
            'RESPONSIBLE1_ADDRESS_LINE_1': responsible1_address_line1_list,
            'RESPONSIBLE1_ADDRESS_LINE_2': responsible1_address_line2_list,
            'RESPONSIBLE1_ADDRESS_LINE_3': responsible1_address_line3_list,
            'RESPONSIBLE1_POSTAL_CODE': responsible1_postal_code_list,
            'RESPONSIBLE1_CITY': responsible1_city_list,
            'RESPONSIBLE1_EMAIL': responsible1_email_list,
            'RESPONSIBLE1_PHONE': responsible1_phone_list,
            'RESPONSIBLE2_ROLE': responsible2_role_list, # Agora preenchido condicionalmente
            'RESPONSIBLE2_LAST_NAME': responsible2_last_name_list,
            'RESPONSIBLE2_FIRST_NAME': responsible2_first_name_list,
            'RESPONSIBLE2_ADDRESS_LINE_1': responsible2_address_line1_list,
            'RESPONSIBLE2_ADDRESS_LINE_2': responsible2_address_line2_list,
            'RESPONSIBLE2_ADDRESS_LINE_3': responsible2_address_line3_list,
            'RESPONSIBLE2_POSTAL_CODE': responsible2_postal_code_list,
            'RESPONSIBLE2_CITY': responsible2_city_list,
            'RESPONSIBLE2_EMAIL': responsible2_email_list,
            'RESPONSIBLE2_PHONE': responsible2_phone_list
        })

        # Preenche outras colunas do df_final a partir do df_merged
        if 'PERSON_NAME' in df_merged.columns:
            df_merged['PERSON_NAME'] = df_merged['PERSON_NAME'].astype(str).fillna('')
            split_person_names = df_merged['PERSON_NAME'].str.split(',', n=1, expand=True)
            df_final['LAST_NAME'] = split_person_names[0].fillna('').str.upper()
            df_final['FIRST_NAME'] = split_person_names[1].str.strip().fillna('').apply(to_title_case)

        if 'BIRTH_DATE' in df_merged.columns:
            dates_to_format = pd.to_datetime(df_merged['BIRTH_DATE'], errors='coerce')
            df_final['BIRTH_DATE'] = dates_to_format.dt.strftime('%d/%m/%Y').fillna('')

        if 'GENDER' in df_merged.columns:
            df_final['GENDER'] = df_merged['GENDER'].fillna('').apply(to_title_case)

        if 'BIRTH_PLACE' in df_merged.columns:
            df_final['BIRTH_PLACE'] = df_merged['BIRTH_PLACE'].fillna('').apply(to_title_case)

        if 'BIRTH_COUNTRY' in df_merged.columns:
            df_final['BIRTH_COUNTRY'] = df_merged['BIRTH_COUNTRY'].fillna('').apply(to_title_case)

        if 'ID_NUMBER' in df_merged.columns:
            df_final['ID_NUMBER'] = df_merged['ID_NUMBER'].astype(str).str.zfill(7)

        # Salva a planilha final
        df_final.to_excel(output_file_path, index=False)
        print(f"Planilha final salva com sucesso em: {output_file_path}")

    except FileNotFoundError:
        print("Erro: Verifique se os caminhos dos arquivos estão corretos.")
    except KeyError as e:
        print(f"Erro: Coluna não encontrada. Verifique os nomes das colunas nas suas planilhas. Detalhe: {e}")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

# --- Exemplo de como usar o script ---
if __name__ == "__main__":
    input_file_path_1 = 'input_file_1.xlsx'
    input_file_path_2 = 'input_file_2.xlsx'
    output_file_path = 'output_file3.xlsx'

    process_data_files(input_file_path_1, input_file_path_2, output_file_path)
