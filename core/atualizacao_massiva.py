# Arquivo: core/atualizacao_massiva.py
import os
import pandas as pd
import unicodedata
from datetime import datetime

def validar_planilha_atualizacao(file_path):
    """Verifica se a planilha contém as abas e colunas necessárias para a atualização."""
    try:
        xls = pd.ExcelFile(file_path)
        abas_necessarias = ['Base de Dados', 'Controle de Radios', 'Controle de Componentes']
        for aba in abas_necessarias:
            if aba not in xls.sheet_names:
                return False, f"Aba necessária '{aba}' não encontrada na planilha."
        
        # Verificando a aba principal de referência
        df_base_cols = pd.read_excel(file_path, sheet_name='Base de Dados', nrows=0).columns
        colunas_base = ['Gerencia', 'CC Cobrança Rateio', 'GESTOR']
        for col in colunas_base:
            if col not in df_base_cols:
                return False, f"Aba 'Base de Dados' não contém a coluna essencial: '{col}'."
        
        return True, ""
    except Exception as e:
        return False, f"Erro ao validar a planilha: {e}"

def normalize_text(text):
    """Função para remover acentos e converter para minúsculas."""
    if not isinstance(text, str):
        return text
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text.lower().strip()

def run_update_process(input_path, progress_callback):
    """
    Executa a atualização massiva, agora salvando em uma pasta dedicada.
    """
    try:
        # ETAPA DE VALIDAÇÃO
        progress_callback.emit("--- Validando estrutura da planilha... ---")
        is_valid, error_message = validar_planilha_atualizacao(input_path)
        if not is_valid:
            progress_callback.emit(f"ERRO DE VALIDAÇÃO: {error_message}")
            return False
        progress_callback.emit("Planilha validada com sucesso.")

        progress_callback.emit("\n--- Iniciando Atualização Massiva ---")
        
        # ... (O restante da lógica de leitura e processamento continua igual) ...
        progress_callback.emit("1/4 - Lendo planilhas...")
        df_base = pd.read_excel(input_path, sheet_name='Base de Dados', engine='openpyxl')
        df_radios = pd.read_excel(input_path, sheet_name='Controle de Radios', engine='openpyxl')
        df_componentes = pd.read_excel(input_path, sheet_name='Controle de Componentes', engine='openpyxl')
        progress_callback.emit("Padronizando nomes de colunas...")
        for df in [df_base, df_radios, df_componentes]:
            df.columns = [normalize_text(col) for col in df.columns]
        progress_callback.emit("Leitura e padronização concluídas.")
        col_gerencia = 'gerencia'
        col_centro_custo_origem = 'cc cobranca rateio'
        col_area = 'area'
        col_gestor = 'gestor'
        if col_centro_custo_origem not in df_base.columns:
            raise KeyError(f"A coluna '{col_centro_custo_origem}' não foi encontrada na aba 'Base de Dados'. Colunas disponíveis: {list(df_base.columns)}")
        df_base_ref = df_base[[col_gerencia, col_centro_custo_origem, col_area, col_gestor]].copy()
        df_base_ref.rename(columns={
            col_centro_custo_origem: 'centro de custo atualizado',
            col_area: 'area atualizada',
            col_gestor: 'gerente atualizado'
        }, inplace=True)

        # --- MELHORIA: Salvar em pasta dedicada ---
        output_dir = "Controles_Atualizados"
        os.makedirs(output_dir, exist_ok=True) # Cria a pasta se não existir

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = f'Controle_Atualizado_{timestamp}.xlsx'
        output_path = os.path.join(output_dir, file_name)

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            sheets_to_update = { "Controle de Radios": df_radios, "Controle de Componentes": df_componentes }
            for i, (sheet_name, df_original) in enumerate(sheets_to_update.items(), 2):
                progress_callback.emit(f"{i}/4 - Processando aba '{sheet_name}'...")
                df_merged = pd.merge(df_original, df_base_ref, on=col_gerencia, how='left')
                df_merged['status da atualizacao'] = df_merged['centro de custo atualizado'].apply(
                    lambda x: 'Atualizado com sucesso' if pd.notna(x) else 'Gerencia nao encontrada'
                )
                df_merged['centro de custo'] = df_merged['centro de custo atualizado']
                df_merged['gerente'] = df_merged['gerente atualizado']
                df_final = df_merged.drop(columns=['centro de custo atualizado', 'area atualizada', 'gerente atualizado'])
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)
        
        progress_callback.emit("4/4 - Processo finalizado.")
        progress_callback.emit(f"SUCESSO: Arquivo gerado na pasta '{output_dir}'.")
        return True

    except KeyError as e:
        progress_callback.emit(f"\nERRO DE CHAVE: A coluna {e} não foi encontrada.")
        progress_callback.emit("Verifique se o nome da coluna no Excel corresponde exatamente ao esperado.")
        return False
    except Exception as e:
        progress_callback.emit(f"\nERRO DURANTE A ATUALIZAÇÃO: {e}")
        return False