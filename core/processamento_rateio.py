# Arquivo: core/processamento_rateio.py
# VERSÃO COMPLETA E CORRIGIDA
import os
import pandas as pd
import numpy as np
from datetime import datetime
import configparser

def validar_planilha_rateio(file_path):
    """Verifica se a planilha contém as abas e colunas necessárias para o processo de rateio."""
    try:
        xls = pd.ExcelFile(file_path)
        abas_necessarias = ['Controle de Radios', 'Controle de Componentes', 'Controle Solicitações (488 Tot)']
        for aba in abas_necessarias:
            if aba not in xls.sheet_names:
                return False, f"Aba necessária '{aba}' não encontrada na planilha."
        
        df_radios_cols = pd.read_excel(file_path, sheet_name='Controle de Radios', nrows=0).columns
        if 'Gerência' not in df_radios_cols or 'Valor' not in df_radios_cols:
            return False, "Aba 'Controle de Radios' não contém as colunas 'Gerência' e/ou 'Valor'."

        df_budget_cols = pd.read_excel(file_path, sheet_name='Controle Solicitações (488 Tot)', nrows=0).columns
        if 'Gerencia Padronizada' not in df_budget_cols:
            return False, "Aba 'Controle Solicitações (488 Tot)' não contém a coluna 'Gerencia Padronizada'."

        return True, ""
    except Exception as e:
        return False, f"Erro ao validar a planilha: {e}"

def processar_sumario_radios(df):
    df['Data'] = pd.to_datetime(df['Data'])
    df.dropna(subset=['ID do rádio', 'Gerência', 'Centro de custo'], inplace=True)
    df_latest = df.sort_values('Data').drop_duplicates(subset='ID do rádio', keep='last')
    df_filtered = df_latest[df_latest['Atividade'].isin(['Entrega', 'Devolução'])].copy()
    pivot = pd.pivot_table(df_filtered, index=['Gerência', 'Centro de custo'], columns='Atividade', values='Valor', aggfunc='sum', fill_value=0)
    pivot_qtd = pd.pivot_table(df_filtered, index=['Gerência', 'Centro de custo'], columns='Atividade', values='ID do rádio', aggfunc='count', fill_value=0)
    df_summary = pd.concat([pivot_qtd, pivot], axis=1)
    for col in ['Devolução', 'Entrega']:
        if col not in df_summary:
            df_summary[col] = 0
    df_summary.columns = ['Devolução Qtd', 'Entrega Qtd', 'Devolução Valor', 'Entrega Valor']
    df_summary.rename(columns={'Entrega Qtd': 'Qtd_Radios_Atual', 'Entrega Valor': 'Valor_Radios_Atual'}, inplace=True)
    return df_summary.reset_index()

def processar_corte_fevereiro(df):
    df['Data'] = pd.to_datetime(df['Data'])
    df_latest = df.sort_values('Data').drop_duplicates(subset='ID do rádio', keep='last')
    data_corte = pd.to_datetime("2024-02-01")
    df_corte = df_latest[(df_latest['Atividade'] == 'Entrega') & (df_latest['Data'] < data_corte)]
    df_summary_corte = df_corte.groupby('Gerência').agg(
        Qtd_Radios_Corte_Fevereiro=('ID do rádio', 'count'),
        Valor_Radios_Corte_Fevereiro=('Valor', 'sum')
    ).reset_index()
    return df_summary_corte

def processar_baterias(df):
    df_baterias = df[df['Equipamento'] == 'Bateria'].copy()
    pivot = pd.pivot_table(df_baterias, index=['Gerência', 'Centro de custo'], columns='Atividade', values=['Quantidade', 'Valor'], aggfunc='sum', fill_value=0)
    pivot.columns = ['_'.join(col).strip() for col in pivot.columns.values]
    return pivot.reset_index()

def processar_budget(df):
    df_summary_budget = df.groupby('Gerencia Padronizada').agg(
        Qtd_Radios_Budget=('Gerencia Padronizada', 'count'),
        Valor_Radios_Budget=('Valor', 'sum')
    ).reset_index().rename(columns={'Gerencia Padronizada': 'Gerência'})
    return df_summary_budget

def processar_generic(df, tipo_atividade, nome_col_sufixo):
    df['Data'] = pd.to_datetime(df['Data'])
    df_filtered = df[(df['Atividade'] == tipo_atividade) & (df['Data'].dt.month == pd.Timestamp.now().month)].copy()
    agg_dict = {'Valor': 'sum'}
    if 'Quantidade' in df.columns:
        agg_dict['Quantidade'] = 'sum'
    else:
        agg_dict['ID do rádio'] = 'count'
    summary = df_filtered.groupby(['Gerência', 'Centro de custo']).agg(agg_dict).reset_index()
    rename_map = {'Valor': f'Valor_{nome_col_sufixo}'}
    if 'Quantidade' in summary.columns:
        rename_map['Quantidade'] = f'Qtd_{nome_col_sufixo}'
    if 'ID do rádio' in summary.columns:
        rename_map['ID do rádio'] = f'Qtd_{nome_col_sufixo}'
    summary.rename(columns=rename_map, inplace=True)
    return summary

def consolidar_e_calcular_rateio(dataframes):
    df_final = dataframes['sumario_radios']
    dfs_to_merge = [
        (dataframes['corte_fevereiro'], 'Gerência'),
        (dataframes['budget'], 'Gerência'),
        (dataframes['baterias'], ['Gerência', 'Centro de custo']),
        (dataframes['ressarc_radios'], ['Gerência', 'Centro de custo']),
        (dataframes['ressarc_comp'], ['Gerência', 'Centro de custo']),
        (dataframes['instal_radios'], ['Gerência', 'Centro de custo']),
    ]
    for df_to_merge, key in dfs_to_merge:
        if not df_to_merge.empty:
            df_final = pd.merge(df_final, df_to_merge, on=key, how='left')
    df_final.fillna(0, inplace=True)
    colunas_necessarias = [
        'Qtd_Radios_Atual', 'Valor_Radios_Atual', 'Qtd_Radios_Corte_Fevereiro', 'Valor_Radios_Corte_Fevereiro',
        'Qtd_Radios_Budget', 'Valor_Radios_Budget', 'Valor_Ressarcimento_Radios', 'Valor_Entrega_Baterias',
        'Valor_Ressarcimento_Componentes', 'Valor_Devolucao_Baterias'
    ]
    for col in colunas_necessarias:
        if col not in df_final.columns:
            df_final[col] = 0
    C, D = df_final['Qtd_Radios_Atual'], df_final['Valor_Radios_Atual']
    E, F = df_final['Qtd_Radios_Corte_Fevereiro'], df_final['Valor_Radios_Corte_Fevereiro']
    G, H = df_final['Qtd_Radios_Budget'], df_final['Valor_Radios_Budget']
    L, N = df_final['Valor_Ressarcimento_Radios'], df_final['Valor_Entrega_Baterias']
    P, R = df_final['Valor_Ressarcimento_Componentes'], df_final['Valor_Devolucao_Baterias']
    T = R
    df_final['Qtd_Radios_Rateio'] = np.where(E <= G, C - E, C - G)
    df_final['Valor_Radios_Rateio'] = np.where(E <= G, D - F, D - H)
    J = df_final['Valor_Radios_Rateio']
    df_final['Coluna_U'] = np.maximum(0, P - T)
    U = df_final['Coluna_U']
    df_final['Total_a_Cobrar_Rateio'] = J + L + N + R + U
    if not df_final.empty:
        total_row_df = pd.DataFrame([df_final.sum(numeric_only=True)])
        total_row_df['Gerência'] = 'TOTAL'
        df_final = pd.concat([df_final, total_row_df], ignore_index=True)
    return df_final

def run_full_process(file_path, progress_callback):
    """Orquestra todo o processo de ETL, lendo configurações de um arquivo externo."""
    try:
        # Lendo o arquivo de configuração
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')
        
        # Nomes das abas a partir do config
        aba_radios = config['Abas']['controle_radios']
        aba_componentes = config['Abas']['controle_componentes']
        aba_budget = config['Abas']['controle_solicitacoes']
        
        # ETAPA DE VALIDAÇÃO (agora também usa o config)
        progress_callback.emit("--- Validando estrutura da planilha... ---")
        # A função de validação agora precisará dos nomes do config, então a integramos aqui
        xls = pd.ExcelFile(file_path)
        for aba in [aba_radios, aba_componentes, aba_budget]:
            if aba not in xls.sheet_names:
                progress_callback.emit(f"ERRO DE VALIDAÇÃO: Aba necessária '{aba}' não encontrada.")
                return False
        progress_callback.emit("Planilha validada com sucesso.")

        # ... (O restante das Fases 1, 2 e 3 continua exatamente igual) ...
        progress_callback.emit("\n--- Fase 1: Carregamento dos Dados ---")
        df_radios = pd.read_excel(file_path, sheet_name='Controle de Radios', engine='openpyxl')
        df_componentes = pd.read_excel(file_path, sheet_name='Controle de Componentes', engine='openpyxl')
        df_budget = pd.read_excel(file_path, sheet_name='Controle Solicitações (488 Tot)', engine='openpyxl')
        progress_callback.emit("Planilhas carregadas com sucesso.")
        progress_callback.emit("\n--- Fase 2: Processamento e Agregação ---")
        dfs = {}
        progress_callback.emit("1/7 - Processando sumário principal de rádios...")
        dfs['sumario_radios'] = processar_sumario_radios(df_radios.copy())
        progress_callback.emit("2/7 - Processando dados do corte de Fevereiro...")
        dfs['corte_fevereiro'] = processar_corte_fevereiro(df_radios.copy())
        progress_callback.emit("3/7 - Processando dados de Baterias...")
        dfs['baterias'] = processar_baterias(df_componentes.copy())
        progress_callback.emit("4/7 - Processando dados de Budget...")
        dfs['budget'] = processar_budget(df_budget.copy())
        progress_callback.emit("5/7 - Processando Ressarcimentos de Rádios...")
        dfs['ressarc_radios'] = processar_generic(df_radios.copy(), 'Ressarcimento', 'Ressarcimento_Radios')
        progress_callback.emit("6/7 - Processando Ressarcimentos de Componentes...")
        dfs['ressarc_comp'] = processar_generic(df_componentes.copy(), 'Ressarcimento', 'Ressarcimento_Componentes')
        progress_callback.emit("7/7 - Processando Instalações de Rádios...")
        dfs['instal_radios'] = processar_generic(df_radios.copy(), 'Instalação', 'Instalacao_Radios')
        progress_callback.emit("\n--- Fase 3: Consolidação e Cálculos Finais ---")
        df_rateio_final = consolidar_e_calcular_rateio(dfs)
        
        # --- MELHORIA: Salvar em pasta dedicada ---
        output_dir = "Relatorios_Rateio"
        os.makedirs(output_dir, exist_ok=True) # Cria a pasta se não existir

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = f'Rateio_Final_{timestamp}.xlsx'
        output_path = os.path.join(output_dir, file_name)
        
        progress_callback.emit(f"\n--- Fase 4: Exportando resultado para '{output_path}' ---")
        df_rateio_final.to_excel(output_path, index=False, sheet_name='Rateio')
        progress_callback.emit(f"SUCESSO: Arquivo gerado na pasta '{output_dir}'.")
        return True
        
    except Exception as e:
        progress_callback.emit(f"\nERRO DURANTE O PROCESSAMENTO: {e}")
        return False