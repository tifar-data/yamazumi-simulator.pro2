import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO
import time

# Configuração da página
st.set_page_config(
    page_title="Yamazumi Simulator Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título da aplicação
st.title("📊 Yamazumi Simulator Pro")
st.markdown("Sistema avançado para balanceamento de linhas de montagem de chicotes")

# Sidebar para upload e parâmetros
with st.sidebar:
    st.header("⚙️ Configurações")
    
    # Upload do arquivo Excel
    uploaded_file = st.file_uploader(
        "Carregar arquivo Excel", 
        type=["xlsx", "xls"],
        help="Faça upload do arquivo Excel com dados de tarefas e parâmetros"
    )
    
    st.divider()
    st.subheader("Parâmetros de Produção")
    
    # Entrada de parâmetros
    disponibilidade = st.number_input(
        "Disponibilidade por turno (minutos)",
        min_value=60,
        max_value=1440,
        value=480,
        step=30
    )
    
    demanda = st.number_input(
        "Demanda de peças por turno",
        min_value=10,
        max_value=1000,
        value=150,
        step=10
    )
    
    max_postos = st.number_input(
        "Número máximo de postos",
        min_value=1,
        max_value=20,
        value=10,
        step=1
    )
    
    tempo_ciclo_min = st.number_input(
        "Tempo mínimo de ciclo (minutos)",
        min_value=1.0,
        max_value=10.0,
        value=3.0,
        step=0.5
    )
    
    tempo_ciclo_max = st.number_input(
        "Tempo máximo de ciclo (minutos)", 
        min_value=1.0,
        max_value=10.0,
        value=4.0,
        step=0.5
    )

# Função para calcular balanceamento
def calcular_balanceamento(df_tarefas, tempo_ciclo, max_postos):
    """
    Calcula o balanceamento da linha baseado nas tarefas e tempo de ciclo
    """
    # Ordenar tarefas por sequência
    df_ordenado = df_tarefas.sort_values('Sequencia')
    
    postos = []
    tempo_posto_atual = 0
    tarefas_posto = []
    tempos_posto = []
    
    for _, tarefa in df_ordenado.iterrows():
        tempo_tarefa = tarefa['Tempo_Padrao_segundos']
        
        if tempo_posto_atual + tempo_tarefa <= tempo_ciclo or not tarefas_posto:
            tarefas_posto.append(tarefa['Tarefa'])
            tempos_posto.append(tempo_tarefa)
            tempo_posto_atual += tempo_tarefa
        else:
            # Finalizar posto atual
            eficiencia = (tempo_posto_atual / tempo_ciclo) * 100
            postos.append({
                'Tarefas': tarefas_posto.copy(),
                'Tempos': tempos_posto.copy(),
                'Tempo Total': tempo_posto_atual,
                'Tempo Ciclo': tempo_ciclo,
                'Eficiência': eficiencia
            })
            
            # Iniciar novo posto
            tarefas_posto = [tarefa['Tarefa']]
            tempos_posto = [tempo_tarefa]
            tempo_posto_atual = tempo_tarefa
    
    # Adicionar último posto
    if tarefas_posto:
        eficiencia = (tempo_posto_atual / tempo_ciclo) * 100
        postos.append({
            'Tarefas': tarefas_posto,
            'Tempos': tempos_posto,
            'Tempo Total': tempo_posto_atual,
            'Tempo Ciclo': tempo_ciclo,
            'Eficiência': eficiencia
        })
    
    return postos

# Função para gerar gráfico Yamazumi
def gerar_grafico_yamazumi(postos, tempo_ciclo):
    """
    Gera gráfico Yamazumi (gráfico de barras empilhadas)
    """
    fig, ax = plt.subplots(figsize=(14, 8))
    
    # Preparar dados para o gráfico
    postos_nomes = [f"Posto {i+1}" for i in range(len(postos))]
    tempos_totais = [p['Tempo Total'] for p in postos]
    
    # Criar barras empilhadas
    bottom = np.zeros(len(postos))
    
    for i, posto in enumerate(postos):
        for j, (tarefa, tempo) in enumerate(zip(posto['Tarefas'], posto['Tempos'])):
            ax.bar(postos_nomes[i], tempo, bottom=bottom[i], 
                  label=tarefa if i == 0 else "", alpha=0.7)
            bottom[i] += tempo
    
    # Linha de tempo de ciclo
    ax.axhline(y=tempo_ciclo, color='red', linestyle='--', 
              label=f'Tempo de Ciclo ({tempo_ciclo}s)')
    
    ax.set_xlabel('Postos de Trabalho', fontsize=12)
    ax.set_ylabel('Tempo (segundos)', fontsize=12)
    ax.set_title('Diagrama Yamazumi - Balanceamento de Linha', fontsize=16)
    ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    ax.grid(True, alpha=0.3)
    
    # Adicionar valores nas barras
    for i, (posto, total) in enumerate(zip(postos_nomes, tempos_totais)):
        ax.text(i, total + 5, f'{total:.1f}s', ha='center', va='bottom', fontweight='bold')
    
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    return fig

# Processamento do arquivo upload
if uploaded_file is not None:
    try:
        # Carregar abas do Excel
        df_tarefas = pd.read_excel(uploaded_file, sheet_name="Tarefas")
        df_parametros = pd.read_excel(uploaded_file, sheet_name="Parametros")
        
        # Verificar colunas necessárias
        colunas_necessarias = ["Tarefa", "Tempo_Padrao_segundos", "Tipo_Tarefa", "Sequencia"]
        for coluna in colunas_necessarias:
            if coluna not in df_tarefas.columns:
                st.error(f"Coluna '{coluna}' não encontrada na aba 'Tarefas'")
                st.stop()
        
        # Calcular tempo de ciclo
        tempo_ciclo = (disponibilidade * 60) / demanda  # em segundos
        
        # Calcular balanceamento
        postos = calcular_balanceamento(df_tarefas, tempo_ciclo, max_postos)
        
        # Exibir resultados
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("📋 Resultados do Balanceamento")
            
            for i, posto in enumerate(postos, 1):
                with st.expander(f"Posto {i} - {posto['Tempo Total']:.1f}s ({posto['Eficiência']:.1f}%)"):
                    st.write("**Tarefas:**")
                    for tarefa, tempo in zip(posto['Tarefas'], posto['Tempos']):
                        st.write(f"- {tarefa}: {tempo}s")
                    
                    st.write(f"**Tempo total:** {posto['Tempo Total']:.1f}s")
                    st.write(f"**Eficiência:** {posto['Eficiência']:.1f}%")
        
        with col2:
            st.subheader("📊 Diagrama Yamazumi")
            fig = gerar_grafico_yamazumi(postos, tempo_ciclo)
            st.pyplot(fig)
        
        # Métricas de desempenho
        st.subheader("📈 Métricas de Desempenho")
        
        tempo_total = sum(p['Tempo Total'] for p in postos)
        eficiencia_total = (tempo_total / (len(postos) * tempo_ciclo)) * 100
        
        col1, col2, col3, col4 = st.columns(4)
        
        col1.metric("Número de Postos", len(postos))
        col2.metric("Eficiência Total", f"{eficiencia_total:.1f}%")
        col3.metric("Tempo de Ciclo", f"{tempo_ciclo:.1f}s")
        col4.metric("Tempo Total da Linha", f"{tempo_total:.1f}s")
        
        # Download dos resultados
        st.subheader("💾 Exportar Resultados")
        
        # Criar Excel com resultados
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Salvar dados dos postos
            dados_postos = []
            for i, posto in enumerate(postos, 1):
                for tarefa, tempo in zip(posto['Tarefas'], posto['Tempos']):
                    dados_postos.append({
                        'Posto': f'Posto {i}',
                        'Tarefa': tarefa,
                        'Tempo (s)': tempo,
                        'Tempo Total Posto': posto['Tempo Total'],
                        'Eficiência Posto': posto['Eficiência']
                    })
            
            df_resultados = pd.DataFrame(dados_postos)
            df_resultados.to_excel(writer, sheet_name='Resultados', index=False)
            
            # Salhar métricas
            metricas = {
                'Métrica': ['Número de Postos', 'Eficiência Total', 'Tempo de Ciclo', 'Tempo Total Linha'],
                'Valor': [len(postos), f"{eficiencia_total:.1f}%", f"{tempo_ciclo:.1f}s", f"{tempo_total:.1f}s"]
            }
            df_metricas = pd.DataFrame(metricas)
            df_metricas.to_excel(writer, sheet_name='Métricas', index=False)
        
        st.download_button(
            label="📥 Baixar Resultados em Excel",
            data=output.getvalue(),
            file_name="resultados_balanceamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Erro ao processar arquivo: {str(e)}")
else:
    # Modo de demonstração com dados exemplo
    st.info("💡 Faça upload de um arquivo Excel ou use os dados de exemplo abaixo")
    
    # Dados de exemplo
    dados_exemplo = {
        'Tarefa': [
            'Corte dos fios', 'Descascamento de pontas', 'Aplicação de terminal A',
            'Aplicação de terminal B', 'Montagem do conector X', 'Montagem do conector Y',
            'Teste de continuidade', 'Encapamento', 'Inspeção final', 'Embalagem'
        ],
        'Tempo_Padrao_segundos': [45, 60, 75, 90, 120, 110, 85, 65, 50, 40],
        'Tipo_Tarefa': [
            'Value-Added', 'Value-Added', 'Value-Added', 'Value-Added', 'Value-Added',
            'Value-Added', 'Necessary Non-Value-Added', 'Value-Added', 
            'Necessary Non-Value-Added', 'Non-Value-Added'
        ],
        'Sequencia': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    }
    
    df_exemplo = pd.DataFrame(dados_exemplo)
    st.dataframe(df_exemplo, use_container_width=True)