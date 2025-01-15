import streamlit as st
import pandas as pd
import numpy as np
import plotly.figure_factory as ff
import matplotlib.pyplot as plt

# Função para calcular o match
def calcular_match(cangaceiros, empresas):
    match_matrix = pd.DataFrame(index=cangaceiros['Cangaceiro'], columns=empresas['Empresa'])
    
    for cangaceiro in cangaceiros.itertuples(index=False, name=None):
        for empresa in empresas.itertuples(index=False, name=None):
            cangaceiro_nome, afinidade_g1, afinidade_g2, afinidade_g3, \
            afinidade_i1, afinidade_i2, afinidade_i3, afinidade_i4, cangaceiro_localidade = cangaceiro
            
            empresa_nome, area_atuacao, maior_necessidade, segunda_necessidade, empresa_localidade = empresa
            
            afinidade_geral = max(afinidade_g1, afinidade_g2, afinidade_g3)
            
            maior_necessidade_match = 0
            if maior_necessidade == 'Projetos e Modelo de Negócios (proposta de valor, execução de projetos, cadeia de valor, inovação em soluções)':
                maior_necessidade_match = afinidade_i1
            elif maior_necessidade == 'Vendas e Mercado (processo de vendas, estratégia comercial, aquisição de clientes, retenção de clientes)':
                maior_necessidade_match = afinidade_i2
            elif maior_necessidade == 'Gestão e Operações (planejamento estratégico, sistema de gestão, gestão financeira, estrutura organizacional)':
                maior_necessidade_match = afinidade_i3
            elif maior_necessidade == 'Time e Cultura (atração e retenção de membros, formação de lideranças, engajamento do time, cultura organizacional)':
                maior_necessidade_match = afinidade_i4
            
            segunda_necessidade_match = 0
            if segunda_necessidade == 'Projetos e Modelo de Negócios (proposta de valor, execução de projetos, cadeia de valor, inovação em soluções)':
                segunda_necessidade_match = afinidade_i1
            elif segunda_necessidade == 'Vendas e Mercado (processo de vendas, estratégia comercial, aquisição de clientes, retenção de clientes)':
                segunda_necessidade_match = afinidade_i2
            elif segunda_necessidade == 'Gestão e Operações (planejamento estratégico, sistema de gestão, gestão financeira, estrutura organizacional)':
                segunda_necessidade_match = afinidade_i3
            elif segunda_necessidade == 'Time e Cultura (atração e retenção de membros, formação de lideranças, engajamento do time, cultura organizacional)':
                segunda_necessidade_match = afinidade_i4
            
            localidade_match = 0.5 if cangaceiro_localidade == empresa_localidade else 0
            
            total_match = (afinidade_geral * 1.5) + (maior_necessidade_match * 3) + (segunda_necessidade_match * 2) + localidade_match
            match_matrix.at[cangaceiro_nome, empresa_nome] = round(total_match / 3.3, 2) 
    
    return match_matrix

# Carregar os dados
cangaceiros = pd.read_excel(r'C:\Users\leona\OneDrive\Área de Trabalho\Python\RN\Base_Cangaceiros_Empresas.xlsx', sheet_name='Cangaceiros2')
empresas = pd.read_excel(r'C:\Users\leona\OneDrive\Área de Trabalho\Python\RN\Base_Cangaceiros_Empresas.xlsx', sheet_name='Empresas2')


# Calcular a matriz de match
match_matrix = calcular_match(cangaceiros, empresas)

# Iniciar a aplicação Streamlit
st.title('Dashboard Interativo - Mapa de Compatibilidade')

# Selecionar o nível de zoom
zoom = st.slider("Escolha o tamanho do gráfico", 5, 30, 15)

# Criar o mapa de calor com plotly
fig = ff.create_annotated_heatmap(
    z=match_matrix.values.astype(float),
    x = [
    " ".join(name.split()[:2]) if len(name.split()) > 1 else name
    for name in empresas['Empresa']
],
    y=match_matrix.index.tolist(),
    colorscale='YlGn',
    annotation_text=match_matrix.values.astype(str)
)

# Ajustar o layout
fig.update_layout(
    title="Mapa de Compatibilidade - Cangaceiros e Empresas",
    xaxis=dict(title="Empresas", tickangle=90),
    yaxis=dict(title="Cangaceiros"),
    autosize=True,
    width=zoom * 100,
    height=zoom * 100
)

# Mostrar o gráfico interativo
st.plotly_chart(fig, use_container_width=True)

# Download da matriz como arquivo Excel
@st.cache_data
def convert_to_excel(df):
    output = pd.ExcelWriter("match_matrix.xlsx")
    df.to_excel(output, index=True)
    output.close()
    return output

if st.button("Baixar Matriz em Excel"):
    st.download_button(
        label="Download",
        data=convert_to_excel(match_matrix).path,
        file_name="match_matrix.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
