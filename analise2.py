import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from pathlib import Path
import warnings
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import webbrowser
import unicodedata
import re

warnings.filterwarnings('ignore')

# Configurações iniciais
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('display.max_rows', 100)

# Definir caminho do arquivo
caminho_planilha = "C:/Users/walace.gorino/Documents/analise chamados do suporte.xlsx"
path = Path(caminho_planilha)
out_dir = path.parent

def normalizar_texto(texto):
    """Normaliza texto removendo acentos e caracteres especiais"""
    if not isinstance(texto, str):
        texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = texto.lower()
    # Remove caracteres especiais, mantém apenas letras, números e espaços
    texto = ''.join(c for c in texto if c.isalnum() or c.isspace())
    return texto

def carregar_dados():
    """Carrega e processa os dados para o dashboard"""
    try:
        if not path.exists():
            print(f"❌ ERRO: O arquivo não foi encontrado em '{path}'")
            print("👉 Por favor, edite a variável 'caminho_planilha' no código com o caminho correto.")
            return None
        print(f"Carregando planilha: {path.name}")
        df = pd.read_excel(path)
        print(f"✅ Dataset carregado com sucesso. Shape: {df.shape}")
        print(f"📊 Colunas disponíveis: {list(df.columns)}")
        return df
    except Exception as e:
        print(f"❌ Erro durante o carregamento: {str(e)}")
        return None

def encontrar_coluna_categoria(df):
    """Encontra automaticamente a coluna de categoria/solicitação"""
    palavras_chave = ['categoria', 'tipo', 'assunto', 'natureza', 'classificacao', 'descricao',
                      'solicitacao', 'problema', 'informado', 'chamado', 'motivo']
    melhor_coluna = None
    melhor_pontuacao = 0
    for col in df.columns:
        col_normalized = normalizar_texto(col)
        pontuacao = sum(1 for palavra in palavras_chave if palavra in col_normalized)
        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_coluna = col
    if melhor_coluna:
        return melhor_coluna
    colunas_nao_numericas = df.select_dtypes(exclude=[np.number]).columns
    return colunas_nao_numericas[0] if len(colunas_nao_numericas) > 0 else None

def encontrar_coluna_solucao(df):
    """Encontra automaticamente a coluna de solução"""
    palavras_chave = ['solucao', 'resolucao', 'apresentada', 'solucaoapresentada', 'resultado', 'procedimento']
    melhor_coluna = None
    melhor_pontuacao = 0
    for col in df.columns:
        col_normalized = normalizar_texto(col)
        pontuacao = sum(1 for palavra in palavras_chave if palavra in col_normalized)
        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_coluna = col
    return melhor_coluna

def encontrar_coluna_data(df):
    """Encontra automaticamente a coluna de data"""
    palavras_chave = ['data', 'date', 'abertura', 'criado', 'timestamp']
    melhor_coluna = None
    melhor_pontuacao = 0
    for col in df.columns:
        col_normalized = normalizar_texto(col)
        pontuacao = sum(1 for palavra in palavras_chave if palavra in col_normalized)
        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_coluna = col
    return melhor_coluna

def encontrar_coluna_status(df):
    """Encontra automaticamente a coluna de status"""
    palavras_chave = ['status', 'estado', 'situacao', 'fechamento', 'andamento']
    melhor_coluna = None
    melhor_pontuacao = 0
    for col in df.columns:
        col_normalized = normalizar_texto(col)
        pontuacao = sum(1 for palavra in palavras_chave if palavra in col_normalized)
        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_coluna = col
    return melhor_coluna

def processar_datas(df, coluna_data):
    """Processa colunas de data и extrai informações temporais"""
    if coluna_data and coluna_data in df.columns:
        try:
            df[coluna_data] = pd.to_datetime(df[coluna_data], errors='coerce')
            df['Ano'] = df[coluna_data].dt.year
            df['Mês'] = df[coluna_data].dt.month
            df['Dia'] = df[coluna_data].dt.day
            df['Dia_Semana'] = df[coluna_data].dt.day_name()
            df['Hora'] = df[coluna_data].dt.hour
            print("✅ Datas processadas com sucesso")
        except Exception as e:
            print(f"❌ Erro ao processar datas: {e}")
    return df

def analise_chamados(df):
    """Realiza análise específica de chamados"""
    print("\n" + "="*60)
    print("ANÁLISE GERAL DE CHAMADOS")
    print("="*60)
    df_clean = df.copy()
    coluna_categoria = encontrar_coluna_categoria(df_clean)
    if not coluna_categoria:
        print("❌ Não foi possível identificar uma coluna de categoria")
        return df_clean, None, None, None, None, None
    
    coluna_solucao = encontrar_coluna_solucao(df_clean)
    coluna_data = encontrar_coluna_data(df_clean)
    coluna_status = encontrar_coluna_status(df_clean)
    
    print(f"📋 Coluna de categoria identificada: '{coluna_categoria}'")
    if coluna_solucao:
        print(f"📋 Coluna de solução identificada: '{coluna_solucao}'")
    if coluna_data:
        print(f"📋 Coluna de data identificada: '{coluna_data}'")
    if coluna_status:
        print(f"📋 Coluna de status identificada: '{coluna_status}'")
    
    # Processar datas se disponível
    if coluna_data:
        df_clean = processar_datas(df_clean, coluna_data)
    
    print(f"\n📊 Estatísticas da coluna '{coluna_categoria}':")
    print(f"   Valores únicos: {df_clean[coluna_categoria].nunique()}")
    contagem_categorias = df_clean[coluna_categoria].value_counts()
    print(f"\n📈 Distribuição de categorias (top 10):")
    for i, (categoria, quantidade) in enumerate(contagem_categorias.head(10).items(), 1):
        percentual = (quantidade / len(df_clean)) * 100
        print(f"   {i}. {categoria}: {quantidade} chamados ({percentual:.1f}%)")
    
    return df_clean, coluna_categoria, coluna_solucao, coluna_data, coluna_status, contagem_categorias

def analisar_solucoes_por_categoria(df, col_categoria, col_solucao):
    """Agrupa por categoria e conta as soluções apresentadas."""
    if not col_categoria or not col_solucao:
        print("⚠️ Colunas de categoria e/ou solução não encontradas. Análise de soluções por categoria pulada.")
        return None

    print("\n" + "="*60)
    print("ANÁLISE DE SOLUÇÕES POR CATEGORIA DE PROBLEMA")
    print("="*60)

    # Agrupa pela categoria do problema e conta a frequência de cada solução
    solucoes_agrupadas = df.groupby(col_categoria)[col_solucao].value_counts().rename('Contagem')
    df_solucoes = solucoes_agrupadas.reset_index()

    # Exibir no console as top 5 categorias e suas top 3 soluções
    top_categorias = df[col_categoria].value_counts().nlargest(5).index
    print("🔍 Exibindo as soluções mais comuns para os problemas mais frequentes:")
    for categoria in top_categorias:
        print(f"\n--- Problema: '{categoria}' ---")
        top_solucoes = df_solucoes[df_solucoes[col_categoria] == categoria].nlargest(3, 'Contagem')
        if top_solucoes.empty:
            print("   (Nenhuma solução registrada para esta categoria)")
        else:
            for _, row in top_solucoes.iterrows():
                print(f"   -> Solução: '{row[col_solucao]}' ( aplicada {row['Contagem']} vezes )")

    return df_solucoes

def criar_graficos_interativos(df, coluna_categoria, coluna_solucao, coluna_data, coluna_status, contagem_categorias, df_solucoes):
    """Cria gráficos interativos para o dashboard"""
    
    # 1. Gráfico de distribuição de categorias (top 15)
    fig_categorias = px.bar(
        x=contagem_categorias.head(15).index, 
        y=contagem_categorias.head(15).values,
        title="Top 15 Categorias de Problemas",
        labels={'x': 'Categoria', 'y': 'Quantidade'},
        color=contagem_categorias.head(15).values,
        color_continuous_scale='reds'
    )
    fig_categorias.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='#fff',
        xaxis_tickangle=-45
    )
    
    # 2. Gráfico de pizza para distribuição percentual (corrigido)
    if len(contagem_categorias) > 10:
        top_10_categorias = contagem_categorias.head(10)
        outros_valor = contagem_categorias[10:].sum()
        categorias_pizza = pd.concat([top_10_categorias, pd.Series([outros_valor], index=['Outros'])])
    else:
        categorias_pizza = contagem_categorias
    
    fig_pizza = px.pie(
        values=categorias_pizza.values,
        names=categorias_pizza.index,
        title="Distribuição Percentual de Chamados por Categoria",
        color_discrete_sequence=px.colors.sequential.Reds
    )
    fig_pizza.update_traces(textposition='inside', textinfo='percent+label')
    fig_pizza.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='#fff',
        showlegend=False
    )
    
    # 3. Gráfico de tendência temporal (se dados de data disponíveis)
    fig_temporal = None
    if 'Mês' in df.columns and 'Ano' in df.columns:
        temporal_data = df.groupby(['Ano', 'Mês']).size().reset_index(name='Quantidade')
        temporal_data['Data'] = pd.to_datetime(temporal_data['Ano'].astype(str) + '-' + temporal_data['Mês'].astype(str))
        
        fig_temporal = px.line(
            temporal_data, 
            x='Data', 
            y='Quantidade',
            title='Evolução Temporal de Chamados',
            markers=True
        )
        fig_temporal.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='#fff'
        )
        fig_temporal.update_traces(line_color='#e50914')
    
    # 4. Gráfico de distribuição por dia da semana (se dados de data disponíveis)
    fig_dia_semana = None
    if 'Dia_Semana' in df.columns:
        dias_ordem = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        dias_portugues = {'Monday': 'Segunda', 'Tuesday': 'Terça', 'Wednesday': 'Quarta', 
                         'Thursday': 'Quinta', 'Friday': 'Sexta', 'Saturday': 'Sábado', 'Sunday': 'Domingo'}
        
        dia_semana_data = df['Dia_Semana'].value_counts().reindex(dias_ordem)
        dia_semana_data.index = dia_semana_data.index.map(dias_portugues)
        
        fig_dia_semana = px.bar(
            x=dia_semana_data.index,
            y=dia_semana_data.values,
            title='Chamados por Dia da Semana',
            labels={'x': 'Dia da Semana', 'y': 'Quantidade'},
            color=dia_semana_data.values,
            color_continuous_scale='reds'
        )
        fig_dia_semana.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='#fff'
        )
    
    # 5. Gráfico de distribuição por hora do dia (se dados disponíveis)
    fig_hora = None
    if 'Hora' in df.columns:
        hora_data = df['Hora'].value_counts().sort_index()
        
        fig_hora = px.bar(
            x=hora_data.index,
            y=hora_data.values,
            title='Chamados por Hora do Dia',
            labels={'x': 'Hora', 'y': 'Quantidade'},
            color=hora_data.values,
            color_continuous_scale='reds'
        )
        fig_hora.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='#fff'
        )
    
    # 6. Gráfico de status (se disponível)
    fig_status = None
    if coluna_status and coluna_status in df.columns:
        status_data = df[coluna_status].value_counts()
        
        fig_status = px.pie(
            values=status_data.values,
            names=status_data.index,
            title="Distribuição por Status",
            color_discrete_sequence=px.colors.sequential.Reds
        )
        fig_status.update_traces(textposition='inside', textinfo='percent+label')
        fig_status.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='#fff'
        )
    
    # 7. Heatmap de correlação entre hora e dia da semana (se dados disponíveis)
    fig_heatmap = None
    if 'Hora' in df.columns and 'Dia_Semana' in df.columns:
        dias_ordem = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        dias_portugues = {'Monday': 'Segunda', 'Tuesday': 'Terça', 'Wednesday': 'Quarta', 
                         'Thursday': 'Quinta', 'Friday': 'Sexta', 'Saturday': 'Sábado', 'Sunday': 'Domingo'}
        
        heatmap_data = df.groupby(['Dia_Semana', 'Hora']).size().unstack(fill_value=0)
        heatmap_data = heatmap_data.reindex(dias_ordem)
        heatmap_data.index = heatmap_data.index.map(dias_portugues)
        
        fig_heatmap = px.imshow(
            heatmap_data,
            title='Heatmap: Chamados por Dia da Semana e Hora',
            labels=dict(x="Hora do Dia", y="Dia da Semana", color="Quantidade"),
            color_continuous_scale='reds'
        )
        fig_heatmap.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='#fff'
        )
    
    return {
        'categorias': fig_categorias,
        'pizza': fig_pizza,
        'temporal': fig_temporal,
        'dia_semana': fig_dia_semana,
        'hora': fig_hora,
        'status': fig_status,
        'heatmap': fig_heatmap
    }

def criar_dashboard_interativo(df, coluna_categoria, coluna_solucao, coluna_data, coluna_status, contagem_categorias, df_solucoes, graficos):
    """Cria um dashboard interativo no estilo Netflix para análise de chamados"""
    total_chamados = len(df)
    total_categorias = len(contagem_categorias) if contagem_categorias is not None else 0
    
    # Preparar dados para tabela de soluções
    tabela_solucoes_html = ""
    if df_solucoes is not None and coluna_solucao and coluna_solucao in df_solucoes.columns:
        col_solucao_nome = coluna_solucao
        top_categorias_dash = df[coluna_categoria].value_counts().nlargest(5).index
        
        for categoria in top_categorias_dash:
            dados_categoria = df_solucoes[df_solucoes[coluna_categoria] == categoria].nlargest(5, 'Contagem')
            if not dados_categoria.empty:
                tabela_solucoes_html += f"<h4>Problema: {categoria}</h4>"
                tabela_solucoes_html += "<table><thead><tr><th>Solução Apresentada</th><th>Quantidade</th></tr></thead><tbody>"
                for _, row in dados_categoria.iterrows():
                    solucao_truncada = row[col_solucao_nome][:100] + "..." if len(str(row[col_solucao_nome])) > 100 else row[col_solucao_nome]
                    tabela_solucoes_html += f"<tr><td>{solucao_truncada}</td><td>{row['Contagem']}</td></tr>"
                tabela_solucoes_html += "</tbody></table><br>"
    else:
        tabela_solucoes_html = "<p>Não foi possível gerar a análise de soluções por categoria.</p>"

    # Converter gráficos para HTML
    grafico_categorias_html = graficos['categorias'].to_html(include_plotlyjs='cdn', div_id="categoria-chart")
    grafico_pizza_html = graficos['pizza'].to_html(include_plotlyjs=False, div_id="pizza-chart")
    
    grafico_temporal_html = ""
    if graficos['temporal']:
        grafico_temporal_html = graficos['temporal'].to_html(include_plotlyjs=False, div_id="temporal-chart")
    
    grafico_dia_semana_html = ""
    if graficos['dia_semana']:
        grafico_dia_semana_html = graficos['dia_semana'].to_html(include_plotlyjs=False, div_id="dia-semana-chart")
    
    grafico_hora_html = ""
    if graficos['hora']:
        grafico_hora_html = graficos['hora'].to_html(include_plotlyjs=False, div_id="hora-chart")
    
    grafico_status_html = ""
    if graficos['status']:
        grafico_status_html = graficos['status'].to_html(include_plotlyjs=False, div_id="status-chart")
    
    grafico_heatmap_html = ""
    if graficos['heatmap']:
        grafico_heatmap_html = graficos['heatmap'].to_html(include_plotlyjs=False, div_id="heatmap-chart")

    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Dashboard de Chamados</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            :root {{
                --netflix-red: #e50914;
                --netflix-dark: #141414;
                --netflix-gray: #2f2f2f;
                --netflix-light-gray: #b3b3b3;
            }}
            
            body {{ 
                background: var(--netflix-dark); 
                color: #fff; 
                font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; 
                margin: 0;
                padding: 0;
            }}
            
            .container {{ 
                max-width: 1400px; 
                margin: 0 auto; 
                padding: 20px; 
            }}
            
            .header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 20px 0;
                border-bottom: 1px solid var(--netflix-gray);
                margin-bottom: 30px;
            }}
            
            .logo {{
                color: var(--netflix-red);
                font-size: 2.5rem;
                font-weight: bold;
            }}
            
            .filters {{
                display: flex;
                gap: 15px;
                margin-bottom: 20px;
                flex-wrap: wrap;
            }}
            
            .filter-item {{
                background: var(--netflix-gray);
                border: none;
                color: white;
                padding: 10px 15px;
                border-radius: 4px;
                cursor: pointer;
            }}
            
            .section-title {{ 
                font-size: 1.8rem; 
                margin-bottom: 20px; 
                border-left: 4px solid var(--netflix-red); 
                padding-left: 10px;
                margin-top: 40px;
            }}
            
            .metrics-grid {{ 
                display: grid; 
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
                gap: 20px; 
                margin-bottom: 40px; 
            }}
            
            .metric-card {{ 
                background: var(--netflix-gray); 
                border-radius: 6px; 
                padding: 20px; 
                text-align: center;
                transition: transform 0.3s ease;
            }}
            
            .metric-card:hover {{
                transform: translateY(-5px);
                box-shadow: 0 10px 20px rgba(0,0,0,0.3);
            }}
            
            .metric-value {{ 
                font-size: 2.2rem; 
                font-weight: bold; 
                color: var(--netflix-red);
                margin-bottom: 5px;
            }}
            
            .metric-label {{ 
                font-size: 0.9rem; 
                color: var(--netflix-light-gray); 
            }}
            
            .chart-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
                gap: 25px;
                margin-bottom: 40px;
            }}
            
            .chart-container {{ 
                background: #181818; 
                border-radius: 8px; 
                padding: 20px; 
                margin-bottom: 30px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.5); 
            }}
            
            .chart-title {{
                font-size: 1.2rem;
                margin-bottom: 15px;
                color: white;
            }}
            
            table {{ 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 15px;
            }}
            
            th, td {{ 
                padding: 12px 15px; 
                text-align: left; 
                border-bottom: 1px solid #333; 
            }}
            
            th {{ 
                background-color: var(--netflix-gray); 
                color: var(--netflix-red);
            }}
            
            tr:hover {{ 
                background-color: var(--netflix-gray); 
            }}
            
            h1 {{ 
                text-align: center; 
                color: var(--netflix-red); 
                font-size: 2.5rem; 
                margin-bottom: 10px;
            }}
            
            .subtitle {{
                text-align: center;
                color: var(--netflix-light-gray);
                margin-bottom: 40px;
            }}
            
            @media (max-width: 768px) {{
                .chart-grid {{
                    grid-template-columns: 1fr;
                }}
                
                .metrics-grid {{
                    grid-template-columns: repeat(2, 1fr);
                }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div class="logo">ANÁLISE DE CHAMADOS</div>
                <div class="filters">
                    <button class="filter-item" onclick="filterData('all')">Todos</button>
                    <button class="filter-item" onclick="filterData('month')">Este Mês</button>
                    <button class="filter-item" onclick="filterData('quarter')">Este Trimestre</button>
                    <button class="filter-item" onclick="filterData('year')">Este Ano</button>
                </div>
            </div>
            
            <h1>Dashboard de Análise de Chamados</h1>
            <div class="subtitle">Visualização interativa dos dados de chamados de suporte</div>
            
            <h2 class="section-title">Visão Geral</h2>
            <div class="metrics-grid">
                <div class="metric-card">
                    <div class="metric-value">{total_chamados}</div>
                    <div class="metric-label">Total de Chamados</div>
                </div>
                <div class="metric-card">
                    <div class="metric-value">{total_categorias}</div>
                    <div class="metric-label">Categorias Diferentes</div>
                </div>
                <div class="metric-card">
                    <div class="metric-value">{contagem_categorias.max() if contagem_categorias is not None else 0}</div>
                    <div class="metric-label">Pico em uma Categoria</div>
                </div>
                <div class="metric-card">
                    <div class="metric-value">{contagem_categorias.min() if contagem_categorias is not None else 0}</div>
                    <div class="metric-label">Mínimo em uma Categoria</div>
                </div>
            </div>

            <div class="chart-grid">
                <div class="chart-container">
                    <div class="chart-title">Distribuição por Categoria (Top 15)</div>
                    {grafico_categorias_html}
                </div>
                
                <div class="chart-container">
                    <div class="chart-title">Distribuição Percentual</div>
                    {grafico_pizza_html}
                </div>
            </div>

            <div class="chart-grid">
                <div class="chart-container">
                    <div class="chart-title">Evolução Temporal</div>
                    {grafico_temporal_html if grafico_temporal_html else "<p>Dados temporais não disponíveis</p>"}
                </div>
                
                <div class="chart-container">
                    <div class="chart-title">Distribuição por Status</div>
                    {grafico_status_html if grafico_status_html else "<p>Dados de status não disponíveis</p>"}
                </div>
            </div>

            <div class="chart-grid">
                <div class="chart-container">
                    <div class="chart-title">Chamados por Dia da Semana</div>
                    {grafico_dia_semana_html if grafico_dia_semana_html else "<p>Dados de dia da semana não disponíveis</p>"}
                </div>
                
                <div class="chart-container">
                    <div class="chart-title">Chamados por Hora do Dia</div>
                    {grafico_hora_html if grafico_hora_html else "<p>Dados de hora não disponíveis</p>"}
                </div>
            </div>

            <div class="chart-container">
                <div class="chart-title">Heatmap: Chamados por Dia da Semana e Hora</div>
                {grafico_heatmap_html if grafico_heatmap_html else "<p>Dados insuficientes para heatmap</p>"}
            </div>
            
            <h2 class="section-title">Análise de Soluções por Problema</h2>
            <div class="chart-container">
                <h3>Soluções Mais Comuns para os Principais Problemas</h3>
                {tabela_solucoes_html}
            </div>

            <h2 class="section-title">Top 20 Categorias de Problemas (Geral)</h2>
            <div class="chart-container">
                <table>
                    <thead><tr><th>Categoria</th><th>Quantidade</th><th>Percentual</th></tr></thead>
                    <tbody>
    """
    if contagem_categorias is not None:
        for categoria, quantidade in contagem_categorias.head(20).items():
            percentual = (quantidade / total_chamados) * 100
            # CORREÇÃO: trocar 'quantitude' por 'quantidade'
            html_content += f"<tr><td>{categoria}</td><td>{quantidade}</td><td>{percentual:.1f}%</td></tr>"
    html_content += """
                    </tbody>
                </table>
            </div>
        </div>
        
        <script>
            // Função para filtros (simulada para este exemplo)
            function filterData(range) {{
                alert('Funcionalidade de filtro para ' + range + ' será implementada em uma versão futura');
                // Em uma implementação real, aqui viria o código para filtrar os dados
                // e atualizar os gráficos dinamicamente
            }}
            
            // Ajustar tamanho dos gráficos ao redimensionar a janela
            window.addEventListener('resize', function() {{
                // Em uma implementação real, os gráficos seriam redimensionados
            }});
        </script>
    </body>
    </html>
    """
    
    dashboard_path = out_dir / "dashboard_interativo_chamados.html"
    with open(dashboard_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"\n✅ Dashboard interativo salvo em: {dashboard_path}")
    try:
        webbrowser.open(str(dashboard_path))
    except Exception as e:
        print(f"ℹ️ Não foi possível abrar o navegador. Abra manualmente o arquivo: {dashboard_path}")
    return dashboard_path

def exportar_analises(contagem_categorias, df_solucoes):
    """Exporta estatísticas gerais e por categoria para um único arquivo Excel com abas."""
    excel_path = out_dir / "analise_completa_chamados.xlsx"
    try:
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            # Aba 1: Estatísticas Gerais
            if contagem_categorias is not None:
                df_estatisticas = pd.DataFrame({
                    'Categoria': contagem_categorias.index,
                    'Quantidade': contagem_categorias.values,
                    'Percentual': (contagem_categorias.values / contagem_categorias.sum() * 100).round(2)
                })
                df_estatisticas.to_excel(writer, sheet_name='Geral_Por_Categoria', index=False)

            # Aba 2: Soluções por Categoria
            if df_solucoes is not None:
                df_solucoes.to_excel(writer, sheet_name='Solucoes_Por_Categoria', index=False)

        print(f"✅ Análises exportadas para: {excel_path}")
        return excel_path
    except Exception as e:
        print(f"❌ Erro ao exportar análises para Excel: {e}")
        return None

# Executar a análise completa
if __name__ == "__main__":
    print("🔍 Iniciando análise de chamados...")
    print("=" * 50)
    
    df = carregar_dados()
    
    if df is not None:
        df_clean, coluna_categoria, coluna_solucao, coluna_data, coluna_status, contagem_categorias = analise_chamados(df)
        
        # Encontrar a coluna de solução e rodar a análise agrupada
        if coluna_solucao:
            print(f"📋 Coluna de solução identificada: '{coluna_solucao}'")
            df_solucoes_agrupadas = analisar_solucoes_por_categoria(df_clean, coluna_categoria, coluna_solucao)
        else:
            print("❌ Não foi possível identificar uma coluna de solução.")
            df_solucoes_agrupadas = None

        if coluna_categoria and contagem_categorias is not None:
            # Criar gráficos interativos
            print("\n📊 Criando gráficos interativos...")
            graficos = criar_graficos_interativos(df_clean, coluna_categoria, coluna_solucao, coluna_data, coluna_status, contagem_categorias, df_solucoes_agrupadas)
            
            # Criar dashboard interativo
            print("\n🎨 Criando dashboard interativo...")
            dashboard_path = criar_dashboard_interativo(df_clean, coluna_categoria, coluna_solucao, coluna_data, coluna_status, contagem_categorias, df_solucoes_agrupadas, graficos)
            
            # Exportar análises para Excel
            print("\n💾 Exportando análises para Excel...")
            excel_path = exportar_analises(contagem_categorias, df_solucoes_agrupadas)
            
            print(f"\n🎉 Análise concluída com sucesso!")
            print(f"📊 Dashboard interativo: {dashboard_path}")
            if excel_path:
                print(f"📋 Planilha com análises: {excel_path}")
            
            print(f"\n📈 RESUMO FINAL:")
            print(f"   • Total de chamados: {len(df_clean)}")
            print(f"   • Categorias diferentes: {len(contagem_categorias)}")
            print(f"   • Categoria mais frequente: '{contagem_categorias.index[0]}' ({contagem_categorias.values[0]} chamados)")
        else:
            print("❌ Não foi possível realizar a análise de categorias")
    else:
        print("❌ Análise interrompida. Não foi possível carregar os dados.")