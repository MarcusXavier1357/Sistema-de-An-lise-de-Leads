import pandas as pd
import numpy as np
import dash
from dash import dcc, html, Input, Output, State, callback, dash_table, no_update
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import re
import base64
import io
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import tempfile
import matplotlib.pyplot as plt
import seaborn as sns

# Configurar tema claro com Font Awesome
app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.LUX,
        "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css"
    ],
    suppress_callback_exceptions=True
)
app.title = "Sistema de Análise de Leads"
server = app.server

# Layout do aplicativo
app.layout = dbc.Container([
    dcc.Markdown(
        """
        <style>
            .custom-tab {
                background-color: #2c3e50;
                color: white;
                padding: 10px;
                border: 1px solid #34495e;
                border-radius: 5px;
                margin-right: 5px;
            }
            .custom-tab--selected {
                background-color: #3498db;
                color: white;
                border: 1px solid #2980b9; /* removido font-weight: bold */
            }
            .dropdown-toggle-custom::after {
            float: right;
            margin-top: 8px;
            }
            .dropdown-menu {
                max-height: 300px;
                overflow-y: auto;
                width: 100%;
            }
            .dropdown-item {
                padding: 5px 15px;
            }
        </style>
        """,
        dangerously_allow_html=True
    ),
    
    dbc.Row([
        dbc.Col([
            # Painel de controle
            html.H1("Sistema de Análise de Leads", className="text-center mt-3 mb-4"),
            
            # Upload de arquivo
            dcc.Upload(
                id='upload-data',
                children=dbc.Button("Carregar Planilha", color="primary", className="w-100 mb-4"),
                multiple=False
            ),
            
            # Seletor de cidade
            html.Label("Selecione a Cidade:", className="mt-3"),
            dcc.Dropdown(
                id='city-selector',
                options=[],  # Inicializado vazio
                value='',
                className="mb-4"
            ),
            
            # Seletor de período inicial
            html.Label("Período Inicial:", className="mt-2"),
            dcc.Dropdown(
                id='period-start',
                options=[],
                className="mb-3"
            ),
            
            # Seletor de período final
            html.Label("Período Final:", className="mt-2"),
            dcc.Dropdown(
                id='period-end',
                options=[],
                className="mb-4"
            ),
            
            # Seletor de visualização
            html.Label("Tipo de Visualização:", className="mt-2"),
            dcc.Dropdown(
                id='view-selector',
                options=[
                    {'label': 'Visão Geral', 'value': 'Visão Geral'},
                    {'label': 'Desempenho por Origem', 'value': 'Desempenho por Origem'},
                    {'label': 'Conversão por Canal', 'value': 'Conversão por Canal'},
                    {'label': 'Evolução Mensal', 'value': 'Evolução Mensal'},
                    {'label': 'Top Canais', 'value': 'Top Canais'},
                    {'label': 'Eficiência de Vendas', 'value': 'Eficiência de Vendas'},
                    {'label': 'Correlação Leads-Vendas', 'value': 'Correlação Leads-Vendas'},
                    {'label': 'Dispersão Leads x Vendas', 'value': 'Dispersão Leads x Vendas'}
                ],
                value='Visão Geral',
                className="mb-4"
            ),

            html.Label("Filtrar por Origem:", className="mt-2", id='origin-filter-label'),
            dbc.DropdownMenu(
                [
                    dbc.Checklist(
                        id='origin-checklist',
                        options=[],
                        value=[],
                        labelStyle={"display": "block", "margin": "5px"},
                    )
                ],
                id='origin-dropdown',
                label="Selecione as origens...",
                className="mb-4",
                direction="down",
                toggle_style={
                    "width": "100%",
                    "text-align": "left",
                    "background": "white",
                    "border": "1px solid #ced4da"
                },
                toggleClassName="dropdown-toggle-custom",
            ),
            
            # Botão de exportar
            dbc.Button("Exportar Relatório", id="export-btn", color="secondary", className="w-100 mt-4")
        ], width=3, className="bg-light p-4"),
        
        # Área principal
        dbc.Col([
            dcc.Tabs(
                id="tabs",
                value='tab-dashboard',
                children=[
                    dcc.Tab(label='Dashboard', value='tab-dashboard', 
                            className='custom-tab', selected_className='custom-tab--selected'),
                    dcc.Tab(label='Dados Detalhados', value='tab-details',
                            className='custom-tab', selected_className='custom-tab--selected'),
                ],
                className="mb-3"
            ),
            html.Div(id='tabs-content', className="p-4")
        ], width=9)
    ]),
    
    # Armazenamento de dados
    dcc.Store(id='stored-data'),
    dcc.Store(id='stored-periods'),
    dcc.Store(id='origin-performance-data'),  
    dcc.Download(id="download-report"),
    dcc.Download(id="download-origin-excel")  
], fluid=True, className="bg-white")

# Funções auxiliares
def clean_column_name(name): # Padroniza nomes de colunas
    """Padroniza nomes de colunas"""
    name = str(name).strip().lower()
    if 'período' in name or 'periodo' in name:
        return 'periodo'
    name = re.sub(r'[^\w\s]', '_', name)
    name = re.sub(r'\s+', '_', name)
    return name

def convert_percentage(series): # Converte porcentagens para valores numéricos
    """Converte porcentagens para valores numéricos"""
    if series.dtype == 'object':
        # Tentar converter para float
        try:
            converted = series.astype(str).str.replace('%', '').str.replace(',', '.').str.strip()
            converted = pd.to_numeric(converted, errors='coerce')
            if not converted.isna().all():  # Se a conversão foi bem sucedida
                return converted / 100
        except:
            pass
    return series

def convert_numeric_columns(df): # Converte colunas numéricas conhecidas para tipo float
    """Converte colunas numéricas conhecidas para tipo float"""
    numeric_columns = ['contatos', 'aproveitados', 'vendas', 'leads', 'conversao']
    for col in numeric_columns:
        if col in df.columns:
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            except:
                pass
    return df

def parse_period(period_str): # Converte período em formato datetime para ordenação
    """Converte período em formato datetime para ordenação"""
    try:
        # Se já for um objeto datetime, retorne diretamente
        if isinstance(period_str, (datetime, pd.Timestamp)):
            return period_str
            
        # Se for um objeto de data do pandas
        if isinstance(period_str, pd.Timestamp):
            return period_str.to_pydatetime()
            
        # Converter para string se necessário
        period_str = str(period_str).strip()
        
        # Tentar formatos conhecidos
        formats = [
            "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m", "%b-%y", 
            "%b %y", "%B %Y", "%m/%d/%Y", "%Y-%m-%dT%H:%M:%S",  # Adicionado formato ISO
            "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%d %B %Y"
        ]
        
        for fmt in formats:
            try:
                return datetime.strptime(period_str, fmt)
            except ValueError:
                continue
        
        # Mapeamento de meses em português
        month_map_full = {
            'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4, 
            'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 
            'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
        }
        
        month_map_abbr = {
            'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
            'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
        }
        
        # Verificar se é apenas o nome do mês
        if period_str.lower() in month_map_full:
            # Usar ano atual se não especificado
            year = datetime.now().year
            return datetime(year, month_map_full[period_str.lower()], 1)
            
        # Tentar padrões como "Jan-23" ou "Janeiro 2023"
        match = re.match(r'(\D{3,}).*?(\d{2,4})', period_str, re.IGNORECASE)
        if match:
            month_str = match.group(1).lower()
            year_str = match.group(2)
            
            # Verificar mês completo
            if month_str in month_map_full:
                month = month_map_full[month_str]
            # Verificar abreviação
            elif month_str[:3] in month_map_abbr:
                month = month_map_abbr[month_str[:3]]
            else:
                return datetime(1900, 1, 1)
                
            year = int(year_str) if len(year_str) == 4 else 2000 + int(year_str)
            return datetime(year, month, 1)
        
        # Tentar padrão de apenas números (MM/AAAA)
        parts = re.findall(r'\d+', period_str)
        if len(parts) >= 2:
            month = int(parts[0])
            year = int(parts[1]) if len(parts[1]) == 4 else 2000 + int(parts[1])
            if 1 <= month <= 12:
                return datetime(year, month, 1)
        
        # Se nada funcionar, retornar data padrão
        return datetime(1900, 1, 1)
            
    except Exception as e:
        print(f"Erro ao analisar período '{period_str}': {str(e)}")
        return datetime(1900, 1, 1)

def format_period_display(dt): # Formata período para exibição
    """Formata período para exibição no formato 'jan/25'"""
    month_abbr = {
        1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 
        5: 'mai', 6: 'jun', 7: 'jul', 8: 'ago', 
        9: 'set', 10: 'out', 11: 'nov', 12: 'dez'
    }
    return f"{month_abbr[dt.month]}/{dt.strftime('%y')}"

def calculate_correlations(df): # Calcula correlação entre Leads e Vendas para cada origem
    """Calcula correlação entre Leads e Vendas para cada origem"""
    correlations = []
    
    for origem in df['origem'].unique():
        df_origem = df[df['origem'] == origem]
        
        # Garantir que temos dados suficientes
        if len(df_origem) >= 3 and 'contatos' in df_origem.columns and 'vendas' in df_origem.columns:
            # Remover zeros para evitar problemas
            df_clean = df_origem[(df_origem['contatos'] > 0) & (df_origem['vendas'] > 0)]
            
            if len(df_clean) >= 3:
                try:
                    corr = df_clean['contatos'].corr(df_clean['vendas'])
                    correlations.append({
                        'origem': origem,
                        'correlacao': corr,
                        'n_periodos': len(df_clean)
                    })
                except:
                    continue
    
    # Criar DataFrame e ordenar
    if correlations:
        df_corr = pd.DataFrame(correlations)
        return df_corr.sort_values('correlacao', ascending=False)
    return pd.DataFrame()

def serialize_dataframe(df): # Função para serializar DataFrames
    return {
        'columns': df.columns.tolist(),
        'data': df.values.tolist(),
        'index': df.index.tolist()
    }

def deserialize_dataframe(serialized_df): # Função  para desserializar
    return pd.DataFrame(
        data=serialized_df['data'],
        columns=serialized_df['columns']
    )

@app.callback(
    [Output('stored-data', 'data'),
     Output('stored-periods', 'data'),
     Output('period-start', 'options'),
     Output('period-end', 'options'),
     Output('city-selector', 'options'),
     Output('city-selector', 'value')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_stored_data(contents, filename): # Atualiza os dados armazenados e opções de período/cidade
    default_return = [no_update, no_update, no_update, no_update, no_update, '']
    
    if contents is None:
        return default_return
    
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    
    try:
        data = {}
        period_tuples = []
        abas_excluidas = ["Salvador", "Planilha1", "Fortaleza", "Deliverysalvador1", "Deliverysalvador4", "SSA", "Brasília", "J4ASSUNCAO"]
        
        with io.BytesIO(decoded) as file:
            with pd.ExcelFile(file) as xlsx:
                for sheet_name in xlsx.sheet_names:
                    if sheet_name in abas_excluidas:
                        continue
                    
                    try:
                        df = pd.read_excel(xlsx, sheet_name)
                        df.columns = [clean_column_name(col) for col in df.columns]
                        df = convert_numeric_columns(df)
                        
                        for col in df.columns:
                            if 'conversao' in col.lower() or '%' in col.lower():
                                df[col] = convert_percentage(df[col])
                        
                        df['cidade'] = sheet_name
                        data[sheet_name] = df
                        
                        if 'periodo' in df.columns:
                            for p in df['periodo'].dropna().unique():
                                dt = parse_period(p)
                                if dt.year > 2000:
                                    period_tuples.append((dt, p))
                                
                    except Exception as e:
                        print(f"Erro na aba {sheet_name}: {str(e)}")
                        continue
        
        if not data:
            return default_return
        
        unique_periods = {}
        for dt, p in period_tuples:
            key = dt.strftime('%Y-%m')
            if key not in unique_periods:
                unique_periods[key] = p
        
        sorted_periods = sorted(
            unique_periods.items(), 
            key=lambda x: datetime.strptime(x[0], '%Y-%m')
        )
        
        original_periods = [p for _, p in sorted_periods]
        
        period_options = []
        for p in original_periods:
            dt = parse_period(p)
            display_value = format_period_display(dt)
            period_options.append({'label': display_value, 'value': p})
        
        # REMOVER OPÇÃO "TOTAL" - APENAS ABAS DA PLANILHA
        city_options = [{'label': sheet_name, 'value': sheet_name} for sheet_name in data.keys()]
        
        serialized_data = {}
        for sheet_name, df in data.items():
            serialized_data[sheet_name] = serialize_dataframe(df)
        
        # Definir primeira cidade como padrão
        first_city = list(data.keys())[0] if data else ''
    
        return [
            serialized_data, 
            original_periods, 
            period_options, 
            period_options, 
            city_options, 
            first_city
        ]
    
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return default_return

@app.callback(
    [Output('period-start', 'value'),
     Output('period-end', 'value')],
    [Input('stored-periods', 'data')]
)
def set_default_periods(periods): # Define os períodos padrão para os seletores
    if not periods:
        return None, None
    return periods[0], periods[-1]

@app.callback(
    Output('tabs-content', 'children'),
    [Input('tabs', 'value'),
     Input('city-selector', 'value'),
     Input('period-start', 'value'),
     Input('period-end', 'value'),
     Input('view-selector', 'value'),
     Input('stored-data', 'data'),
     Input('stored-periods', 'data'),
     Input('origin-checklist', 'value')]
)
def render_tabs_content(tab, city, period_start, period_end, view, data, periods, selected_origins): # Renderiza o conteúdo das abas
    try:
        if data is None:
            return dbc.Alert("Por favor, carregue um arquivo Excel para começar.", color="info")
        
        # Obter os dados com o novo filtro de origens
        df = get_current_data(data, city, period_start, period_end, selected_origins)
        
        if df is None or df.empty:
            return dbc.Alert("Nenhum dado encontrado para os filtros selecionados.", color="warning")
        
        if tab == 'tab-details':
            return render_details_tab(df)
        else:
            return render_dashboard(view, df)
    
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"Erro em render_tabs_content: {str(e)}")
        print(error_trace)
        return dbc.Alert(f"Erro ao renderizar conteúdo: {str(e)}", color="danger")

def render_dashboard(view, df): # Renderiza o dashboard com base na visualização selecionada
    try:
        if view == "Visão Geral":
            return render_summary_view(df)
        elif view == "Desempenho por Origem":
            return render_origin_performance(df)
        elif view == "Conversão por Canal":
            return render_conversion_by_channel(df)
        elif view == "Evolução Mensal":
            return render_monthly_trend(df)
        elif view == "Top Canais":
            return render_top_channels(df)
        elif view == "Eficiência de Vendas":
            return render_sales_efficiency(df)
        elif view == "Correlação Leads-Vendas":  
            return render_correlation_view(df)
        elif view == "Dispersão Leads x Vendas":
            return render_scatter_plots(df)
        
        return dbc.Alert("Visualização não reconhecida", color="warning")
    except Exception as e:
        print(f"Erro em render_dashboard: {str(e)}")
        return dbc.Alert(f"Erro ao renderizar dashboard: {str(e)}", color="danger")

def render_details_tab(df): # Renderiza a aba de detalhes com uma tabela
    try:
        return dash_table.DataTable(
            data=df.to_dict('records'),
            columns=[{'name': i, 'id': i} for i in df.columns],
            page_size=20,
            style_table={'overflowX': 'auto'},
            style_header={
                'backgroundColor': '#e9ecef',
                'color': '#212529',
                'fontWeight': 'bold',
                'border': '1px solid #dee2e6'
            },
            style_cell={
                'backgroundColor': 'white',
                'color': '#212529',
                'padding': '10px',
                'border': '1px solid #dee2e6'
            },
            style_data_conditional=[
                {
                    'if': {'row_index': 'odd'},
                    'backgroundColor': '#f8f9fa'
                }
            ]
        )
    except Exception as e:
        print(f"Erro em render_details_tab: {str(e)}")
        return dbc.Alert(f"Erro ao exibir dados detalhados: {str(e)}", color="danger")

def get_current_data(data, city, period_start, period_end, selected_origins=None): # Obtém os dados atuais filtrados por cidade, período e origens selecionadas
    if data is None or city not in data:
        return None
    
    try:
        # Desserializar apenas a aba selecionada
        df = deserialize_dataframe(data[city])
        df = convert_numeric_columns(df)

        # Filtrar por período se disponível
        if 'periodo' in df.columns and period_start and period_end:
            start_dt = parse_period(period_start)
            end_dt = parse_period(period_end)
            
            # Criar coluna de data se necessário
            if 'periodo_dt' not in df.columns:
                df['periodo_dt'] = df['periodo'].apply(parse_period)
            
            df = df[(df['periodo_dt'] >= start_dt) & (df['periodo_dt'] <= end_dt)]
        
        # Filtrar linhas com "Total" na origem
        if 'origem' in df.columns:
            # Converter para minúsculas e remover espaços
            origens = df['origem'].astype(str).str.lower().str.strip()
            
            # Criar máscara para excluir qualquer variação de 'total'
            mask = ~(
                origens.str.contains('total') |
                origens.str.contains('geral') |
                origens.str.contains('consolidado')
            )
            
            # Aplicar filtro
            df = df[mask]
        
        # Filtro por origem
        if selected_origins and 'origem' in df.columns:
            df = df[df['origem'].isin(selected_origins)]

        return df
    
    except Exception as e:
        print(f"Erro em get_current_data: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

@app.callback(
    Output('origin-checklist', 'options'),
    Output('origin-checklist', 'value'),
    [Input('stored-data', 'data'),
     Input('city-selector', 'value'),
     Input('period-start', 'value'),
     Input('period-end', 'value')],
    [State('origin-checklist', 'value')]
)
def update_origin_checklist(data, city, period_start, period_end, current_selection): # Atualiza a lista de origens disponíveis no checklist
    if not data or not city:
        return [], []
    
    df = get_current_data(data, city, period_start, period_end, None)
    
    if df is None or 'origem' not in df.columns:
        return [], current_selection
    
    # Obter origens únicas e ordenar alfabeticamente
    origins = sorted(df['origem'].unique().tolist())
    
    options = [{"label": orig, "value": orig} for orig in origins if orig]
    
    # Selecionar todas as origens por padrão
    if not current_selection:
        selected_values = origins
    else:
        # Manter seleção atual se ainda for válida
        selected_values = [v for v in current_selection if v in origins]
    
    return options, selected_values

@app.callback(
    Output('origin-dropdown', 'label'),
    [Input('origin-checklist', 'value')],
    [State('origin-checklist', 'options')]
)
def update_dropdown_label(selected_values, options): # Atualiza o rótulo do dropdown com base nas origens selecionadas
    if not selected_values:
        return "Selecione as origens..."
    
    # Verificar se todas as origens estão selecionadas
    all_options = {opt['value'] for opt in options}
    if set(selected_values) == all_options:
        return "Todas as origens"
    
    labels = {opt['value']: opt['label'] for opt in options}
    selected_labels = [labels[v] for v in selected_values if v in labels]
    
    if len(selected_labels) > 2:
        label = f"{len(selected_labels)} origens selecionadas"
    else:
        label = ", ".join(selected_labels)
    
    return label

@app.callback(
    Output('origin-filter-label', 'style'),
    Output('origin-dropdown', 'style'),
    [Input('view-selector', 'value')]
)
def toggle_origin_filter(view): # Mostra ou esconde o filtro de origem com base na visualização selecionada
    show_style = {'display': 'block'}
    hide_style = {'display': 'none'}
    
    if view in ['Visão Geral', 'Desempenho por Origem', 'Conversão por Canal', 'Top Canais', 'Eficiência de Vendas', 'Correlação Leads-Vendas', 'Dispersão Leads x Vendas']:
        return show_style, show_style
    return hide_style, hide_style

def prepare_origin_export_data(df): # Prepara os dados para exportação em Excel
    """Prepara os dados para exportação em Excel"""
    if 'origem' not in df.columns:
        return None
        
    grouped = df.groupby('origem').agg({
        'contatos': 'sum',
        'aproveitados': 'sum',
        'vendas': 'sum'
    }).reset_index()
    
    grouped['%_aproveitamento'] = grouped['aproveitados'] / grouped['contatos']
    grouped['taxa_conversao'] = grouped['vendas'] / grouped['contatos']
    grouped['conversao_ap'] = grouped['vendas'] / grouped['aproveitados']
    grouped['lead_por_venda'] = grouped['contatos'] / grouped['vendas']
    
    # Ordenar por vendas (decrescente)
    return grouped.sort_values('vendas', ascending=False)

@app.callback(
    Output('download-origin-excel', 'data'),
    Input('export-origin-excel-btn', 'n_clicks'),
    State('origin-performance-data', 'data'),
    prevent_initial_call=True
)
def export_origin_excel(n_clicks, data):
    if n_clicks is None or not data:
        return dash.no_update
    
    try:
        # Criar DataFrame a partir dos dados armazenados
        df = pd.DataFrame(data)
        
        # Criar arquivo Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Desempenho_por_Origem', index=False)
            
            # Acessar a planilha e o workbook
            workbook = writer.book
            worksheet = writer.sheets['Desempenho_por_Origem']
            
            # Formatar cabeçalhos
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#2c3e50',
                'border': 1,
                'color': 'white'
            })
            
            # Aplicar formatação aos cabeçalhos
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Ajustar largura das colunas
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
        
        output.seek(0)
        return dcc.send_bytes(
            output.getvalue(), 
            filename=f"desempenho_por_origem_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        
    except Exception as e:
        print(f"Erro ao exportar Excel: {str(e)}")
        return dash.no_update
    
@app.callback(
    Output('download-report', 'data'),
    [Input('export-btn', 'n_clicks')],
    [State('city-selector', 'value'),
     State('period-start', 'value'),
     State('period-end', 'value'),
     State('view-selector', 'value'),
     State('stored-data', 'data'),
     State('stored-periods', 'data'),
     State('origin-checklist', 'value')],
    prevent_initial_call=True
)
def export_report(n_clicks, city, period_start, period_end, view, data, periods, selected_origins): # Exporta o relatório em PDF
    if n_clicks is None or not city or not period_start or not period_end or not data:
        return dash.no_update
    
    df = get_current_data(data, city, period_start, period_end, selected_origins)
    if df is None or df.empty:
        return dash.no_update
    
    try:
        # Criar PDF em memória
        pdf_buffer = io.BytesIO()
        generate_pdf_report(df, city, period_start, period_end, view, pdf_buffer)
        
        # Retornar para download
        filename = f"Relatorio_{city}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        return dcc.send_bytes(pdf_buffer.getvalue(), filename=filename)
    
    except Exception as e:
        print(f"Erro ao gerar relatório: {str(e)}")
        return dash.no_update

def generate_pdf_report(df, city, period_start, period_end, view, buffer): # Gera o relatório em PDF com base nos dados filtrados
    # Configurações do documento
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )
    
    styles = getSampleStyleSheet()
    
    # Definir estilos personalizados apenas se não existirem
    if not hasattr(styles, 'TitleStyle'):
        styles.add(ParagraphStyle(
            name='TitleStyle',
            fontSize=18,
            alignment=TA_CENTER,
            spaceAfter=12
        ))
    
    if not hasattr(styles, 'SubtitleStyle'):
        styles.add(ParagraphStyle(
            name='SubtitleStyle',
            fontSize=12,
            alignment=TA_CENTER,
            textColor=colors.grey,
            spaceAfter=20
        ))
    
    if not hasattr(styles, 'SectionStyle'):
        styles.add(ParagraphStyle(
            name='SectionStyle',
            fontSize=14,
            spaceBefore=20,
            spaceAfter=10
        ))
    
    if not hasattr(styles, 'BodyStyle'):
        styles.add(ParagraphStyle(
            name='BodyStyle',
            fontSize=10,
            alignment=TA_JUSTIFY,
            leading=14
        ))
    
    if not hasattr(styles, 'MetricStyle'):
        styles.add(ParagraphStyle(
            name='MetricStyle',
            fontSize=14,
            textColor=colors.darkblue,
            spaceAfter=5
        ))
    
    if not hasattr(styles, 'FooterStyle'):
        styles.add(ParagraphStyle(
            name='FooterStyle',
            fontSize=8,
            textColor=colors.grey,
            alignment=TA_CENTER
        ))
    
    elements = []
    
    # Cabeçalho
    elements.append(Paragraph("Relatório de Performance de Leads", styles['TitleStyle']))
    
    # Converter e formatar as datas
    start_dt = parse_period(period_start)
    end_dt = parse_period(period_end)
    formatted_start = format_period_display(start_dt)
    formatted_end = format_period_display(end_dt)
    
    elements.append(Paragraph(f"{city} | {formatted_start} - {formatted_end}", styles['SubtitleStyle']))
    elements.append(Spacer(1, 12))
    
    # Seção: Visão Geral
    elements.append(Paragraph("Visão Geral", styles['SectionStyle']))
    
    # Métricas da visão geral
    try:
        total_contacts = df['contatos'].sum() if 'contatos' in df.columns else 0
        total_leads = df['aproveitados'].sum() if 'aproveitados' in df.columns else 0
        total_sales = df['vendas'].sum() if 'vendas' in df.columns else 0
        
        conversion_rate = total_sales / total_contacts if total_contacts > 0 else 0
        lead_conversion_rate = total_sales / total_leads if total_leads > 0 else 0
        
        metrics_data = [
            ["Métrica", "Valor", "Insight"],
            ["Total de Contatos", f"{total_contacts:,.0f}", "Volume total de oportunidades geradas"],
            ["Leads Aproveitados", f"{total_leads:,.0f}", f"({total_leads/total_contacts:.1%} dos contatos)" if total_contacts > 0 else ""],
            ["Vendas Fechadas", f"{total_sales:,.0f}", f"({conversion_rate:.1%} de conversão geral)" if total_contacts > 0 else ""],
            ["Conversão de Leads Aproveitados", f"{lead_conversion_rate:.1%}" if total_leads > 0 else "N/A", "Eficiência no aproveitamento de oportunidades"]
        ]
        
        metrics_table = Table(metrics_data)
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ]))
        elements.append(metrics_table)
        elements.append(Spacer(1, 20))
    except:
        pass

    # Gráficos da visão geral
    fig_paths = []
    
    # 1. Distribuição de contatos por origem
    if 'origem' in df.columns and 'contatos' in df.columns and not df.empty:
        try:
            origin_counts = df.groupby('origem')['contatos'].sum()
            if len(origin_counts) > 5:
                top_origins = origin_counts.nlargest(5)
                other = origin_counts.sum() - top_origins.sum()
                top_origins = pd.concat([top_origins, pd.Series({'Outros': other})])
            else:
                top_origins = origin_counts
            
            fig = px.pie(
                top_origins, 
                values=top_origins.values, 
                names=top_origins.index,
                title="Distribuição por Origem (Top 5)",
                hole=0.3
            )
            fig.update_layout(showlegend=True)
            
            fig_path = tempfile.mktemp(suffix='.png')
            fig.write_image(fig_path)
            fig_paths.append(("Distribuição de Contatos por Origem", fig_path))
        except:
            pass

    # 2. Top canais por conversão
    if 'origem' in df.columns and 'contatos' in df.columns and 'vendas' in df.columns and not df.empty:
        try:
            channel_efficiency = df.groupby('origem').agg({
                'contatos': 'sum',
                'vendas': 'sum'
            })
            channel_efficiency['taxa_conversao'] = channel_efficiency['vendas'] / channel_efficiency['contatos']
            channel_efficiency = channel_efficiency[channel_efficiency['contatos'] >= 50]
            
            if not channel_efficiency.empty:
                top_conversion = channel_efficiency.nlargest(5, 'taxa_conversao').sort_values('taxa_conversao', ascending=True)
                
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    y=top_conversion.index,
                    x=top_conversion['taxa_conversao'],
                    orientation='h',
                    marker_color='#3498db'
                ))
                fig.update_layout(
                    title='Top 5 Canais - Conversão Total',
                    xaxis_title="Taxa de Conversão",
                    yaxis_title="Canal",
                    xaxis_tickformat=".0%",
                    showlegend=False
                )
                
                fig_path = tempfile.mktemp(suffix='.png')
                fig.write_image(fig_path)
                fig_paths.append(("Top Canais por Conversão", fig_path))
        except:
            pass

    # Adicionar gráficos da visão geral ao PDF
    for title, path in fig_paths:
        try:
            elements.append(Paragraph(title, styles['BodyStyle']))
            elements.append(Spacer(1, 5))
            elements.append(Image(path, width=400, height=250))
            elements.append(Spacer(1, 15))
        except:
            pass

    # Seção: Evolução Mensal
    elements.append(PageBreak())
    elements.append(Paragraph("Evolução Mensal", styles['SectionStyle']))
    elements.append(Paragraph("Análise: Acompanhe a evolução dos principais indicadores ao longo do tempo. Tendências de crescimento em contatos e vendas indicam eficácia nas estratégias. Quedas consistentes podem sinalizar problemas operacionais ou de mercado.", styles['BodyStyle']))
    
    fig_paths = []
    
    # 1. Gráfico de volume
    if 'periodo_dt' in df.columns and not df.empty:
        try:
            monthly = df.groupby('periodo_dt').agg({
                'contatos': 'sum',
                'aproveitados': 'sum',
                'vendas': 'sum'
            }).sort_index()
            
            monthly['periodo_formatado'] = monthly.index.map(format_period_display)
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=monthly['periodo_formatado'], 
                y=monthly['contatos'], 
                name='Contatos',
                line=dict(color='#3498db')
            ))
            fig.add_trace(go.Scatter(
                x=monthly['periodo_formatado'], 
                y=monthly['aproveitados'], 
                name='Aproveitados',
                line=dict(color='#2ecc71')
            ))
            fig.add_trace(go.Scatter(
                x=monthly['periodo_formatado'], 
                y=monthly['vendas'], 
                name='Vendas',
                line=dict(color='#e74c3c')
            ))
            fig.update_layout(
                title="Evolução Mensal - Volume",
                xaxis_title="Período",
                yaxis_title="Quantidade",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            
            fig_path = tempfile.mktemp(suffix='.png')
            fig.write_image(fig_path)
            fig_paths.append(("Volume de Contatos, Leads e Vendas", fig_path))
        except:
            pass

    # 2. Gráfico de taxas
    if 'periodo_dt' in df.columns and not df.empty:
        try:
            monthly = df.groupby('periodo_dt').agg({
                'contatos': 'sum',
                'aproveitados': 'sum',
                'vendas': 'sum'
            }).sort_index()
            
            monthly['taxa_conversao'] = monthly['vendas'] / monthly['contatos']
            monthly['conversao_ap'] = monthly['vendas'] / monthly['aproveitados']
            monthly['periodo_formatado'] = monthly.index.map(format_period_display)
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=monthly['periodo_formatado'], 
                y=monthly['taxa_conversao'], 
                name='Conversão Total',
                line=dict(color='#f39c12')
            ))
            fig.add_trace(go.Scatter(
                x=monthly['periodo_formatado'], 
                y=monthly['conversao_ap'], 
                name='Conversão Aproveitados',
                line=dict(color='#9b59b6')
            ))
            fig.update_layout(
                title="Evolução Mensal - Taxas de Conversão",
                xaxis_title="Período",
                yaxis_title="Taxa",
                yaxis_tickformat=".0%",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            
            fig_path = tempfile.mktemp(suffix='.png')
            fig.write_image(fig_path)
            fig_paths.append(("Taxas de Conversão", fig_path))
        except:
            pass

    # Adicionar gráficos de evolução mensal ao PDF
    for title, path in fig_paths:
        try:
            elements.append(Paragraph(title, styles['BodyStyle']))
            elements.append(Spacer(1, 5))
            elements.append(Image(path, width=400, height=250))
            elements.append(Spacer(1, 15))
        except:
            pass

    # Seção: Top Canais
    elements.append(PageBreak())
    elements.append(Paragraph("Top Canais", styles['SectionStyle']))
    elements.append(Paragraph("Análise: Identifique os canais com melhor desempenho. Canais com alta conversão representam oportunidades de investimento. Canais com baixa conversão podem precisar de otimização ou realocação de recursos.", styles['BodyStyle']))
    
    fig_paths = []
    
    # 1. Top por conversão total
    if 'origem' in df.columns and 'contatos' in df.columns and 'vendas' in df.columns and not df.empty:
        try:
            channel_efficiency = df.groupby('origem').agg({
                'contatos': 'sum',
                'vendas': 'sum'
            })
            channel_efficiency['taxa_conversao'] = channel_efficiency['vendas'] / channel_efficiency['contatos']
            channel_efficiency = channel_efficiency[channel_efficiency['contatos'] >= 50]
            
            if not channel_efficiency.empty:
                top_conversion = channel_efficiency.nlargest(5, 'taxa_conversao').sort_values('taxa_conversao', ascending=True)
                
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    y=top_conversion.index,
                    x=top_conversion['taxa_conversao'],
                    orientation='h',
                    marker_color='#3498db'
                ))
                fig.update_layout(
                    title='Top 5 Canais - Conversão Total',
                    xaxis_title="Taxa de Conversão",
                    yaxis_title="Canal",
                    xaxis_tickformat=".0%",
                    showlegend=False
                )
                
                fig_path = tempfile.mktemp(suffix='.png')
                fig.write_image(fig_path)
                fig_paths.append(("Conversão Total (Vendas/Contatos)", fig_path))
        except:
            pass

    # 2. Top por conversão de leads
    if 'origem' in df.columns and 'aproveitados' in df.columns and 'vendas' in df.columns and not df.empty:
        try:
            channel_efficiency = df.groupby('origem').agg({
                'aproveitados': 'sum',
                'vendas': 'sum'
            })
            channel_efficiency['conversao_ap'] = channel_efficiency['vendas'] / channel_efficiency['aproveitados']
            channel_efficiency = channel_efficiency[channel_efficiency['aproveitados'] >= 30]
            
            if not channel_efficiency.empty:
                top_conversion = channel_efficiency.nlargest(5, 'conversao_ap').sort_values('conversao_ap', ascending=True)
                
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    y=top_conversion.index,
                    x=top_conversion['conversao_ap'],
                    orientation='h',
                    marker_color='#2ecc71'
                ))
                fig.update_layout(
                    title='Top 5 Canais - Conversão de Leads',
                    xaxis_title="Taxa de Conversão",
                    yaxis_title="Canal",
                    xaxis_tickformat=".0%",
                    showlegend=False
                )
                
                fig_path = tempfile.mktemp(suffix='.png')
                fig.write_image(fig_path)
                fig_paths.append(("Conversão de Leads (Vendas/Aproveitados)", fig_path))
        except:
            pass

    # Adicionar gráficos de top canais ao PDF
    for title, path in fig_paths:
        try:
            elements.append(Paragraph(title, styles['BodyStyle']))
            elements.append(Spacer(1, 5))
            elements.append(Image(path, width=400, height=250))
            elements.append(Spacer(1, 15))
        except:
            pass

    # Seção: Eficiência de Vendas
    elements.append(PageBreak())
    elements.append(Paragraph("Eficiência de Vendas", styles['SectionStyle']))
    elements.append(Paragraph("Análise: Mede a porcentagem de vendas por lead. Quanto maior o valor, melhor será para investir", styles['BodyStyle']))
    
    if 'origem' in df.columns and 'contatos' in df.columns and 'vendas' in df.columns and not df.empty:
        try:
            grouped = df.groupby('origem').agg({
                'contatos': 'sum',
                'vendas': 'sum'
            }).reset_index()
            
            grouped = grouped[grouped['vendas'] > 0]
            grouped['eficiencia'] = (grouped['vendas'] / grouped['contatos']) * 100
            filtered = grouped[grouped['vendas'] >= 10]
            
            if not filtered.empty:
                top_15 = filtered.nlargest(15, 'eficiencia')
                top_15 = top_15.sort_values('eficiencia', ascending=True)
                
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    y=top_15['origem'],
                    x=top_15['eficiencia'],
                    orientation='h',
                    marker_color='#3498db'
                ))
                fig.update_layout(
                    title="Eficiência de Vendas (Vendas por Lead - Top 15)",
                    xaxis_title="Eficiência (%)",
                    yaxis_title="Origem",
                    showlegend=False
                )
                
                fig_path = tempfile.mktemp(suffix='.png')
                fig.write_image(fig_path)
                elements.append(Image(fig_path, width=500, height=350))
                elements.append(Spacer(1, 15))
        except:
            pass

    # Seção: Correlação Leads-Vendas
    elements.append(PageBreak())
    elements.append(Paragraph("Correlação Leads-Vendas", styles['SectionStyle']))
    elements.append(Paragraph("Análise: Correlação mede a relação entre leads e vendas. Valores próximos a 1 indicam que o aumento de leads acompanha o aumento de vendas. Valores próximos a -1 indicam relação inversa. Valores próximos a 0 indicam pouca relação entre as variáveis.", styles['BodyStyle']))
    elements.append(Paragraph("Interpretação: Correlação > 0.7 = Forte relação positiva | 0.3-0.7 = Relação moderada | < 0.3 = Fraca relação", styles['BodyStyle']))
    
    try:
        df_corr = calculate_correlations(df)
        
        if not df_corr.empty:
            # Melhores correlações
            top_15 = df_corr.head(15).sort_values('correlacao', ascending=False)
            
            fig_top = go.Figure()
            fig_top.add_trace(go.Bar(
                x=top_15['correlacao'],
                y=top_15['origem'],
                orientation='h',
                marker_color='#2ecc71'
            ))
            fig_top.update_layout(
                title='TOP 15 - Melhores Correlações (Leads x Vendas)',
                xaxis_title="Coeficiente de Correlação",
                yaxis_title="Canal",
                xaxis=dict(range=[-1.1, 1.1]),
                showlegend=False
            )
            
            fig_top_path = tempfile.mktemp(suffix='.png')
            fig_top.write_image(fig_top_path)
            elements.append(Image(fig_top_path, width=500, height=400))
            elements.append(Spacer(1, 15))
            
            # Piores correlações
            bottom_15 = df_corr.tail(15).sort_values('correlacao', ascending=True)
            
            fig_bottom = go.Figure()
            fig_bottom.add_trace(go.Bar(
                x=bottom_15['correlacao'],
                y=bottom_15['origem'],
                orientation='h',
                marker_color='#e74c3c'
            ))
            fig_bottom.update_layout(
                title='TOP 15 - Piores Correlações (Leads x Vendas)',
                xaxis_title="Coeficiente de Correlação",
                yaxis_title="Canal",
                xaxis=dict(range=[-1.1, 1.1]),
                showlegend=False
            )
            
            fig_bottom_path = tempfile.mktemp(suffix='.png')
            fig_bottom.write_image(fig_bottom_path)
            elements.append(Image(fig_bottom_path, width=500, height=400))
    except:
        pass

    # Seção: Dispersão Leads x Vendas
    elements.append(PageBreak())
    elements.append(Paragraph("Dispersão Leads x Vendas", styles['SectionStyle']))
    elements.append(Paragraph("Análise: Mostra a relação entre volume de leads e vendas geradas. Canais no canto superior direito (muitos leads e vendas) são os mais eficientes. Canais com muitos leads e poucas vendas precisam de otimização.", styles['BodyStyle']))

    try:
        # Agregar dados por origem
        df_agg = df.groupby('origem', as_index=False).agg({
            'contatos': 'sum',
            'aproveitados': 'sum',
            'vendas': 'sum'
        })
        
        df_agg = df_agg[
            (df_agg['contatos'] > 0) & 
            (df_agg['vendas'] > 0) & 
            (df_agg['aproveitados'] > 0)
        ]
        
        if not df_agg.empty:
            # ADIÇÃO: Selecionar top 8 origens por contatos
            top_8_contatos = df_agg.nlargest(8, 'contatos')
            
            # Gráfico 1: Leads vs Vendas
            fig1 = px.scatter(
                df_agg,
                x='contatos',
                y='vendas',
                size='vendas',
                color='origem',
                title='Relação Leads vs Vendas por Canal',
                labels={'contatos': 'Total de Leads', 'vendas': 'Total de Vendas'}
            )
            fig1.update_layout(showlegend=False)
            
            # ADIÇÃO: Adicionar rótulos para top 8
            for i, row in top_8_contatos.iterrows():
                fig1.add_annotation(
                    x=row['contatos'],
                    y=row['vendas'],
                    text=row['origem'],
                    showarrow=True,
                    arrowhead=1,
                    ax=0,
                    ay=-30
                )
            
            fig1_path = tempfile.mktemp(suffix='.png')
            fig1.write_image(fig1_path)
            elements.append(Image(fig1_path, width=500, height=400))
            elements.append(Spacer(1, 15))
            
            # ADIÇÃO: Selecionar top 8 origens por aproveitados
            top_8_aproveitados = df_agg.nlargest(8, 'aproveitados')
            
            # Gráfico 2: Leads Aproveitados vs Vendas
            fig2 = px.scatter(
                df_agg,
                x='aproveitados',
                y='vendas',
                size='vendas',
                color='origem',
                title='Relação Leads Aproveitados vs Vendas por Canal',
                labels={'aproveitados': 'Leads Aproveitados', 'vendas': 'Total de Vendas'}
            )
            fig2.update_layout(showlegend=False)
            
            # ADIÇÃO: Adicionar rótulos para top 8
            for i, row in top_8_aproveitados.iterrows():
                fig2.add_annotation(
                    x=row['aproveitados'],
                    y=row['vendas'],
                    text=row['origem'],
                    showarrow=True,
                    arrowhead=1,
                    ax=0,
                    ay=-40
                )
            
            fig2_path = tempfile.mktemp(suffix='.png')
            fig2.write_image(fig2_path)
            elements.append(Image(fig2_path, width=500, height=400))
    except:
        pass

    # Rodapé
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"Relatório gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['FooterStyle']))
    elements.append(Paragraph("Sistema de Análise de Leads | Dados Confidenciais", styles['FooterStyle']))
    
    # Construir o PDF
    doc.build(elements)
    
    # Remover arquivos temporários das imagens
    for _, path in fig_paths:
        try:
            os.remove(path)
        except:
            pass

# Funções de renderização para cada visualização
def render_summary_view(df): # Visão Geral
    try:
        # Verificar se as colunas necessárias existem
        required_columns = ['contatos', 'aproveitados', 'vendas']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            return dbc.Alert(
                f"Colunas obrigatórias ausentes: {', '.join(missing_columns)}",
                color="danger"
            )
        
        # Calcular métricas
        total_contacts = df['contatos'].sum()
        total_leads = df['aproveitados'].sum()
        total_sales = df['vendas'].sum()
        
        conversion_rate = total_sales / total_contacts if total_contacts > 0 else 0
        lead_conversion_rate = total_sales / total_leads if total_leads > 0 else 0
        lead_per_sale = total_contacts / total_sales if total_sales > 0 else 0
        
        metrics = [
            ("Total de Contatos", total_contacts, "{:,.0f}"),
            ("Leads Aproveitados", total_leads, "{:,.0f}"),
            ("Vendas Fechadas", total_sales, "{:,.0f}"),
            ("Taxa de Conversão", conversion_rate, "{:.1%}"),
            ("Conversão de Leads", lead_conversion_rate, "{:.1%}"),
            ("Lead por Venda", lead_per_sale, "{:,.1f}")
        ]
        
        # Cards de métricas
        metric_cards = []
        for title, value, fmt in metrics:
            card = dbc.Card(
                dbc.CardBody([
                    html.H5(title, className="card-title"),
                    html.H4(fmt.format(value), className="card-text")
                ]),
                className="text-center m-2",
                style={'width': '18rem', 'height': '10rem'}
            )
            metric_cards.append(card)
    
        # Gráficos
        graphs = []
        
        # Gráfico 1: Distribuição de origens
        if 'origem' in df.columns and 'contatos' in df.columns:
            if df['contatos'].dtype == 'object':
                df['contatos'] = pd.to_numeric(df['contatos'], errors='coerce')
            
            origin_counts = df.groupby('origem')['contatos'].sum()
            
            if len(origin_counts) > 5:
                top_origins = origin_counts.nlargest(5)
                other = origin_counts.sum() - top_origins.sum()
                top_origins = pd.concat([top_origins, pd.Series({'Outros': other})])
            else:
                top_origins = origin_counts
            
            fig1 = px.pie(
                top_origins, 
                values=top_origins.values, 
                names=top_origins.index,
                title="Distribuição por Origem (Top 5)",
                hole=0.3
            )
            fig1.update_layout(paper_bgcolor="#ffffff", plot_bgcolor="#999CA0", font_color='#212529')
            graphs.append(dcc.Graph(figure=fig1, className="mb-4"))
        
        # Gráfico 2: Top 5 Canais por Conversão Total
        if 'origem' in df.columns and 'contatos' in df.columns and 'vendas' in df.columns:
            channel_efficiency = df.groupby('origem').agg({
                'contatos': 'sum',
                'vendas': 'sum'
            })
            channel_efficiency['taxa_conversao'] = channel_efficiency['vendas'] / channel_efficiency['contatos']
            channel_efficiency = channel_efficiency[channel_efficiency['contatos'] >= 50]
            
            if not channel_efficiency.empty:
                top_conversion = channel_efficiency.nlargest(5, 'taxa_conversao').sort_values('taxa_conversao', ascending=True)
                
                fig2 = go.Figure()
                fig2.add_trace(go.Bar(
                    y=top_conversion.index,
                    x=top_conversion['taxa_conversao'],
                    orientation='h',
                    text=[f'{v:.1%}' for v in top_conversion['taxa_conversao']],
                    textposition='auto',
                    marker_color='#3498db'
                ))
                fig2.update_layout(
                    title="Top 5 Canais - Conversão Total",
                    xaxis_title="Taxa de Conversão",
                    yaxis_title="Canal",
                    xaxis_tickformat=".0%",
                    paper_bgcolor='white',
                    plot_bgcolor="#999CA0",
                    font_color='#212529'
                )
                graphs.append(dcc.Graph(figure=fig2, className="mb-4"))
        
        return html.Div([
                dbc.Row([dbc.Col(card) for card in metric_cards], className="mb-4"),
            dbc.Row([
                dbc.Col(graphs[0], width=6),
                dbc.Col(graphs[1], width=6)
            ]) if len(graphs) == 2 else dbc.Row(dbc.Col(graphs[0]))
        ])

    except Exception as e:
        print(f"Erro em render_summary_view: {str(e)}")
        return dbc.Alert(f"Erro ao renderizar visão geral: {str(e)}", color="danger")
    
def render_origin_performance(df): # Desempenho por origem
    try:
        if 'origem' not in df.columns:
            return dbc.Alert("Dados de origem não disponíveis", color="danger")
        
        grouped = prepare_origin_export_data(df)  # Usar a nova função
        
        columns = [
            {'name': 'Origem', 'id': 'origem'},
            {'name': 'Contatos', 'id': 'contatos', 'type': 'numeric', 'format': {'specifier': ',.0f'}},
            {'name': 'Aproveitados', 'id': 'aproveitados', 'type': 'numeric', 'format': {'specifier': ',.0f'}},
            {'name': 'Vendas', 'id': 'vendas', 'type': 'numeric', 'format': {'specifier': ',.0f'}},
            {'name': '% Aproveit.', 'id': '%_aproveitamento', 'type': 'numeric', 'format': {'specifier': '.1%'}},
            {'name': 'Conv. Total', 'id': 'taxa_conversao', 'type': 'numeric', 'format': {'specifier': '.1%'}},
            {'name': 'Conv. Aproveitados', 'id': 'conversao_ap', 'type': 'numeric', 'format': {'specifier': '.1%'}},
            {'name': 'Lead/Venda', 'id': 'lead_por_venda', 'type': 'numeric', 'format': {'specifier': ',.1f'}}
        ]
        
        # RETORNO MODIFICADO: Agora inclui o botão de exportação
        return html.Div([
            dbc.Button(
                "Exportar para Excel", 
                id="export-origin-excel-btn", 
                color="success", 
                className="mb-3"
            ),
            dash_table.DataTable(
                data=grouped.to_dict('records'),
                columns=columns,
                page_size=10,
                style_table={'overflowX': 'auto'},
                style_header={
                    'backgroundColor': '#e9ecef',
                    'color': '#212529',
                    'fontWeight': 'bold',
                    'border': '1px solid #dee2e6'
                },
                style_cell={
                    'backgroundColor': 'white',
                    'color': '#212529',
                    'padding': '10px',
                    'border': '1px solid #dee2e6'
                },
                style_data_conditional=[
                    {
                        'if': {'row_index': 'odd'},
                        'backgroundColor': '#f8f9fa'
                    }
                ]
            ),
            dcc.Store(
                id='origin-performance-data',
                data=grouped.to_dict('records')
            )
        ])

    except Exception as e:
        print(f"Erro em render_origin_performance: {str(e)}")
        return dbc.Alert(f"Erro ao exibir desempenho por origem: {str(e)}", color="danger")
    
def render_conversion_by_channel(df): # Conversão por Canal
    try: 
        if 'origem' not in df.columns:
            return dbc.Alert("Dados de origem não disponíveis", color="danger")
        
        grouped = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum',
            'aproveitados': 'sum'
        }).reset_index()
        
        grouped['taxa_conversao'] = grouped['vendas'] / grouped['contatos']
        grouped['conversao_ap'] = grouped['vendas'] / grouped['aproveitados']
        grouped = grouped[grouped['contatos'] >= 100]
        
        if grouped.empty:
            return dbc.Alert("Sem dados suficientes para análise", color="warning")
        
        # Gráfico 1: Conversão total
        fig1 = px.bar(
            grouped.sort_values('taxa_conversao', ascending=True),
            y='origem',
            x='taxa_conversao',
            orientation='h',
            title="Conversão Total por Origem",
            text='taxa_conversao',
            labels={'taxa_conversao': 'Taxa de Conversão', 'origem': 'Origem'}
        )
        fig1.update_traces(texttemplate='%{text:.1%}', textposition='auto')
        fig1.update_layout(
            xaxis_tickformat=".0%",
            paper_bgcolor='white', plot_bgcolor="#C8CCD1", font_color='#212529'
        )
        
        # Gráfico 2: Conversão de leads
        fig2 = px.bar(
            grouped.sort_values('conversao_ap', ascending=True),
            y='origem',
            x='conversao_ap',
            orientation='h',
            title="Conversão de Leads por Origem",
            text='conversao_ap',
            labels={'conversao_ap': 'Taxa de Conversão', 'origem': 'Origem'}
        )
        fig2.update_traces(texttemplate='%{text:.1%}', textposition='auto')
        fig2.update_layout(
            xaxis_tickformat=".0%",
            paper_bgcolor='white', plot_bgcolor='#C8CCD1', font_color='#212529'
        )
        
        return dbc.Row([
            dbc.Col(dcc.Graph(figure=fig1)), 
            dbc.Col(dcc.Graph(figure=fig2))
        ])

    except Exception as e:
        print(f"Erro em render_conversion_by_channel: {str(e)}")
        return dbc.Alert(f"Erro ao exibir conversão por canal: {str(e)}", color="danger")
    
def render_monthly_trend(df): # Evolução Mensal
    try:
        if 'periodo_dt' not in df.columns:
            return dbc.Alert("Dados temporais não disponíveis", color="danger")
        
        monthly = df.groupby('periodo_dt').agg({
            'contatos': 'sum',
            'aproveitados': 'sum',
            'vendas': 'sum'
        }).sort_index()
        
        monthly['taxa_conversao'] = monthly['vendas'] / monthly['contatos']
        monthly['conversao_ap'] = monthly['vendas'] / monthly['aproveitados']
        
        # Formatar datas para exibição
        monthly['periodo_formatado'] = monthly.index.map(format_period_display)
        
        # Gráfico de volume
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(
            x=monthly['periodo_formatado'], 
            y=monthly['contatos'], 
            name='Contatos',
            line=dict(color='#3498db')
        ))
        fig1.add_trace(go.Scatter(
            x=monthly['periodo_formatado'], 
            y=monthly['aproveitados'], 
            name='Aproveitados',
            line=dict(color='#2ecc71')
        ))
        fig1.add_trace(go.Scatter(
            x=monthly['periodo_formatado'], 
            y=monthly['vendas'], 
            name='Vendas',
            line=dict(color='#e74c3c')
        ))
        fig1.update_layout(
            title="Evolução Mensal - Volume",
            xaxis_title="Período",
            yaxis_title="Quantidade",
            paper_bgcolor='white', plot_bgcolor="#F0F1F3", font_color='#212529'
        )
        
        # Gráfico de taxas
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=monthly['periodo_formatado'], 
            y=monthly['taxa_conversao'], 
            name='Conversão Total',
            line=dict(color='#f39c12')
        ))
        fig2.add_trace(go.Scatter(
            x=monthly['periodo_formatado'], 
            y=monthly['conversao_ap'], 
            name='Conversão Aproveitados',
            line=dict(color='#9b59b6')
        ))
        fig2.update_layout(
            title="Evolução Mensal - Taxas de Conversão",
            xaxis_title="Período",
            yaxis_title="Taxa",
            yaxis_tickformat=".0%",
            paper_bgcolor='white', plot_bgcolor="#F0F1F3", font_color='#212529'
        )
        
        return html.Div([
            dcc.Graph(figure=fig1, className="mb-4"),
            dcc.Graph(figure=fig2)
        ])

    except Exception as e:
        print(f"Erro em render_monthly_trend: {str(e)}")
        return dbc.Alert(f"Erro ao exibir evolução mensal: {str(e)}", color="danger")

def render_top_channels(df): # Top Canais
    try:
        if 'origem' not in df.columns:
            return dbc.Alert("Dados de origem não disponíveis", color="danger")
        
        channel_efficiency = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum',
            'aproveitados': 'sum'
        }).reset_index()
        
        channel_efficiency['taxa_conversao'] = channel_efficiency['vendas'] / channel_efficiency['contatos']
        channel_efficiency['conversao_ap'] = channel_efficiency['vendas'] / channel_efficiency['aproveitados']
        channel_efficiency = channel_efficiency[channel_efficiency['contatos'] >= 50]
        
        if channel_efficiency.empty:
            return dbc.Alert("Sem dados suficientes para análise", color="warning")
        
        # Top 5 por conversão total
        top_conversion = channel_efficiency.nlargest(5, 'taxa_conversao').sort_values('taxa_conversao', ascending=True)
        fig1 = px.bar(
            top_conversion,
            y='origem',
            x='taxa_conversao',
            orientation='h',
            title="Top 5 Canais - Conversão Total",
            text='taxa_conversao',
            labels={'taxa_conversao': 'Taxa de Conversão', 'origem': 'Canal'}
        )
        fig1.update_traces(texttemplate='%{text:.1%}', textposition='auto')
        fig1.update_layout(
            xaxis_tickformat=".0%",
            paper_bgcolor='white', plot_bgcolor="#999CA0", font_color='#212529'
        )
        
        # Top 5 por conversão de leads
        top_lead_conversion = channel_efficiency.nlargest(5, 'conversao_ap').sort_values('conversao_ap', ascending=True)
        fig2 = px.bar(
            top_lead_conversion,
            y='origem',
            x='conversao_ap',
            orientation='h',
            title="Top 5 Canais - Conversão de Aproveitados",
            text='conversao_ap',
            labels={'conversao_ap': 'Taxa de Conversão', 'origem': 'Canal'}
        )
        fig2.update_traces(texttemplate='%{text:.1%}', textposition='auto')
        fig2.update_layout(
            xaxis_tickformat=".0%",
            paper_bgcolor='white', plot_bgcolor="#999CA0", font_color='#212529'
        )
        
        return dbc.Row([
            dbc.Col(dcc.Graph(figure=fig1)), 
            dbc.Col(dcc.Graph(figure=fig2))
        ])

    except Exception as e:
        print(f"Erro em render_top_channels: {str(e)}")
        return dbc.Alert(f"Erro ao exibir Top Canais: {str(e)}", color="danger")
    
def render_sales_efficiency(df): # Eficiência de Vendas
    try:
        if 'contatos' not in df.columns or 'vendas' not in df.columns:
            return dbc.Alert("Dados de eficiência não disponíveis", color="danger")
        
        grouped = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum'
        }).reset_index()
        
        grouped = grouped[grouped['vendas'] > 0]
        # Calcular eficiência como porcentagem
        grouped['eficiencia'] = (grouped['vendas'] / grouped['contatos']) * 100
        # Filtrar origens com pelo menos 10 vendas
        filtered = grouped[grouped['vendas'] >= 10]
        
        if filtered.empty:
            return dbc.Alert("Nenhuma origem com dados suficientes", color="warning")
        
        # Selecionar apenas os top 15
        top_15 = filtered.nlargest(15, 'eficiencia')
        top_15 = top_15.sort_values('eficiencia', ascending=True)
        
        fig = px.bar(
            top_15,
            y='origem',
            x='eficiencia',
            orientation='h',
            title="Eficiência de Vendas (Vendas por Lead - Top 15)",
            text='eficiencia',
            labels={'eficiencia': 'Eficiência (%)', 'origem': 'Origem'},
            color='eficiencia',
            color_continuous_scale='Blues'
        )
        fig.update_traces(
            texttemplate='%{text:.1f}%', 
            textposition='auto',
            hovertemplate='<b>%{y}</b><br>Eficiência: %{x:.1f}%<br>Vendas: %{customdata}',
            customdata=top_15['vendas']
        )
        fig.update_layout(
            paper_bgcolor='white', 
            plot_bgcolor="#F0F1F3", 
            font_color='#212529',
            xaxis=dict(
                title='Eficiência (%)',
                ticksuffix='%',
                showgrid=True,
                gridcolor='#dddddd'
            ),
            yaxis=dict(
                title='',
                automargin=True
            ),
            coloraxis_showscale=False
        )
        
        return dcc.Graph(figure=fig)

    except Exception as e:
        print(f"Erro em render_sales_efficiency: {str(e)}")
        return dbc.Alert(f"Erro ao exibir Eficiência de Vendas: {str(e)}", color="danger")

def render_correlation_view(df): # Correlação Leads-Vendas
    try:
        df_corr = calculate_correlations(df)
        
        if df_corr.empty:
            return dbc.Alert(
                "Dados insuficientes para calcular correlações (mínimo 3 períodos por origem)",
                color="warning"
            )
        
        # Separar em melhores e piores
        top_15 = df_corr.head(15).sort_values('correlacao', ascending=False)
        bottom_15 = df_corr.tail(15).sort_values('correlacao', ascending=True)
        
        # Criar gráfico dos melhores
        fig_top = px.bar(
            top_15,
            x='correlacao',
            y='origem',
            orientation='h',
            title='TOP 15 - Melhores Correlações (Leads x Vendas)',
            text='correlacao',
            labels={'correlacao': 'Coeficiente', 'origem': 'Canal'},
            color='correlacao',
            color_continuous_scale='Greens',
            range_color=[-1, 1]
        )
        
        # Personalizar layout dos melhores
        fig_top.update_traces(
            texttemplate='%{text:.2f}',
            textposition='auto',
            hovertemplate='<b>%{y}</b><br>Correlação: %{x:.2f}<br>Períodos: %{customdata}',
            customdata=top_15['n_periodos']
        )
        fig_top.update_layout(
            xaxis=dict(range=[-1.1, 1.1]),
            yaxis=dict(autorange="reversed"),
            paper_bgcolor='white',
            plot_bgcolor='#f8f9fa',
            font_color='#212529',
            coloraxis_showscale=False,
            height=500
        )
        
        # Criar gráfico dos piores
        fig_bottom = px.bar(
            bottom_15,
            x='correlacao',
            y='origem',
            orientation='h',
            title='TOP 15 - Piores Correlações (Leads x Vendas)',
            text='correlacao',
            labels={'correlacao': 'Coeficiente', 'origem': 'Canal'},
            color='correlacao',
            color_continuous_scale='Reds_r',  # Invertido para vermelho mais forte = pior
            range_color=[-1, 1]
        )
        
        # Personalizar layout dos piores
        fig_bottom.update_traces(
            texttemplate='%{text:.2f}',
            textposition='auto',
            hovertemplate='<b>%{y}</b><br>Correlação: %{x:.2f}<br>Períodos: %{customdata}',
            customdata=bottom_15['n_periodos']
        )
        fig_bottom.update_layout(
            xaxis=dict(range=[-1.1, 1.1]),
            yaxis=dict(autorange="reversed"),
            paper_bgcolor='white',
            plot_bgcolor='#f8f9fa',
            font_color='#212529',
            coloraxis_showscale=False,
            height=500
        )
        
        return html.Div([
            dcc.Graph(figure=fig_top),
            dcc.Graph(figure=fig_bottom)
        ])
    
    except Exception as e:
        print(f"Erro em render_correlation_view: {str(e)}")
        return dbc.Alert(f"Erro ao calcular correlações: {str(e)}", color="danger")
    
def render_scatter_plots(df): # Dispersão Leads x Vendas
    try:
        required_columns = ['contatos', 'aproveitados', 'vendas', 'origem']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            return dbc.Alert(
                f"Colunas obrigatórias ausentes: {', '.join(missing_columns)}",
                color="danger"
            )
        
        # Agregar dados por origem (somar todos os períodos)
        df_agg = df.groupby('origem', as_index=False).agg({
            'contatos': 'sum',
            'aproveitados': 'sum',
            'vendas': 'sum'
        })
        
        # Filtrar origens com dados válidos
        df_agg = df_agg[
            (df_agg['contatos'] > 0) & 
            (df_agg['vendas'] > 0) & 
            (df_agg['aproveitados'] > 0)
        ]
        
        if df_agg.empty:
            return dbc.Alert("Sem dados válidos para análise", color="warning")
        
        # Limitar a 15 canais para melhor visualização
        top_origins = df_agg.nlargest(15, 'contatos')['origem']
        df_filtered = df_agg[df_agg['origem'].isin(top_origins)]
        
        # Gráfico 1: Leads vs Vendas
        fig1 = px.scatter(
            df_filtered,
            x='contatos',
            y='vendas',
            color='origem',
            title='Relação Leads vs Vendas por Canal (Totais)',
            labels={
                'contatos': 'Total de Leads',
                'vendas': 'Total de Vendas',
                'origem': 'Canal'
            },
            size='vendas',  # Tamanho proporcional às vendas
            size_max=30,
            hover_data=['contatos', 'vendas'],
            text='origem'  # Mostrar nome da origem no ponto
        )
        
        # Personalizar layout
        fig1.update_traces(
            textposition='top center',
            marker=dict(line=dict(width=1, color='DarkSlateGrey')),
            textfont=dict(size=10)
        )
        fig1.update_layout(
            paper_bgcolor='white',
            plot_bgcolor='#f8f9fa',
            font_color='#212529',
            legend_title_text='Canal',
            height=500,
            showlegend=False  # Nomes já estão nos pontos
        )
        
        # Adicionar linha de tendência global
        try:
            fig1.add_trace(
                px.scatter(df_filtered, x='contatos', y='vendas', trendline="ols").data[1]
            )
        except:
            pass
        
        # Gráfico 2: Aproveitados vs Vendas
        fig2 = px.scatter(
            df_filtered,
            x='aproveitados',
            y='vendas',
            color='origem',
            title='Relação Leads Aproveitados vs Vendas por Canal (Totais)',
            labels={
                'aproveitados': 'Total de Leads Aproveitados',
                'vendas': 'Total de Vendas',
                'origem': 'Canal'
            },
            size='vendas',
            size_max=30,
            hover_data=['aproveitados', 'vendas'],
            text='origem'
        )
        
        # Personalizar layout
        fig2.update_traces(
            textposition='top center',
            marker=dict(line=dict(width=1, color='DarkSlateGrey')),
            textfont=dict(size=10)
        )
        fig2.update_layout(
            paper_bgcolor='white',
            plot_bgcolor='#f8f9fa',
            font_color='#212529',
            legend_title_text='Canal',
            height=500,
            showlegend=False
        )
        
        try:
            fig2.add_trace(
                px.scatter(df_filtered, x='aproveitados', y='vendas', trendline="ols").data[1]
            )
        except:
            pass
        
        return html.Div([
            dcc.Graph(figure=fig1),
            dcc.Graph(figure=fig2)
        ])
    
    except Exception as e:
        print(f"Erro em render_scatter_plots: {str(e)}")
        return dbc.Alert(f"Erro ao gerar gráficos de dispersão: {str(e)}", color="danger")
    
server = app.server

if __name__ == '__main__':
    app.run(debug=True)