import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
from typing import Dict, List, Tuple, Optional
import numpy as np

# Configuração da página
st.set_page_config(
    page_title="Análise de Matrículas - Educação Especial",
    page_icon="📊",
    layout="wide"
)

# Definir paletas de cores pastéis
PASTEL_COLORS = {
    # Para categorias binárias (2 valores) - cores contrastantes
    'binary': ['#a8d8b9', '#fdb4b4'],  # Verde pastel e Rosa pastel

    # Para múltiplas categorias - gradiente de azul pastel
    'gradient_blue': [
        '#2a5a7a',  # Azul mais escuro
        '#4b7c9e',
        '#6c9ec2',
        '#8db9d8',
        '#aed4ee',
        '#cfe8f7',
        '#e0f2fc',
        '#f0f9ff'  # Azul mais claro
    ],

    # Cor única para gráficos de linha
    'single': '#6c9ec2'
}

# Configurações globais de estilo
STYLE_CONFIG = {
    'font_family': 'Open Sans, sans-serif',
    'title_font': {
        'family': 'Open Sans, sans-serif',
        'size': 20,
        'color': 'black'
    },
    'margins': {
        't': 100,  # top
        'b': 80,  # bottom
        'l': 210,  # left
        'r': 80,  # right
        'pad': 4  # padding
    },
    'grid': {
        'x': {
            'showgrid': True,
            'gridwidth': 1,
            'gridcolor': '#EEEEEE',
            'zeroline': True,
            'zerolinewidth': 1,
            'zerolinecolor': '#444444'
        },
        'y': {
            'showgrid': False,
            'zeroline': False
        }
    }
}


class DataParser:
    """Parser especializado para a estrutura específica dos dados"""

    @staticmethod
    def parse_number_br(text: str) -> float:
        """Converte número formato BR para float"""
        if pd.isna(text) or text == 'N/A':
            return np.nan
        text = str(text).strip()
        # Remove pontos (separador de milhares) e troca vírgula por ponto
        text = text.replace('.', '').replace(',', '.')
        try:
            return float(text)
        except:
            return np.nan

    @staticmethod
    def extract_value_and_percent(text: str) -> Tuple[Optional[float], Optional[float]]:
        """Extrai valor absoluto e percentual de uma string"""
        if pd.isna(text) or text == 'N/A':
            return None, None

        text = str(text).strip()

        # Padrão: valor (percentual%)
        match = re.match(r'^([\d\.]+)\s*\((\d+[,\.]\d+)%\)$', text)
        if match:
            valor = DataParser.parse_number_br(match.group(1))
            percent = DataParser.parse_number_br(match.group(2))
            return valor, percent

        # Apenas valor
        match = re.match(r'^([\d\.]+)$', text)
        if match:
            valor = DataParser.parse_number_br(match.group(1))
            return valor, None

        return None, None

    @staticmethod
    def is_section_header(text: str) -> bool:
        """Verifica se é um cabeçalho de seção"""
        if pd.isna(text):
            return False
        return '═══' in str(text)

    @staticmethod
    def clean_section_name(text: str) -> str:
        """Limpa o nome da seção"""
        return str(text).replace('═', '').strip()

    @staticmethod
    def parse_composite_line(metric: str, value: str) -> List[Dict]:
        """Parse linhas compostas como Top 5 ou composição por idade"""
        results = []

        # Padrão para composição de idades: "X anos: Y, Z anos: W, ..."
        if 'anos:' in str(value):
            # Ex: "12 anos: 803, 13 anos: 2.459, ..."
            parts = str(value).split(',')
            for part in parts:
                match = re.match(r'(\d+)\s*anos:\s*([\d\.]+)', part.strip())
                if match:
                    results.append({
                        'idade': int(match.group(1)),
                        'quantidade': DataParser.parse_number_br(match.group(2))
                    })

        return results


def load_and_parse_excel(file) -> Dict:
    """Carrega e faz o parsing do arquivo Excel"""
    try:
        # Lê todas as abas
        excel_file = pd.ExcelFile(file)
        data = {}

        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file, sheet_name=sheet_name, header=None)

            # Assume que primeira coluna é Métrica e segunda é Valor
            if df.shape[1] < 2:
                st.warning(f"Aba '{sheet_name}' não tem 2 colunas. Pulando...")
                continue

            df.columns = ['Metrica', 'Valor']

            # Parse dos dados
            parsed_data = parse_sheet_data(df)
            data[sheet_name] = parsed_data

        return data
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {str(e)}")
        return {}


def parse_sheet_data(df: pd.DataFrame) -> Dict:
    """Faz o parsing completo de uma aba"""
    sections = {}
    current_section = None
    current_subsection = None
    total_matriculas = None

    for idx, row in df.iterrows():
        metric = row['Metrica']
        value = row['Valor']

        # Pula linhas vazias
        if pd.isna(metric):
            continue

        metric = str(metric)  # Não strip ainda para preservar espaços iniciais

        # Verifica se a linha está indentada (começa com espaços)
        is_indented = metric.startswith('  ') or metric.startswith('\t')

        # Agora faz o strip normal
        metric = metric.strip()

        # Verifica se é cabeçalho de seção
        if DataParser.is_section_header(metric):
            current_section = DataParser.clean_section_name(metric)
            current_subsection = None
            sections[current_section] = {
                'items': [],
                'subsections': {},
                'composites': {}
            }
            continue

        # Se não há seção atual, cria uma "Geral"
        if current_section is None:
            current_section = "Geral"
            sections[current_section] = {
                'items': [],
                'subsections': {},
                'composites': {}
            }

        # Procura por Total de matrículas para cálculo de percentuais
        if 'Total de matrículas' in metric:
            valor, _ = DataParser.extract_value_and_percent(value)
            if valor:
                total_matriculas = valor

        # Detecta subsecções específicas (apenas se não está indentado)
        if not is_indented and metric.endswith(':'):
            subsection_name = metric.rstrip(':')

            # Mapeia nomes de subsecções para categorias apropriadas
            if 'Distribuição por Sexo' in subsection_name:
                current_subsection = 'Sexo'
            elif 'Distribuição por Zona' in subsection_name:
                current_subsection = 'Zona'
            elif 'Distribuição por Cor/Raça' in subsection_name:
                current_subsection = 'Cor/Raça'
            elif 'Distribuição por número de deficiências' in subsection_name:
                current_subsection = 'Distribuição por número de deficiências'
            elif 'Comorbidades' in subsection_name:
                current_subsection = 'Comorbidades'
            elif 'Detalhamento da distorção por etapa' in subsection_name:
                current_subsection = 'Detalhamento da distorção por etapa'
            elif 'Top 5 municípios' in subsection_name:
                current_subsection = 'Top 5 municípios'
            elif 'Composição de idades por etapa' in subsection_name:
                current_subsection = 'Composição por idade/etapa'
            else:
                current_subsection = subsection_name

            if current_subsection not in sections[current_section]['subsections']:
                sections[current_section]['subsections'][current_subsection] = []
            continue

        # Parse do valor
        if pd.notna(value):
            value_str = str(value).strip()

            # Tenta extrair valor e percentual
            valor, percentual = DataParser.extract_value_and_percent(value_str)

            # Se é uma linha de composição por idade
            if 'anos:' in value_str:
                composite_data = DataParser.parse_composite_line(metric, value_str)
                if composite_data:
                    sections[current_section]['composites'][metric] = composite_data
                    continue

            # Calcula percentual se não existir e tivermos o total
            if valor and percentual is None and total_matriculas:
                percentual = (valor / total_matriculas) * 100

            # Adiciona ao local apropriado
            item_data = {
                'metrica': metric,
                'valor': valor,
                'percentual': percentual,
                'valor_original': value_str
            }

            if current_subsection:
                sections[current_section]['subsections'][current_subsection].append(item_data)
            else:
                sections[current_section]['items'].append(item_data)

    return {
        'sections': sections,
        'total_matriculas': total_matriculas
    }


def format_number_br(value: float, is_percent: bool = False) -> str:
    """Formata número no padrão brasileiro"""
    if pd.isna(value):
        return "N/A"

    if is_percent:
        return f"{value:,.1f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    else:
        return f"{value:,.0f}".replace(',', '.')


def create_bar_chart(data: pd.DataFrame, title: str, x_col: str, y_col: str,
                     percent_col: str = None, deficiency_type: str = "",
                     font_sizes: dict = None) -> go.Figure:
    """Cria gráfico de barras interativo com formatação pt-BR e estilo aprimorado"""

    if font_sizes is None:
        font_sizes = {
            'title': 20,
            'subtitle': 14,
            'labels': 12,
            'values': 11,
            'reference': 10
        }

    # Ordena por valor decrescente
    data = data.sort_values(by=y_col, ascending=True)

    # Define cores baseado no número de categorias
    num_bars = len(data)

    if num_bars == 2:
        # Para categorias binárias (ex: M/F, Rural/Urbana)
        colors = PASTEL_COLORS['binary']
    else:
        # Para múltiplas categorias, usa gradiente de tons pastéis
        gradient = PASTEL_COLORS['gradient_blue']
        if num_bars <= len(gradient):
            # Seleciona cores espaçadas do gradiente
            step = len(gradient) // num_bars
            colors = [gradient[i * step] for i in range(num_bars)]
        else:
            # Se tiver mais barras que cores, repete o gradiente
            colors = [gradient[i % len(gradient)] for i in range(num_bars)]
        # Inverte para que valores maiores tenham cores mais escuras
        colors = colors[::-1]

    # Texto para hover e labels com formatação pt-BR
    if percent_col and percent_col in data.columns:
        hover_text = [
            f"<b>{row[x_col]}</b><br>"
            f"Valor: {format_number_br(row[y_col])}<br>"
            f"Percentual: {format_number_br(row[percent_col], True)}%"
            for _, row in data.iterrows()
        ]
        text_labels = [f"{format_number_br(row[y_col])} ({format_number_br(row[percent_col], True)}%)"
                       for _, row in data.iterrows()]
    else:
        hover_text = [
            f"<b>{row[x_col]}</b><br>"
            f"Valor: {format_number_br(row[y_col])}"
            for _, row in data.iterrows()
        ]
        text_labels = [f"{format_number_br(row[y_col])}" for _, row in data.iterrows()]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=data[x_col],
        x=data[y_col],
        orientation='h',
        text=text_labels,
        textposition='outside',
        textfont=dict(
            size=font_sizes['values'],
            family='Open Sans, sans-serif'
        ),
        hovertemplate='%{hovertext}<extra></extra>',
        hovertext=hover_text,
        marker_color=colors,
        width=0.8  # Bar width 80%
    ))

    # Título completo com tamanhos personalizados
    full_title = f"<b><span style='font-size:{font_sizes['title']}px'>Quantidade de matrículas da Educação Especial por {title}</span></b><br>"
    full_title += f"<span style='font-size:{font_sizes['subtitle']}px'>Tipo de deficiência: {deficiency_type} | "
    full_title += "Rede: Pública — estadual e municipal | Pernambuco | 2024</span>"

    max_value = data[y_col].max()

    fig.update_layout(
        template='plotly',  # Theme classic
        title={
            'text': full_title,
            'x': 0.5,
            'xanchor': 'center',
            'font': {
                'family': 'Open Sans, sans-serif'
            }
        },
        xaxis=dict(
            title={
                'text': "Quantidade de Matrículas",
                'font': dict(
                    size=font_sizes['labels'],
                    family='Open Sans, sans-serif'
                )
            },
            tickfont=dict(
                size=font_sizes['values'],
                family='Open Sans, sans-serif'
            ),
            tickformat=',.0f',
            separatethousands=True,
            range=[-25, max_value * 1.15],  # Min: -25 conforme solicitado
            showgrid=True,
            gridwidth=1,
            gridcolor='#EEEEEE',
            zeroline=True,
            zerolinewidth=2,  # Aumentado para maior visibilidade
            zerolinecolor='#444444',
            layer='above traces',  # Garante que a linha zero fique acima das barras
            automargin=True
        ),
        yaxis=dict(
            title={
                'text': "",
                'font': dict(
                    size=font_sizes['labels'],
                    family='Open Sans, sans-serif'
                )
            },
            tickfont=dict(
                size=font_sizes['labels'],
                family='Open Sans, sans-serif'
            ),
            range=[-1.3, len(data)],  # Min: -1.3 conforme solicitado
            showgrid=False,
            zeroline=False,
            automargin=True
        ),
        height=max(400, len(data) * 50),
        showlegend=False,
        hovermode='closest',
        margin=dict(
            l=STYLE_CONFIG['margins']['l'],
            r=STYLE_CONFIG['margins']['r'],
            t=STYLE_CONFIG['margins']['t'],
            b=STYLE_CONFIG['margins']['b'],
            pad=STYLE_CONFIG['margins']['pad']
        ),
        font=dict(
            family='Open Sans, sans-serif'
        ),
        bargap=0,  # Bar padding: 0
        bargroupgap=0
    )

    # Adiciona rodapé com negrito apenas em "Fonte:"
    fig.add_annotation(
        text="<b>Fonte:</b> Elaboração própria, com base nos dados informados pelo Inep (doc. 2).",
        xref="paper", yref="paper",
        x=0, y=-0.45,
        showarrow=False,
        font=dict(
            size=font_sizes['reference'],
            color="gray",
            family='Open Sans, sans-serif'
        ),
        xanchor='left'
    )

    return fig


def create_line_chart(data: pd.DataFrame, title: str, x_col: str, y_col: str,
                      deficiency_type: str = "", font_sizes: dict = None) -> go.Figure:
    """Cria gráfico de linha interativo com formatação pt-BR e estilo aprimorado"""

    if font_sizes is None:
        font_sizes = {
            'title': 20,
            'subtitle': 14,
            'labels': 12,
            'values': 11,
            'reference': 10
        }

    fig = go.Figure()

    # Ordena por x (geralmente idade)
    data = data.sort_values(by=x_col)
    data = data.reset_index(drop=True)  # Reset index para garantir sequência correta

    # Usa cor única da paleta pastel
    line_color = PASTEL_COLORS['single']

    # Prepara os textos - mostra valores apenas em índices pares (0, 2, 4, etc.)
    # Isso corresponde a mostrar valores para idades 1, 3, 5, etc. se começar de 1
    text_values = []
    for idx, value in enumerate(data[y_col]):
        # Mostra valor se o índice for par (0, 2, 4...)
        if idx % 2 == 0:
            text_values.append(format_number_br(value))
        else:
            text_values.append("")  # Não mostra valor para índices ímpares

    fig.add_trace(go.Scatter(
        x=data[x_col],
        y=data[y_col],
        mode='lines+markers+text',
        text=text_values,
        textposition="top center",
        textfont=dict(
            size=font_sizes['values'],
            family='Open Sans, sans-serif'
        ),
        line=dict(color=line_color, width=3),
        marker=dict(size=10, color=line_color),
        hovertemplate='<b>%{x}</b><br>Quantidade: ' +
                      '%{y:,.0f}<extra></extra>'
    ))

    # Título completo com tamanhos personalizados
    full_title = f"<b><span style='font-size:{font_sizes['title']}px'>Quantidade de matrículas da Educação Especial por {title}</span></b><br>"
    full_title += f"<span style='font-size:{font_sizes['subtitle']}px'>Tipo de deficiência: {deficiency_type} | "
    full_title += "Rede: Pública — estadual e municipal | Pernambuco | 2024</span>"

    fig.update_layout(
        template='plotly',  # Theme classic
        title={
            'text': full_title,
            'x': 0.5,
            'xanchor': 'center',
            'font': {
                'family': 'Open Sans, sans-serif'
            }
        },
        xaxis=dict(
            title={
                'text': x_col.capitalize(),
                'font': dict(
                    size=font_sizes['labels'],
                    family='Open Sans, sans-serif'
                )
            },
            tickfont=dict(
                size=font_sizes['values'],
                family='Open Sans, sans-serif'
            ),
            showgrid=True,
            gridwidth=1,
            gridcolor='#EEEEEE',
            zeroline=True,
            zerolinewidth=2,  # Aumentado para maior visibilidade
            zerolinecolor='#444444',
            automargin=True
        ),
        yaxis=dict(
            title={
                'text': "Quantidade de Matrículas",
                'font': dict(
                    size=font_sizes['labels'],
                    family='Open Sans, sans-serif'
                )
            },
            tickfont=dict(
                size=font_sizes['values'],
                family='Open Sans, sans-serif'
            ),
            tickformat=',.0f',
            separatethousands=True,
            showgrid=False,
            zeroline=False,
            automargin=True
        ),
        height=500,
        showlegend=False,
        hovermode='x unified',
        margin=dict(
            l=STYLE_CONFIG['margins']['l'],
            r=STYLE_CONFIG['margins']['r'],
            t=STYLE_CONFIG['margins']['t'],
            b=STYLE_CONFIG['margins']['b'],
            pad=STYLE_CONFIG['margins']['pad']
        ),
        font=dict(
            family='Open Sans, sans-serif'
        )
    )

    # Adiciona rodapé com negrito apenas em "Fonte:"
    fig.add_annotation(
        text="<b>Fonte:</b> Elaboração própria, com base nos dados informados pelo Inep (doc. 2).",
        xref="paper", yref="paper",
        x=0, y=-0.35,
        showarrow=False,
        font=dict(
            size=font_sizes['reference'],
            color="gray",
            family='Open Sans, sans-serif'
        ),
        xanchor='left'
    )

    return fig


def categorize_data(section_data: dict, section_name: str) -> dict:
    """Categoriza os dados de uma seção de forma mais inteligente"""
    categories = {}

    # Primeiro, adiciona subsecções como categorias
    for subsection_name, items in section_data['subsections'].items():
        if items:
            categories[subsection_name] = items

    # Depois, categoriza os itens não subseccionados
    for item in section_data['items']:
        metric = item['metrica'].lower()

        # Categorização para itens não subseccionados
        categorized = False

        # Análise de distorção idade-série
        if 'idade apropriada' in metric or 'distorção idade-série' in metric:
            if 'Status idade-série' not in categories:
                categories['Status idade-série'] = []
            categories['Status idade-série'].append(item)
            categorized = True

        # Dependência administrativa
        elif any(x in metric for x in ['municipal', 'estadual', 'federal', 'rede']):
            if 'Dependência Administrativa' not in categories:
                categories['Dependência Administrativa'] = []
            categories['Dependência Administrativa'].append(item)
            categorized = True

        # Idade (quando não está em subsecção)
        elif re.search(r'\d+\s*anos', metric) and 'ensino' not in metric:
            if 'Idade' not in categories:
                categories['Idade'] = []
            categories['Idade'].append(item)
            categorized = True

        # Etapas de ensino
        elif any(x in metric for x in ['infantil', 'fundamental', 'médio', 'eja', 'profissional']):
            if 'Etapa de Ensino' not in categories:
                categories['Etapa de Ensino'] = []
            categories['Etapa de Ensino'].append(item)
            categorized = True

        # Municípios
        elif any(x in metric for x in ['recife', 'jaboatão', 'olinda', 'paulista', 'caruaru',
                                       'petrolina', 'garanhuns', 'camaragibe']):
            if 'Municípios' not in categories:
                categories['Municípios'] = []
            categories['Municípios'].append(item)
            categorized = True

        # Limpa o nome da seção para usar como categoria padrão (ex: "3. NOME" -> "NOME")
        default_category_name = re.sub(r'^\d+\.\s*', '', section_name).strip().capitalize()

        # Se não foi categorizado, usa o nome limpo da seção como categoria
        if not categorized and item['valor']:
            if default_category_name not in categories:
                categories[default_category_name] = []
            categories[default_category_name].append(item)

    # Adiciona composites como categorias especiais
    if section_data['composites']:
        for key, value in section_data['composites'].items():
            if 'composição' in key.lower() and 'idade' in key.lower():
                categories['Composição por Idade/Etapa'] = value

    return categories


def main():
    st.title("📊 Análise de Matrículas - Educação Especial")
    st.markdown("### Sistema de Visualização de Dados do Censo Escolar")

    # Sidebar com controles de customização
    with st.sidebar:
        st.markdown("### ⚙️ Configurações de Visualização")

        # Controles de tamanho de fonte
        st.markdown("#### Tamanhos de Fonte")
        font_sizes = {
            'title': st.slider("Título", 14, 30, 20),
            'subtitle': st.slider("Subtítulo", 10, 20, 14),
            'labels': st.slider("Rótulos dos eixos", 10, 18, 12),
            'values': st.slider("Valores", 8, 16, 11),
            'reference': st.slider("Referências", 8, 14, 10)
        }

        st.divider()

        st.markdown("### ℹ️ Informações")
        st.info(
            "**Como usar:**\n"
            "1. Faça upload do arquivo Excel\n"
            "2. Selecione o tipo de deficiência (aba)\n"
            "3. Escolha a seção de análise\n"
            "4. Selecione a categoria para visualização\n\n"
            "**Observações:**\n"
            "- Números formatados em pt-BR\n"
            "- Percentuais calculados automaticamente\n"
            "- Gráficos interativos (hover para detalhes)\n"
            "- Cores: Categorias binárias contrastantes, múltiplas em gradiente\n"
            "- Fonte: Open Sans"
        )

    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Selecione o arquivo Excel",
        type=['xlsx', 'xls'],
        help="Arquivo deve conter abas com dados estruturados de matrículas por tipo de deficiência"
    )

    if uploaded_file is not None:
        # Carrega e processa os dados
        with st.spinner("Processando arquivo..."):
            data = load_and_parse_excel(uploaded_file)

        if not data:
            st.error("Não foi possível processar o arquivo.")
            return

        # Interface de seleção
        col1, col2, col3 = st.columns(3)

        with col1:
            # Seleção de aba (tipo de deficiência)
            selected_sheet = st.selectbox(
                "Tipo de Deficiência",
                options=list(data.keys()),
                help="Selecione o tipo de deficiência para análise"
            )

        if selected_sheet:
            sheet_data = data[selected_sheet]
            sections = sheet_data['sections']

            with col2:
                # Seleção de seção
                section_options = list(sections.keys())
                selected_section = st.selectbox(
                    "Seção",
                    options=section_options,
                    help="Selecione a seção de análise"
                )

            if selected_section:
                section_data = sections[selected_section]

                with col3:
                    # Categoriza os dados de forma inteligente
                    categories = categorize_data(section_data, selected_section)

                    # Seleção de categoria
                    if categories:
                        selected_category = st.selectbox(
                            "Categoria de Análise",
                            options=list(categories.keys()),
                            help="Selecione a categoria para visualização"
                        )
                    else:
                        st.warning("Nenhuma categoria de dados encontrada nesta seção.")
                        return

        # Separador visual
        st.divider()

        # Geração do gráfico
        if selected_sheet and selected_section and selected_category:
            with st.container():
                st.subheader(f"Visualização: {selected_category}")

                # Prepara dados para visualização
                if selected_category == 'Composição por Idade/Etapa':
                    # Dados de composição especial
                    comp_data = categories[selected_category]
                    if comp_data:
                        df = pd.DataFrame(comp_data)
                        fig = create_line_chart(
                            df,
                            "Idade",
                            'idade',
                            'quantidade',
                            selected_sheet,
                            font_sizes
                        )
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    # Dados regulares
                    cat_data = categories[selected_category]
                    if cat_data:
                        # Cria DataFrame
                        df_data = []
                        for item in cat_data:
                            df_data.append({
                                'Categoria': item['metrica'],
                                'Valor': item['valor'] if item['valor'] else 0,
                                'Percentual': item['percentual'] if item['percentual'] else 0
                            })

                        df = pd.DataFrame(df_data)

                        # Remove linhas com valor zero ou nulo
                        df = df[df['Valor'] > 0]

                        if not df.empty:
                            # Determina o nome correto para o título
                            title_category = selected_category
                            if selected_category == 'Indicadores Gerais':
                                title_category = "Indicadores"
                            elif selected_category == 'Status idade-série':
                                title_category = "Status Idade-Série"

                            # Cria gráfico apropriado
                            if selected_category == 'Idade' or 'idade' in selected_category.lower():
                                # Para idade, usa gráfico de linha
                                df['Idade_Num'] = df['Categoria'].str.extract(r'(\d+)').astype(float)
                                df = df.sort_values('Idade_Num')
                                fig = create_line_chart(
                                    df,
                                    title_category,
                                    'Idade_Num',
                                    'Valor',
                                    selected_sheet,
                                    font_sizes
                                )
                            else:
                                # Para outras categorias, usa gráfico de barras
                                fig = create_bar_chart(
                                    df,
                                    title_category,
                                    'Categoria',
                                    'Valor',
                                    'Percentual',
                                    selected_sheet,
                                    font_sizes
                                )

                            st.plotly_chart(fig, use_container_width=True)

                            # Exibe tabela de dados
                            with st.expander("📋 Ver dados tabulares"):
                                # Formata a tabela com padrão brasileiro
                                df_display = df.copy()
                                df_display['Valor'] = df_display['Valor'].apply(lambda x: format_number_br(x))
                                df_display['Percentual'] = df_display['Percentual'].apply(
                                    lambda x: f"{format_number_br(x, True)}%" if x > 0 else "—"
                                )
                                st.dataframe(df_display, use_container_width=True)

                            # Validação e estatísticas
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Total da Categoria", format_number_br(df['Valor'].sum()))
                            with col2:
                                if sheet_data['total_matriculas']:
                                    st.metric("Total de Matrículas",
                                              format_number_br(sheet_data['total_matriculas']))
                            with col3:
                                if sheet_data['total_matriculas']:
                                    coverage = (df['Valor'].sum() / sheet_data['total_matriculas']) * 100
                                    st.metric("Cobertura", f"{format_number_br(coverage, True)}%")
                        else:
                            st.warning("Nenhum dado válido encontrado para esta categoria.")

        # Resumo do arquivo no sidebar
        if data:
            with st.sidebar:
                st.markdown("### 📊 Resumo do Arquivo")
                st.write(f"**Abas encontradas:** {len(data)}")
                for sheet in data.keys():
                    sections_count = len(data[sheet]['sections'])
                    st.write(f"- {sheet}: {sections_count} seções")


if __name__ == "__main__":
    main()