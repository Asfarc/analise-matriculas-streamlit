import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
from typing import Dict, List, Tuple, Optional
import numpy as np
import os
import requests
from io import BytesIO
import textwrap

# Configuração da página
st.set_page_config(
    page_title="Análise de Matrículas - Educação Especial",
    page_icon="📊",
    layout="wide"
)

# Configuração para exportação A4
A4_CONFIG = {
    'width_cm': 16,  # Largura útil A4 (21cm - 3cm - 2cm margens)
    'dpi': 150,  # DPI adequado para documentos digitais
    'width_px': 945,  # 16cm × 59.06 px/cm @ 150 DPI
    'bar_height_base': 400,  # Altura base para gráficos de barras
    'bar_height_per_item': 50,  # Altura adicional por barra
    'bar_height_max': 1200,  # Altura máxima para gráficos de barras
    'line_height': 600  # Altura fixa para gráficos de linha
}

# Definir paletas de cores pastéis
PASTEL_COLORS = {
    # Para categorias binárias (2 valores) - cores contrastantes
    'binary': ['#fdb4b4', '#5c88ab'],  # Verde pastel e Rosa pastel

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
            'zerolinewidth': 10,
            'zerolinecolor': '#444444',
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
        if pd.isna(text):
            return None, None

        text = str(text).strip()

        # Verifica se é N/A ou contém N/A
        if text.upper().startswith('N/A') or text.upper() == 'N/A':
            return None, None

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


def load_default_excel():
    """Carrega o arquivo Excel padrão do repositório ou local"""
    default_file = "analise_deficiencias"

    # Primeiro tenta carregar localmente
    if os.path.exists(default_file):
        try:
            return pd.ExcelFile(default_file)
        except:
            pass
    return None


def load_and_parse_excel(file) -> Dict:
    """Carrega e faz o parsing do arquivo Excel"""
    try:
        # Se file é um caminho string, carrega o arquivo
        if isinstance(file, str):
            excel_file = pd.ExcelFile(file)
        else:
            excel_file = pd.ExcelFile(file)

        data = {}

        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

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

        # Preserva espaços iniciais para detectar indentação
        metric_raw = str(metric)
        is_indented = metric_raw.startswith('  ') or metric_raw.startswith('\t')

        # Remove espaços extras
        metric = metric_raw.strip()

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

        # Detecta subsecções específicas (apenas linhas não-indentadas que terminam com ':')
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
        if pd.notna(value) or (pd.isna(value) and metric):  # Inclui linhas com métrica mas sem valor
            value_str = str(value).strip() if pd.notna(value) else ""

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

            # Sempre adiciona, mesmo com valor None (para preservar N/A)
            if current_subsection:
                sections[current_section]['subsections'][current_subsection].append(item_data)
            elif metric:  # Só adiciona se houver métrica
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
    """Cria gráfico de barras com linha vertical em x=0 sempre visível"""

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

    # Define cores
    num_bars = len(data)
    if num_bars == 2:
        colors = PASTEL_COLORS['binary']
    else:
        gradient = PASTEL_COLORS['gradient_blue']
        if num_bars <= len(gradient):
            step = max(1, len(gradient) // num_bars)
            colors = [gradient[min(i * step, len(gradient) - 1)] for i in range(num_bars)]
        else:
            colors = [gradient[i % len(gradient)] for i in range(num_bars)]
        colors = colors[::-1]

    # Prepara textos
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

    # Adiciona as barras
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
        width=0.8
    ))

    # Título principal
    main_title_html = f"<span style='font-weight: bold; font-size:{font_sizes['title']}px'>Matrículas da Educação Especial — {title}</span>"

    # Subtítulo
    subtitle_html = (
        f"<span style='font-size:{font_sizes['subtitle']}px'>"
        f"<span style='font-weight: bold;'>Tipo de deficiência:</span> {deficiency_type} <span style='font-weight: bold;'>|</span> "
        f"<span style='font-weight: bold;'>Rede:</span> Pública <span style='font-weight: bold;'>—</span> Estadual e Municipal <span style='font-weight: bold;'>|</span> "
        f"<span style='font-weight: bold;'>Pernambuco</span> <span style='font-weight: bold;'>|</span> <span style='font-weight: bold;'>2024</span>"
        f"</span>"
    )

    # Combina título e subtítulo
    full_title = f"{main_title_html}<br>{subtitle_html}"

    max_value = data[y_col].max() if not data.empty else 100

    # Calcula altura dinâmica baseada no número de barras
    height = min(
        max(A4_CONFIG['bar_height_base'], num_bars * A4_CONFIG['bar_height_per_item']),
        A4_CONFIG['bar_height_max']
    )

    # Layout
    fig.update_layout(
        template='seaborn',
        title={
            'text': full_title,
            'x': 0.5,
            'xanchor': 'center',
            'font': {'family': 'Open Sans, sans-serif'}
        },
        xaxis=dict(
            title={
                'text': "Quantidade de Matrículas",
                'font': dict(size=font_sizes['labels'], family='Open Sans, sans-serif')
            },
            tickfont=dict(size=font_sizes['values'], family='Open Sans, sans-serif'),
            tickformat=',.0f',
            separatethousands=True,
            # Ajusta range dinamicamente baseado no tamanho da fonte
            range=[-25, max_value * (1.25 + (font_sizes['values'] - 11) * 0.02)],
            showgrid=True,
            gridwidth=1,
            gridcolor='#EEEEEE',
            zeroline=False,
            automargin=True
        ),
        yaxis=dict(
            title="",
            tickfont=dict(size=font_sizes['labels'], family='Open Sans, sans-serif'),
            range=[-1.3, len(data)],
            showgrid=False,
            zeroline=False,
            automargin=True
        ),
        height=height,
        width=A4_CONFIG['width_px'],
        showlegend=False,
        hovermode='closest',
        margin=dict(
            l=150,
            # Margem direita dinâmica baseada no tamanho da fonte
            r=max(80, 40 + (font_sizes['values'] - 11) * 8),
            t=120,
            b=80,
            pad=4
        ),
        font=dict(family='Open Sans, sans-serif'),
        bargap=0.2,
        bargroupgap=0
    )

    # Adiciona linha vertical em x=0
    fig.add_shape(
        type="line",
        x0=0, x1=0,
        y0=-1.5, y1=len(data),
        line=dict(color="#333333", width=2.5),
        layer="above",
        xref="x", yref="y"
    )

    # Rodapé
    fig.add_annotation(
        text="<b>Fonte:</b> Elaboração própria, com base nos dados informados pelo Inep (doc. 2).",
        xref="paper",
        yref="paper",
        x=0.0,
        y=-0.30,
        showarrow=False,
        font=dict(
            size=font_sizes['reference'],
            color="#666666",
            family='Open Sans, sans-serif'
        ),
        xanchor='left',
        yanchor='top',
        align='left'
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
    data = data.reset_index(drop=True)

    # Usa cor única da paleta pastel
    line_color = PASTEL_COLORS['single']

    # Prepara os textos - mostra valores apenas em índices pares
    text_values = []
    for idx, value in enumerate(data[y_col]):
        if idx % 2 == 0:
            text_values.append(format_number_br(value))
        else:
            text_values.append("")

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
        hovertemplate='<b>%{x}</b><br>Quantidade: %{y:,.0f}<extra></extra>'
    ))

    # Título principal
    main_title_html = f"<span style='font-weight: bold; font-size:{font_sizes['title']}px'>Matrículas da Educação Especial — {title}</span>"

    # Subtítulo
    subtitle_html = (
        f"<span style='font-size:{font_sizes['subtitle']}px'>"
        f"<span style='font-weight: bold;'>Tipo de deficiência:</span> {deficiency_type} <span style='font-weight: bold;'>|</span> "
        f"<span style='font-weight: bold;'>Rede:</span> Pública <span style='font-weight: bold;'>—</span> Estadual e Municipal <span style='font-weight: bold;'>|</span> "
        f"<span style='font-weight: bold;'>Pernambuco</span> <span style='font-weight: bold;'>|</span> <span style='font-weight: bold;'>2024</span>"
        f"</span>"
    )

    # Combina título e subtítulo
    full_title = f"{main_title_html}<br>{subtitle_html}"

    fig.update_layout(
        template='seaborn',
        title={
            'text': full_title,
            'x': 0.5,
            'xanchor': 'center',
            'y': 0.98,
            'yanchor': 'top',
            'font': {'family': 'Open Sans, sans-serif'}
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
            zeroline=False,
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
            automargin=True,
            layer='above traces'
        ),
        height=A4_CONFIG['line_height'],
        width=A4_CONFIG['width_px'],
        showlegend=False,
        hovermode='x unified',
        margin=dict(
            l=100,
            r=40,
            t=120,
            b=80,
            pad=4
        ),
        font=dict(
            family='Open Sans, sans-serif'
        )
    )

    # Rodapé
    fig.add_annotation(
        text="<b>Fonte:</b> Elaboração própria, com base nos dados informados pelo Inep (doc. 2).",
        xref="paper",
        yref="paper",
        x=0.0,
        y=-0.30,
        showarrow=False,
        font=dict(
            size=font_sizes['reference'],
            color="#666666",
            family='Open Sans, sans-serif'
        ),
        xanchor='left',
        yanchor='top',
        align='left'
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
        if not categorized and (item['valor'] or item.get('valor_original')):
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

    # Inicializa o estado da sessão
    if 'data' not in st.session_state:
        st.session_state.data = None
        st.session_state.file_name = None

    # Tenta carregar o arquivo padrão automaticamente
    default_file = "analise_deficiencias.xlsx"

    # Upload do arquivo ou uso do padrão
    col1, col2 = st.columns([3, 1])

    with col1:
        uploaded_file = st.file_uploader(
            "Selecione outro arquivo Excel (opcional)",
            type=['xlsx', 'xls'],
            help="Arquivo deve conter abas com dados estruturados de matrículas por tipo de deficiência"
        )

    with col2:
        if st.button("🔄 Recarregar arquivo padrão"):
            st.session_state.data = None
            st.session_state.file_name = None
            st.rerun()

    # Decide qual arquivo usar
    file_to_load = None
    file_name = None

    if uploaded_file is not None:
        file_to_load = uploaded_file
        file_name = uploaded_file.name
    elif st.session_state.data is None:  # Carrega o padrão apenas se não há dados
        if os.path.exists(default_file):
            file_to_load = default_file
            file_name = default_file
            st.success(f"✅ Arquivo padrão '{default_file}' carregado automaticamente!")
        else:
            st.warning(f"⚠️ Arquivo padrão '{default_file}' não encontrado. Por favor, faça upload de um arquivo.")

    # Carrega e processa o arquivo se necessário
    if file_to_load and (st.session_state.file_name != file_name):
        with st.spinner("Processando arquivo..."):
            data = load_and_parse_excel(file_to_load)
            if data:
                st.session_state.data = data
                st.session_state.file_name = file_name

    # Usa os dados da sessão
    data = st.session_state.data

    if data:
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
                        selected_category = None

        # Separador visual
        st.divider()

        # Geração do gráfico
        if selected_sheet and selected_section and selected_category:
            with st.container():
                st.subheader(f"Visualização: {selected_category}")

                # Debug opcional
                if st.checkbox("🔍 Mostrar debug", value=False):
                    st.write(f"Categoria: {selected_category}")
                    st.write(f"Número de itens encontrados: {len(categories.get(selected_category, []))}")
                    if selected_category in categories and len(categories[selected_category]) > 0:
                        st.write("Primeiros 3 itens:")
                        for i, item in enumerate(categories[selected_category][:3]):
                            st.write(f"  {i + 1}. {item}")

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

                        # Configuração de exportação
                        config = {
                            'toImageButtonOptions': {
                                'format': 'png',
                                'filename': f'grafico_{selected_sheet}_{selected_category}'.lower().replace(' ', '_'),
                                'width': A4_CONFIG['width_px'],
                                'height': A4_CONFIG['line_height'],
                                'scale': 1  # Importante: sem multiplicação
                            },
                            'displaylogo': False
                        }
                        st.plotly_chart(fig, use_container_width=True, config=config)
                else:
                    # Dados regulares
                    cat_data = categories[selected_category]
                    if cat_data:
                        # Cria DataFrame
                        df_data = []
                        for item in cat_data:
                            # Filtro para remover métricas específicas do RESUMO GERAL
                            if selected_section == "RESUMO GERAL DO DATASET":
                                metricas_excluir = [
                                    'total de matrículas',
                                    'total de registros',
                                    'total de escolas únicas',
                                    'total de municípios únicos'
                                ]
                                if any(x in item['metrica'].lower() for x in metricas_excluir):
                                    continue

                            df_data.append({
                                'Categoria': item['metrica'],
                                'Valor': item['valor'] if item['valor'] is not None else 0,
                                'Percentual': item['percentual'] if item['percentual'] is not None else 0,
                                'Valor_Original': item.get('valor_original', '')
                            })

                        df = pd.DataFrame(df_data)

                        # Para gráficos, remove linhas com valor zero ou nulo
                        df_plot = df[df['Valor'] > 0].copy()

                        if not df_plot.empty:
                            # Determina o nome correto para o título
                            title_category = selected_category
                            if selected_category == 'Indicadores Gerais':
                                title_category = "Indicadores"
                            elif selected_category == 'Status idade-série':
                                title_category = "Status Idade-Série"

                            # Cria gráfico apropriado
                            if selected_category == 'Idade' or 'idade' in selected_category.lower():
                                # Para idade, usa gráfico de linha
                                df_plot['Idade_Num'] = df_plot['Categoria'].str.extract(r'(\d+)').astype(float)
                                df_plot = df_plot.sort_values('Idade_Num')
                                fig = create_line_chart(
                                    df_plot,
                                    title_category,
                                    'Idade_Num',
                                    'Valor',
                                    selected_sheet,
                                    font_sizes
                                )
                                export_height = A4_CONFIG['line_height']
                            else:
                                # Para outras categorias, usa gráfico de barras
                                fig = create_bar_chart(
                                    df_plot,
                                    title_category,
                                    'Categoria',
                                    'Valor',
                                    'Percentual',
                                    selected_sheet,
                                    font_sizes
                                )
                                # Calcula altura para exportação baseada no número de barras
                                num_bars = len(df_plot)
                                export_height = min(
                                    max(A4_CONFIG['bar_height_base'], num_bars * A4_CONFIG['bar_height_per_item']),
                                    A4_CONFIG['bar_height_max']
                                )

                            # Configuração de exportação
                            config = {
                                'toImageButtonOptions': {
                                    'format': 'png',
                                    'filename': f'grafico_{selected_sheet}_{selected_category}'.lower().replace(' ',
                                                                                                                '_'),
                                    'width': A4_CONFIG['width_px'],
                                    'height': export_height,
                                    'scale': 1  # Importante: sem multiplicação
                                },
                                'displaylogo': False
                            }
                            st.plotly_chart(fig, use_container_width=True, config=config)

                            # Instruções para exportação
                            with st.expander("💡 Como exportar para seu documento", expanded=False):
                                st.markdown(f"""
                                **Para salvar o gráfico:**
                                1. Passe o mouse sobre o gráfico
                                2. Clique no ícone 📷 (câmera) no canto superior direito
                                3. O gráfico será baixado com **{A4_CONFIG['width_px']}px de largura**

                                **No Google Docs:**
                                - Insira a imagem e ajuste a largura para **16cm**
                                - A altura será ajustada automaticamente
                                """)

                        # Sempre exibe tabela de dados (incluindo valores N/A)
                        with st.expander("📋 Ver dados tabulares", expanded=(df_plot.empty)):
                            # Formata a tabela com padrão brasileiro
                            df_display = df.copy()

                            # Formatação especial para valores
                            def format_value_display(row):
                                if row['Valor'] == 0 and 'N/A' in str(row.get('Valor_Original', '')):
                                    return "N/A"
                                else:
                                    return format_number_br(row['Valor'])

                            df_display['Valor'] = df_display.apply(format_value_display, axis=1)
                            df_display['Percentual'] = df_display['Percentual'].apply(
                                lambda x: f"{format_number_br(x, True)}%" if x > 0 else "—"
                            )

                            # Remove coluna auxiliar antes de exibir
                            if 'Valor_Original' in df_display.columns:
                                df_display = df_display.drop('Valor_Original', axis=1)
                            if 'Idade_Num' in df_display.columns:
                                df_display = df_display.drop('Idade_Num', axis=1)

                            st.dataframe(df_display, use_container_width=True)

                        # Validação e estatísticas (usando apenas valores válidos)
                        if not df_plot.empty:
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Total da Categoria", format_number_br(df_plot['Valor'].sum()))
                            with col2:
                                if sheet_data['total_matriculas']:
                                    st.metric("Total de Matrículas",
                                              format_number_br(sheet_data['total_matriculas']))
                            with col3:
                                if sheet_data['total_matriculas']:
                                    coverage = (df_plot['Valor'].sum() / sheet_data['total_matriculas']) * 100
                                    st.metric("Cobertura", f"{format_number_br(coverage, True)}%")

                        # Mensagem informativa se não há dados para gráfico
                        if df_plot.empty and not df.empty:
                            st.info(
                                "ℹ️ Esta categoria contém apenas valores N/A ou sem dados numéricos. Veja os detalhes na tabela acima.")
                    else:
                        st.warning("Nenhum dado encontrado para esta categoria.")


if __name__ == "__main__":
    main()