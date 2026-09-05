import streamlit as st
import pandas as pd
import sqlite3
import datetime
import os
import re
import math
import io
import unicodedata
import requests
import pypdf
import folium
from streamlit_folium import st_folium
from streamlit_js_eval import get_geolocation

st.set_page_config(
    page_title="Central RIA - CREL",
    page_icon="🛗",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_NAME = "gestao_ria.db"
ARQUIVO_GEO_BASE = "clientes_geolocalizados.xlsx"
PIN_ACESSO = "2026"


# -----------------------------------------------------------------------------
# 1. CONTROLE DE ACESSO POR PIN
# -----------------------------------------------------------------------------
def autenticar():
    if "auth" not in st.session_state:
        st.session_state["auth"] = False
    if not st.session_state["auth"]:
        st.markdown("""
        <div style="text-align: center; margin-top: 50px; margin-bottom: 20px;">
            <h2 style="color: #1B365D;">🔒 Central de Gestão e Vistorias RIA</h2>
            <p style="color: #666;">CREL Elevadores — Sistema Operacional Integrado</p>
        </div>
        """, unsafe_allow_html=True)
        _, col, _ = st.columns([1, 2, 1])
        with col:
            with st.form("pin_form"):
                pin = st.text_input("PIN de Acesso da Equipe:", type="password", placeholder="••••")
                if st.form_submit_button("Acessar Central", use_container_width=True):
                    if pin == PIN_ACESSO:
                        st.session_state["auth"] = True
                        st.rerun()
                    else:
                        st.error("PIN incorreto. Acesso restrito.")
        st.stop()


autenticar()


# -----------------------------------------------------------------------------
# 2. BANCO DE DADOS & FUNÇÕES GEOGRÁFICAS
# -----------------------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.execute('''
        CREATE TABLE IF NOT EXISTS rotas_ativas (
            id TEXT PRIMARY KEY,
            mes_exec TEXT,
            dia_rota TEXT,
            ordem_parada INTEGER,
            cliente TEXT,
            cep TEXT,
            endereco TEXT,
            qtd_elev INTEGER,
            marca TEXT,
            vencto_ria TEXT,
            grupo_status TEXT,
            lat REAL,
            lon REAL,
            waze_url TEXT,
            status_vistoria TEXT,
            data_baixa TEXT,
            vistoriador TEXT,
            obs TEXT
        )
    ''')
    return conn


def normalizar_codigo(texto):
    """Normaliza o código do cliente removendo espaços, traços, sublinhados e acentos."""
    t = str(texto or '').upper().replace(' ', '').replace('_', '').replace('-', '')
    return ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')


@st.cache_data
def carregar_base_georreferenciada():
    if os.path.exists(ARQUIVO_GEO_BASE):
        df_g = pd.read_excel(ARQUIVO_GEO_BASE)
        df_g['CLI_NORM'] = df_g['CLIENTE'].apply(normalizar_codigo)
        # Importa apenas chave e coordenadas geográficas (sem coluna de supervisor)
        return df_g[['CLI_NORM', 'lat', 'lon']].drop_duplicates(subset=['CLI_NORM'])
    return pd.DataFrame()


def buscar_coords_fallback(cep, endereco):
    cep_limpo = re.sub(r'\D', '', str(cep or ''))
    if len(cep_limpo) == 8:
        try:
            r = requests.get(f"https://cep.awesomeapi.com.br/json/{cep_limpo}", timeout=3)
            if r.status_code == 200:
                d = r.json()
                return float(d['lat']), float(d['lng'])
        except:
            pass
    return -23.5505, -46.6333


def calc_dist_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) ** 2
    return R * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))


# -----------------------------------------------------------------------------
# 3. EXTRATOR AUTOMÁTICO DE PDF DO SIGA
# -----------------------------------------------------------------------------
def extrair_dados_pdf_siga(pdf_bytes):
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    texto_completo = ""
    for page in reader.pages:
        texto_completo += page.extract_text() or ""

    linhas = texto_completo.split('\n')
    registros = []

    start = False
    table_lines = []
    for l in linhas:
        if 'Listados:' in l:
            break
        if '------------------------------------------------------------------------------------------------------------------------------------' in l:
            if not start:
                start = True
                continue
        if start:
            table_lines.append(l)

    for l_str in table_lines:
        l_clean = l_str.strip()
        if not l_clean or 'CLIENTE' in l_clean or '---' in l_clean:
            continue

        m = re.match(r'^([A-Z0-9_\s]{7,9})\s+(\d{8})\s+(.*)$', l_clean)
        if m:
            cli = m.group(1).strip()
            cep = m.group(2).strip()
            resto = m.group(3).strip()

            m_tech = re.search(r'\s+([A-Z])\s+(\d)\s+(\d{1,2})\s+([A-Z0-9/_-]+)\s*(.*)$', resto)
            if m_tech:
                end = resto[:m_tech.start()].strip()
                tec = m_tech.group(1)
                com = m_tech.group(2)
                qtd = int(m_tech.group(3))
                marca = m_tech.group(4)
                datas_str = m_tech.group(5)

                datas = re.findall(r'\d{2}/\d{2}/\d{4}', datas_str)
                if len(datas) >= 3:
                    vencto = datas[2]
                elif len(datas) == 1:
                    vencto = datas[0]
                else:
                    vencto = "Sem Data"

                registros.append({
                    'CLIENTE': cli,
                    'CEP': f"{cep[:5]}-{cep[5:]}",
                    'ENDEREÇO': end,
                    'QTD': qtd,
                    'MARCA': marca,
                    'VENCTO RIA': vencto
                })
    return pd.DataFrame(registros)


# -----------------------------------------------------------------------------
# 4. INTERFACE PRINCIPAL
# -----------------------------------------------------------------------------
menu = st.sidebar.radio("Navegação Principal", [
    "⚙️ 1. Importador & Roteirizador Logístico",
    "📄 2. Relatório de Vistorias e Impressão",
    "📱 3. Terminal de Campo (Vistoriador)"
])

# =============================================================================
# ETAPA 1: IMPORTADOR, FILTROS E ROTEIRIZADOR
# =============================================================================
if menu == "⚙️ 1. Importador & Roteirizador Logístico":
    st.title("⚙️ Gerador Unificado e Roteirizador RIA")
    st.markdown("Importe o relatório (**PDF** ou **Excel**) para gerar jornadas sequenciadas por menor distância.")

    arquivo = st.file_uploader("Carregue o arquivo (PDF ou Excel):", type=["pdf", "xlsx", "xls"])

    if arquivo:
        if arquivo.name.lower().endswith(".pdf"):
            df = extrair_dados_pdf_siga(arquivo.read())
            col_cli = 'CLIENTE'
            col_cep = 'CEP'
            col_end = 'ENDEREÇO'
            col_qtd = 'QTD'
            col_marca = 'MARCA'
            col_venc = 'VENCTO RIA'
            df['QTD_CLEAN'] = df['QTD']
            df['MARCA_CLEAN'] = df['MARCA'].astype(str).str.upper().str.strip()
            df['VENC_CLEAN'] = df['VENCTO RIA']
        else:
            df_raw = pd.read_excel(arquivo)
            linha_cabecalho = None
            for i, row in df_raw.head(15).iterrows():
                valores = [str(v).upper() for v in row.values]
                if any("CLIENTE" in v for v in valores) and any("ENDEREÇO" in v or "ENDERECO" in v for v in valores):
                    linha_cabecalho = i
                    break

            if linha_cabecalho is not None:
                df_raw.columns = df_raw.iloc[linha_cabecalho]
                df = df_raw.iloc[linha_cabecalho + 1:].copy()
            else:
                df = df_raw.copy()

            df.columns = [str(c).strip().upper() for c in df.columns]
            col_cli = next((c for c in df.columns if "CLIENTE" in c), None)
            col_cep = next((c for c in df.columns if "CEP" in c), None)
            col_end = next((c for c in df.columns if "ENDEREÇO" in c or "ENDERECO" in c), None)
            col_qtd = next((c for c in df.columns if "QTDE" in c or "QTD" in c), None)
            col_marca = next((c for c in df.columns if "MARCA" in c), None)
            col_venc = next((c for c in df.columns if "MÊS" in c or "MES" in c or "VENCTO" in c or "VENC" in c), None)

            if not (col_cli and col_end):
                st.error("Não foi possível identificar as colunas de CLIENTE e ENDEREÇO.")
                st.stop()

            df = df[df[col_cli].notna() & (df[col_cli].astype(str).str.strip() != '')].copy()
            df['QTD_CLEAN'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(1).astype(int)
            df['MARCA_CLEAN'] = df[col_marca].astype(str).str.strip().str.upper() if col_marca else 'DIVERSAS'
            df['VENC_CLEAN'] = df[col_venc].astype(str).str.strip() if col_venc else 'Sem Data'

        # Cruzamento direto com coordenadas (desconsiderando supervisor)
        df_geo = carregar_base_georreferenciada()
        df['CLI_NORM'] = df[col_cli].apply(normalizar_codigo)

        if not df_geo.empty:
            df = df.merge(df_geo[['CLI_NORM', 'lat', 'lon']], on='CLI_NORM', how='left')
        else:
            df['lat'] = None
            df['lon'] = None

        # FILTROS OPERACIONAIS
        st.subheader("🎯 Filtros Operacionais")
        f1, f2 = st.columns(2)
        with f1:
            excluir_alfa = st.toggle("Excluir Marca ALFA", value=True)
        with f2:
            modo_venc = st.selectbox(
                "Escopo de Vencimento:",
                ["Todos (Vencimentos + Sem Data / Pendências)", "Apenas Vencimentos com Data",
                 "Apenas Sem Data / Pendências"]
            )

        df_f = df.copy()
        if excluir_alfa:
            df_f = df_f[~df_f['MARCA_CLEAN'].str.contains('ALFA', na=False)]


        def classificar_grupo(v):
            v_str = str(v).lower()
            if "sem data" in v_str or "//" in v_str or not re.search(r'\d{2}/\d{2}/\d{4}', v_str):
                return "Sem Data / Pendência"
            return "Vencimento Programado"


        df_f['GRUPO_STATUS'] = df_f['VENC_CLEAN'].apply(classificar_grupo)

        if modo_venc == "Apenas Vencimentos com Data":
            df_f = df_f[df_f['GRUPO_STATUS'] == "Vencimento Programado"]
        elif modo_venc == "Apenas Sem Data / Pendências":
            df_f = df_f[df_f['GRUPO_STATUS'] == "Sem Data / Pendência"]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Locais Filtrados", len(df_f))
        m2.metric("Com Vencimento", len(df_f[df_f['GRUPO_STATUS'] == 'Vencimento Programado']))
        m3.metric("Sem Data / Pendências", len(df_f[df_f['GRUPO_STATUS'] == 'Sem Data / Pendência']))
        m4.metric("Total de Elevadores", int(df_f['QTD_CLEAN'].sum()))

        # SELEÇÃO MANUAL
        st.markdown("---")
        st.subheader("☑️ Seleção Manual de Endereços")
        df_edit = df_f[[col_cli, col_cep, col_end, 'QTD_CLEAN', 'MARCA_CLEAN', 'VENC_CLEAN', 'GRUPO_STATUS']].copy()
        df_edit.insert(0, "Incluir", True)
        df_edit.columns = ["Incluir", "CLIENTE", "CEP", "ENDEREÇO", "QTD", "MARCA", "VENCTO RIA", "GRUPO"]

        grid_selecao = st.data_editor(df_edit, use_container_width=True, hide_index=True)
        selecionados = grid_selecao[grid_selecao["Incluir"] == True].copy()

        # PARÂMETROS DE ROTA
        st.markdown("---")
        st.subheader("🚚 Configuração das Jornadas Diárias")
        c1, c2 = st.columns(2)
        with c1:
            max_locais_dia = st.slider("Máximo de Locais por Dia:", 3, 15, 7)
        with c2:
            max_elev_dia = st.slider("Máximo de Elevadores por Dia:", 5, 30, 14)

        if st.button("🚀 Otimizar Trajetos e Salvar para Campo", type="primary", use_container_width=True):
            if selecionados.empty:
                st.warning("Nenhum local selecionado.")
            else:
                with st.spinner("Otimizando trajetos por menor distância..."):
                    dados_processados = []
                    for _, r in selecionados.iterrows():
                        norm = normalizar_codigo(r['CLIENTE'])
                        row_orig = df_f[df_f['CLI_NORM'] == norm].iloc[0]
                        p_lat = row_orig['lat']
                        p_lon = row_orig['lon']

                        if pd.isna(p_lat) or p_lat is None:
                            p_lat, p_lon = buscar_coords_fallback(r['CEP'], r['ENDEREÇO'])

                        waze = f"https://waze.com/ul?ll={p_lat},{p_lon}&navigate=yes"
                        dados_processados.append({
                            'cliente': r['CLIENTE'],
                            'cep': str(r['CEP'] or ''),
                            'endereco': r['ENDEREÇO'],
                            'qtd_elev': r['QTD'],
                            'marca': r['MARCA'],
                            'vencto_ria': r['VENCTO RIA'],
                            'grupo_status': r['GRUPO'],
                            'lat': float(p_lat),
                            'lon': float(p_lon),
                            'waze_url': waze
                        })

                    # OTIMIZAÇÃO GLOBAL POR MENOR PERCURSO (SEM DIVISÃO POR SUPERVISOR)
                    rotas_finais = []
                    pontos_restantes = dados_processados.copy()
                    dia_num = 1

                    while pontos_restantes:
                        dia_nome = f"Dia {dia_num}"
                        parada_num = 1
                        elev_acum = 0
                        locais_acum = 0

                        atual = pontos_restantes.pop(0)
                        atual['dia_rota'] = dia_nome
                        atual['ordem_parada'] = parada_num
                        elev_acum += atual['qtd_elev']
                        locais_acum += 1
                        rotas_finais.append(atual)

                        while pontos_restantes:
                            # REGRA RÍGIDA: só adiciona se respeitar AMBOS os limites (locais e elevadores)
                            candidatos = [
                                p
                                for p in pontos_restantes
                                if (locais_acum + 1 <= max_locais_dia)
                                   and (elev_acum + p['qtd_elev'] <= max_elev_dia)
                            ]

                            # Se nenhum outro endereço couber sem estourar o limite, encerra o dia
                            if not candidatos:
                                break

                            proximo = min(
                                candidatos,
                                key=lambda p: calc_dist_km(
                                    atual['lat'],
                                    atual['lon'],
                                    p['lat'],
                                    p['lon'],
                                ),
                            )
                            pontos_restantes.remove(proximo)
                            parada_num += 1
                            locais_acum += 1
                            elev_acum += proximo['qtd_elev']
                            proximo['dia_rota'] = dia_nome
                            proximo['ordem_parada'] = parada_num
                            rotas_finais.append(proximo)
                            atual = proximo
                        dia_num += 1

                    conn = get_db()
                    conn.execute("DELETE FROM rotas_ativas")
                    for p in rotas_finais:
                        uid = f"{p['dia_rota']}_{p['ordem_parada']}_{p['cliente']}".replace(" ", "_")
                        conn.execute('''
                            INSERT INTO rotas_ativas 
                            (id, mes_exec, dia_rota, ordem_parada, cliente, cep, endereco, qtd_elev, marca, vencto_ria, grupo_status, lat, lon, waze_url, status_vistoria)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente')
                        ''', (uid, p['vencto_ria'], p['dia_rota'], p['ordem_parada'], p['cliente'], p['cep'],
                              p['endereco'], p['qtd_elev'], p['marca'], p['vencto_ria'], p['grupo_status'], p['lat'],
                              p['lon'], p['waze_url']))
                    conn.commit()
                    conn.close()

                    st.success(
                        f"✅ Roteirização concluída! {len(rotas_finais)} locais distribuídos em {dia_num - 1} jornadas diárias.")

# =============================================================================
# ETAPA 2: RELATÓRIO DE VISTORIAS E IMPRESSÃO
# =============================================================================
elif menu == "📄 2. Relatório de Vistorias e Impressão":
    conn = get_db()
    df_rel = pd.read_sql("SELECT * FROM rotas_ativas", conn)
    conn.close()

    if not df_rel.empty:
        # Extrai apenas o número do dia para ordenação numérica real (1, 2, 3... 10, 11, 12)
        df_rel['dia_num'] = df_rel['dia_rota'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
        df_rel = df_rel.sort_values(by=['dia_num', 'ordem_parada']).drop(columns=['dia_num']).reset_index(drop=True)

    if df_rel.empty:
        st.info("Nenhuma rota cadastrada no momento. Processe o arquivo na Etapa 1.")
        st.stop()

    # CSS GLOBAL DE IMPRESSÃO QUE DESTRAVA TODAS AS PÁGINAS DO STREAMLIT
    st.markdown("""
    <style>
    @media print {
        /* 1. Destrava alturas e remove barras de rolagem que cortavam o conteúdo */
        html, body, [data-testid="stAppViewContainer"], [data-testid="stMainBlockContainer"], 
        section.main, .main, div[tabindex="0"] {
            overflow: visible !important;
            height: auto !important;
            min-height: 100% !important;
            max-height: none !important;
            position: static !important;
        }

        /* 2. Oculta tudo que não pertence ao relatório */
        header, footer, [data-testid="stSidebar"], [data-testid="stHeader"], 
        .stButton, [data-testid="stDownloadButton"], iframe, .no-print {
            display: none !important;
        }

        /* 3. Margens e dimensionamento A4 Paisagem */
        @page {
            size: A4 landscape;
            margin: 8mm;
        }

        .main .block-container {
            max-width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
        }

        /* 4. Quebra de linha e tabela limpa */
        table {
            width: 100% !important;
            border-collapse: collapse !important;
            page-break-inside: auto !important;
        }
        tr {
            page-break-inside: avoid !important;
            page-break-after: auto !important;
        }
        th, td {
            border: 1px solid #999 !important;
            padding: 4px 6px !important;
            font-size: 11px !important;
        }
        th {
            background-color: #1B365D !important;
            color: white !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }
    }

    /* Estilo da tabela na tela normal */
    .tabela-relatorio {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 20px;
        font-family: Arial, sans-serif;
    }
    .tabela-relatorio th, .tabela-relatorio td {
        border: 1px solid #ddd;
        padding: 6px 8px;
        font-size: 12px;
    }
    .tabela-relatorio th {
        background-color: #1B365D;
        color: white;
        text-align: left;
    }
    .tabela-relatorio tr:nth-child(even) {
        background-color: #f9f9f9;
    }
    </style>
    """, unsafe_allow_html=True)

    # CABEÇALHO DO RELATÓRIO
    st.markdown(f"""
    <div style="text-align: center; border-bottom: 2px solid #1B365D; padding-bottom: 6px; margin-bottom: 12px;">
        <h3 style="margin:0; color:#1B365D;">PROGRAMAÇÃO DE VISTORIAS DE RIA — CREL ELEVADORES</h3>
        <p style="margin:2px 0 0 0; color:#555; font-size:13px;">Relatório Consolidado de Execução e Rotas Diárias</p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Endereços", len(df_rel))
    c2.metric("Total de Elevadores", int(df_rel['qtd_elev'].sum()))
    c3.metric("Dias Programados", df_rel['dia_rota'].nunique())

    # BOTÕES DE EXPORTAÇÃO
    st.markdown('<div class="no-print">', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        buffer = pd.ExcelWriter("Relatorio_Consolidado_RIA.xlsx", engine='openpyxl')
        df_rel.to_excel(buffer, index=False, sheet_name="Roteiro_Geral")
        buffer.close()
        with open("Relatorio_Consolidado_RIA.xlsx", "rb") as f:
            st.download_button("📥 Baixar Relatório em Excel", f, "Relatorio_Consolidado_RIA.xlsx", use_container_width=True)

    with col2:
        import streamlit.components.v1 as components
        components.html("""
        <button onclick="window.parent.print()" style="
            width: 100%;
            height: 42px;
            background-color: #1B365D;
            color: #ffffff;
            border: none;
            border-radius: 8px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            box-shadow: 0 2px 4px rgba(0,0,0,0.15);
        ">
            🖨️ Imprimir / Salvar em PDF (Página Inteira)
        </button>
        """, height=50)
    st.markdown('</div>', unsafe_allow_html=True)


    # FUNÇÃO DE RENDERIZAÇÃO ESTÁTICA EM HTML SEM ESPAÇOS DE INDENTAÇÃO
    def renderizar_tabela_html(df_sub):
        linhas = []
        for _, r in df_sub.iterrows():
            linhas.append(
                f"<tr>"
                f"<td style='text-align:center; font-weight:bold;'>{r['dia_rota']}</td>"
                f"<td style='text-align:center;'>#{r['ordem_parada']}</td>"
                f"<td><b>{r['cliente']}</b></td>"
                f"<td>{r['cep']}</td>"
                f"<td>{r['endereco']}</td>"
                f"<td style='text-align:center;'>{r['qtd_elev']}</td>"
                f"<td>{r['marca']}</td>"
                f"<td style='text-align:center;'>{r['vencto_ria']}</td>"
                f"<td style='text-align:center;'>{r['status_vistoria']}</td>"
                f"</tr>"
            )
        corpo = "".join(linhas)

        html_final = (
            "<table class='tabela-relatorio'>"
            "<thead><tr>"
            "<th style='text-align:center;'>Dia</th>"
            "<th style='text-align:center;'>Parada</th>"
            "<th>Cliente</th>"
            "<th>CEP</th>"
            "<th>Endereço</th>"
            "<th style='text-align:center;'>Qtd</th>"
            "<th>Marca</th>"
            "<th style='text-align:center;'>Vencimento</th>"
            "<th style='text-align:center;'>Status</th>"
            "</tr></thead>"
            f"<tbody>{corpo}</tbody>"
            "</table>"
        )
        return html_final


    # BLOCO 1: VENCIMENTOS CONFIRMADOS
    df_out = df_rel[df_rel['grupo_status'] == 'Vencimento Programado']
    if not df_out.empty:
        st.markdown(
            f"#### 1. Vistorias com Vencimento Programado ({len(df_out)} locais | {int(df_out['qtd_elev'].sum())} elevadores)")
        st.html(renderizar_tabela_html(df_out))

    # BLOCO 2: SEM DATA / PENDÊNCIAS
    df_sem = df_rel[df_rel['grupo_status'] != 'Vencimento Programado']
    if not df_sem.empty:
        st.markdown(
            f"#### 2. Endereços sem Data Cadastrada / Pendências ({len(df_sem)} locais | {int(df_sem['qtd_elev'].sum())} elevadores)")
        st.html(renderizar_tabela_html(df_sem))
# =============================================================================
# ETAPA 3: TERMINAL DE CAMPO (COM MAPA FOLIUM E BAIXA EM TEMPO REAL)
# =============================================================================
elif menu == "📱 3. Terminal de Campo (Vistoriador)":
    st.title("📱 Roteiro de Campo - Vistoriador")

    conn = get_db()
    df_todos = pd.read_sql("SELECT * FROM rotas_ativas", conn)
    conn.close()

    if df_todos.empty:
        st.warning("Nenhuma vistoria atribuída no momento.")
        st.stop()

    dias_disp = sorted(df_todos['dia_rota'].unique(), key=lambda x: int(re.sub(r'\D', '', x) or 0))
    dia_sel = st.selectbox("Selecione a Jornada / Dia:", dias_disp)

    df_dia = df_todos[df_todos['dia_rota'] == dia_sel].sort_values("ordem_parada").copy()

    # GPS do celular
    col_gps1, col_gps2 = st.columns([1, 2])
    with col_gps1:
        loc = get_geolocation()

    gps_ativo = False
    lat_user, lon_user = None, None
    if loc and 'coords' in loc and loc['coords']:
        lat_user = loc['coords']['latitude']
        lon_user = loc['coords']['longitude']
        gps_ativo = True
        with col_gps2:
            st.success(f"📍 GPS Ativo ({lat_user:.4f}, {lon_user:.4f})")

    # MAPA FOLIUM
    with st.expander("🗺️ Ver Mapa do Dia", expanded=True):
        centro_lat = float(df_dia['lat'].mean())
        centro_lon = float(df_dia['lon'].mean())
        mapa = folium.Map(location=[centro_lat, centro_lon], zoom_start=13, tiles="OpenStreetMap")

        coords = []
        if gps_ativo and lat_user and lon_user:
            coords.append([lat_user, lon_user])
            folium.Marker([lat_user, lon_user], popup="Sua Posição", tooltip="Você está aqui",
                          icon=folium.Icon(color="purple", icon="user")).add_to(mapa)

        for _, r in df_dia.iterrows():
            p = [float(r['lat']), float(r['lon'])]
            coords.append(p)
            cor = "green" if r['status_vistoria'] == "Concluído" else (
                "red" if r['status_vistoria'] == "Improdutiva" else "blue")
            folium.Marker(p, popup=f"#{r['ordem_parada']} - {r['cliente']}<br>{r['endereco']}",
                          tooltip=f"#{r['ordem_parada']} {r['cliente']}", icon=folium.Icon(color=cor)).add_to(mapa)

        if len(coords) > 1:
            folium.PolyLine(coords, color="#1B365D", weight=4, dash_array="6, 8").add_to(mapa)

        st_folium(mapa, width="100%", height=350, returned_objects=[])

    st.markdown("---")
    # Cards de Paradas
    for _, row in df_dia.iterrows():
        st_cor = "#28a745" if row['status_vistoria'] == "Concluído" else (
            "#dc3545" if row['status_vistoria'] == "Improdutiva" else "#1B365D")
        st.markdown(f"""
        <div style="background:#fff; border-left:6px solid {st_cor}; padding:12px; border-radius:8px; margin-bottom:12px; box-shadow:0 2px 5px rgba(0,0,0,0.08);">
            <div style="display:flex; justify-content:space-between;">
                <b>PARADA #{row['ordem_parada']} — {row['cliente']}</b>
                <span style="font-size:12px; font-weight:bold; color:{st_cor};">{row['status_vistoria']}</span>
            </div>
            <div style="font-size:13px; color:#555; margin-top:4px;">
                📍 {row['endereco']} | CEP: {row['cep']}<br>
                🛗 {row['qtd_elev']} elevador(es) — Marca: {row['marca']}
            </div>
        </div>
        """, unsafe_allow_html=True)

        b1, b2 = st.columns([1, 1])
        with b1:
            st.link_button("🚗 Navegar no Waze", row['waze_url'], use_container_width=True)
        with b2:
            with st.expander("📝 Baixa da Visita"):
                with st.form(f"baixa_{row['id']}"):
                    n_status = st.selectbox("Status:", ["Concluído", "Improdutiva", "Pendente"], index=0)
                    tecnico = st.text_input("Vistoriador:", value=row['vistoriador'] or "")
                    obs = st.text_area("Observações:", value=row['obs'] or "")
                    if st.form_submit_button("Confirmar Baixa", use_container_width=True):
                        dt = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                        conn = get_db()
                        conn.execute(
                            "UPDATE rotas_ativas SET status_vistoria = ?, data_baixa = ?, vistoriador = ?, obs = ? WHERE id = ?",
                            (n_status, dt, tecnico, obs, row['id']))
                        conn.commit()
                        conn.close()
                        st.success("Baixa realizada!")
                        st.rerun()