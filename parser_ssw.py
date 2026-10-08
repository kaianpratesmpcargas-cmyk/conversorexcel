"""
Módulo de interpretação e limpeza de dados do SSW.
Suporta tanto tabelas delimitadas por barras verticais (|) quanto
tabelas copiadas como texto/colunas separadas por tabulações ou quebras de linha.
"""
import re
import html
from typing import List, Dict, Any, Tuple


# Regex para identificar padrões típicos de CTRC (ex: ETR513286-0, MIM007293-1, EUC513370-0, SSA009990-2, BAR513389-1)
CTRC_PATTERN = re.compile(r'^[A-Z0-9]{2,5}\d{4,8}-\d$', re.IGNORECASE)

# Lista padrão de colunas esperadas e seus sinônimos/normalizações
COLUMN_MAPPINGS = {
    'ctrc': 'CTRC',
    'emissao': 'Emissão',
    'emissão': 'Emissão',
    'remetente': 'Remetente',
    'destinatario': 'Destinatário',
    'destinatário': 'Destinatário',
    'pagador': 'Pagador',
    'destino': 'Destino',
    'n fiscal': 'N Fiscal',
    'nfiscal': 'N Fiscal',
    'nf': 'N Fiscal',
    'nota fiscal': 'N Fiscal',
    'peso': 'Peso',
    'frete': 'Frete',
    'pv entr': 'Pv Entr',
    'pventr': 'Pv Entr',
    'previsao entrega': 'Pv Entr',
    'previsão entrega': 'Pv Entr',
    'ocorrencia': 'Ocorrência',
    'ocorrência': 'Ocorrência',
    'atraso': 'Atraso',
    'localizacao': 'Localização',
    'localização': 'Localização',
}

CANONICAL_COLUMNS = [
    'CTRC', 'Emissão', 'Remetente', 'Destinatário', 'Pagador',
    'Destino', 'N Fiscal', 'Peso', 'Frete', 'Pv Entr',
    'Ocorrência', 'Atraso', 'Localização'
]


def clean_text_cell(text: str) -> str:
    """
    Limpa elementos HTML, Markdown, links e caracteres invisíveis de uma célula.
    """
    if not text:
        return ""

    # Decodifica entidades HTML como &#xA0;, &nbsp;, &amp;, etc.
    s = html.unescape(text)

    # Converte tags <br>, <br/>, <br /> em espaço
    s = re.sub(r'<br\s*/?>', ' ', s, flags=re.IGNORECASE)

    # Remove qualquer outra tag HTML residual
    s = re.sub(r'<[^>]+>', ' ', s)

    # Extrai o texto visível de links Markdown: [Texto Visível](URL) -> Texto Visível
    s = re.sub(r'\[([^\]]*)\]\([^)]*\)', r'\1', s)

    # Remove colchetes residuais de referências ou marcadores Markdown
    s = s.replace('[', '').replace(']', '')

    # Remove negrito/itálico do Markdown (** ou * ou __)
    s = re.sub(r'[*_]{1,3}', '', s)

    # Substitui caracteres invisíveis / non-breaking space
    s = s.replace('\u00a0', ' ').replace('\ufeff', ' ').replace('\u200b', ' ')

    # Substitui quebras de linha ou tabs internas por espaço
    s = re.sub(r'[\r\n\t]+', ' ', s)

    # Normaliza múltiplos espaços consecutivos
    s = re.sub(r'\s+', ' ', s)

    return s.strip()


def parse_number_br(val_str: str) -> float:
    """
    Converte número formatado no padrão brasileiro (ex: 5.230,50 ou 5230 ou 1.250 ou 0,01) para float.
    """
    if not val_str:
        return 0.0

    cleaned = re.sub(r'[^\d,\.-]', '', str(val_str)).strip()
    if not cleaned:
        return 0.0

    if '.' in cleaned and ',' in cleaned:
        cleaned = cleaned.replace('.', '').replace(',', '.')
    elif ',' in cleaned:
        cleaned = cleaned.replace(',', '.')
    elif '.' in cleaned:
        parts = cleaned.split('.')
        if len(parts) == 2 and len(parts[1]) == 3:
            cleaned = cleaned.replace('.', '')
        elif len(parts) > 2:
            cleaned = cleaned.replace('.', '')

    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_atraso(atraso_str: str) -> Tuple[Any, str]:
    """
    Processa o campo atraso e retorna:
    (dias_atraso_int_ou_None, situacao)
    """
    clean_val = clean_text_cell(str(atraso_str))
    if not clean_val or clean_val.lower() in ('-', 'vazio', 'null', 'none', ' '):
        return None, "SEM INFORMAÇÃO"

    match = re.search(r'-?\d+', clean_val)
    if match:
        try:
            dias = int(match.group(0))
            if dias > 0:
                return dias, "ATRASADO"
            else:
                return dias, "NO PRAZO"
        except ValueError:
            pass

    return None, "SEM INFORMAÇÃO"


def normalize_col_name(raw_name: str) -> str:
    """Normaliza o nome da coluna para identificação."""
    cleaned = clean_text_cell(raw_name).lower()
    cleaned = re.sub(r'[^a-z0-9áéíóúãõâêîôûç ]', '', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()

    for key, canon in COLUMN_MAPPINGS.items():
        if key == cleaned or key in cleaned:
            return canon
    return clean_text_cell(raw_name)


def is_valid_ctrc(ctrc_val: str) -> bool:
    """
    Verifica se o valor parece ser um CTRC válido.
    Exemplos: ETR513286-0, MIM007293-1, EUC513370-0, SSA009990-2, BAR513389-1
    """
    if not ctrc_val:
        return False
    
    val = clean_text_cell(ctrc_val).upper()
    if CTRC_PATTERN.match(val):
        return True
    
    if '-' in val and len(val) >= 7:
        parts = val.split('-')
        if len(parts) == 2 and parts[1].isdigit() and len(parts[0]) >= 4:
            return True

    return False


def is_markdown_separator_row(cells: List[str]) -> bool:
    joined = "".join(cells).strip()
    return bool(joined and re.fullmatch(r'[-:\s|]+', joined))


def is_system_footer_or_button(text_line: str) -> bool:
    line_clean = text_line.strip()
    if not line_clean:
        return True
    if re.fullmatch(r'[\s|×xX\-_=]+', line_clean):
        return True
    lower_line = line_clean.lower()
    if any(ignore in lower_line for ignore in ['total geral', 'registros encontrados', 'página', 'pagina', 'fechar']):
        return True
    return False


def parse_pipe_table(raw_text: str) -> List[Dict[str, Any]]:
    """Parser para tabelas formatadas com barras verticais (|)."""
    lines = raw_text.strip().splitlines()
    header_indices: Dict[str, int] = {}
    header_found = False
    data_rows: List[List[str]] = []

    for line_idx, line in enumerate(lines):
        line_str = line.strip()
        if not line_str or '|' not in line_str:
            continue

        raw_cells = [c.strip() for c in line_str.split('|')]
        if raw_cells and raw_cells[0] == "":
            raw_cells.pop(0)
        if raw_cells and raw_cells[-1] == "":
            raw_cells.pop()

        cleaned_cells = [clean_text_cell(c) for c in raw_cells]
        if is_markdown_separator_row(cleaned_cells):
            continue

        norm_cells = [normalize_col_name(c) for c in cleaned_cells]
        if 'CTRC' in norm_cells:
            for idx, col_name in enumerate(norm_cells):
                if col_name and col_name not in header_indices:
                    header_indices[col_name] = idx
            header_found = True

            for data_line in lines[line_idx + 1:]:
                d_str = data_line.strip()
                if not d_str or '|' not in d_str:
                    continue

                d_cells = [c.strip() for c in d_str.split('|')]
                if d_cells and d_cells[0] == "":
                    d_cells.pop(0)
                if d_cells and d_cells[-1] == "":
                    d_cells.pop()

                if is_markdown_separator_row(d_cells) or is_system_footer_or_button(data_line):
                    continue
                if all(not clean_text_cell(c) for c in d_cells):
                    continue

                data_rows.append(d_cells)
            break

    if not header_found or not data_rows:
        return []

    records = []
    for raw_cells in data_rows:
        def get_val(col_name: str) -> str:
            idx = header_indices.get(col_name)
            if idx is not None and idx < len(raw_cells):
                return clean_text_cell(raw_cells[idx])
            return ""

        ctrc_val = get_val('CTRC')
        if not is_valid_ctrc(ctrc_val):
            continue

        peso_num = parse_number_br(get_val('Peso'))
        frete_num = parse_number_br(get_val('Frete'))
        dias_atraso, situacao = parse_atraso(get_val('Atraso'))

        records.append({
            'CTRC': ctrc_val,
            'DATA EMISSÃO': get_val('Emissão'),
            'REMETENTE': get_val('Remetente'),
            'DESTINATÁRIO': get_val('Destinatário'),
            'PAGADOR': get_val('Pagador'),
            'DESTINO': get_val('Destino'),
            'N FISCAL': get_val('N Fiscal'),
            'PESO': peso_num,
            'FRETE': frete_num,
            'PREVISÃO ENTREGA': get_val('Pv Entr'),
            'OCORRÊNCIA': get_val('Ocorrência'),
            'DIAS DE ATRASO': dias_atraso,
            'LOCALIZAÇÃO': get_val('Localização'),
            'SITUAÇÃO': situacao,
        })
    return records


def parse_stream_lines(raw_text: str) -> List[Dict[str, Any]]:
    """
    Parser avançado para quando a tabela do SSW é copiada como texto puro sem pipes (|),
    onde colunas ou linhas foram separadas por quebra de linha ou tabs.
    Detecta cada bloco iniciado por um CTRC válido.
    """
    # Tratar linhas
    lines = raw_text.splitlines()
    cleaned_lines = []
    
    for l in lines:
        c = clean_text_cell(l)
        if not c:
            # Mantém linha vazia como potencial separador ou célula em branco
            cleaned_lines.append("")
        else:
            if not is_system_footer_or_button(c):
                cleaned_lines.append(c)

    # Localizar índices de todas as linhas que contêm um CTRC válido
    ctrc_indices = []
    for idx, line in enumerate(cleaned_lines):
        # Se na linha tiver algo separado por tab
        subparts = [p.strip() for p in line.split('\t') if p.strip()]
        first_token = subparts[0] if subparts else line
        if is_valid_ctrc(first_token):
            ctrc_indices.append((idx, first_token))

    if not ctrc_indices:
        return []

    records = []
    date_regex = re.compile(r'^\d{2}/\d{2}/\d{2,4}$')

    for i, (start_idx, ctrc_code) in enumerate(ctrc_indices):
        end_idx = ctrc_indices[i + 1][0] if (i + 1) < len(ctrc_indices) else len(cleaned_lines)
        block = cleaned_lines[start_idx:end_idx]

        # Se o bloco for tab-delimited em uma linha só
        if len(block) == 1 and '\t' in lines[start_idx]:
            tokens = [clean_text_cell(t) for t in lines[start_idx].split('\t')]
        else:
            # Bloco com quebras de linha verticais
            tokens = [t for t in block]

        # Remover mensagens de erro residuais se existirem
        tokens = [t for t in tokens if "Não foi possível identificar" not in t]

        # Estrutura padrão esperada das 13 colunas:
        # [0] CTRC
        # [1] Emissão (dd/mm/yy)
        # [2] Remetente
        # [3] Destinatário
        # [4] Pagador
        # [5] Destino
        # [6] N Fiscal
        # [7] Peso
        # [8] Frete
        # [9] Pv Entr (dd/mm/yy)
        # [10] Ocorrência
        # [11] Atraso (número)
        # [12] Localização

        # Mapeamento dinâmico inteligente
        emissao = ""
        remetente = ""
        destinatario = ""
        pagador = ""
        destino = ""
        nf = ""
        peso_raw = ""
        frete_raw = ""
        pv_entr = ""
        ocorrencia = ""
        atraso_raw = ""
        localizacao = ""

        if len(tokens) >= 13:
            emissao = tokens[1]
            remetente = tokens[2]
            destinatario = tokens[3]
            pagador = tokens[4]
            destino = tokens[5]
            nf = tokens[6]
            peso_raw = tokens[7]
            frete_raw = tokens[8]
            pv_entr = tokens[9]
            ocorrencia = tokens[10]
            atraso_raw = tokens[11]
            localizacao = " ".join(tokens[12:])
        elif len(tokens) >= 10:
            # Pode haver campos vazios omitidos (ex: Ocorrência vazia gerando 12 itens)
            # Buscar pela data de emissão
            idx_curr = 1
            if idx_curr < len(tokens) and date_regex.match(tokens[idx_curr]):
                emissao = tokens[idx_curr]
                idx_curr += 1

            if idx_curr < len(tokens):
                remetente = tokens[idx_curr]
                idx_curr += 1
            if idx_curr < len(tokens):
                destinatario = tokens[idx_curr]
                idx_curr += 1
            if idx_curr < len(tokens):
                pagador = tokens[idx_curr]
                idx_curr += 1
            if idx_curr < len(tokens):
                destino = tokens[idx_curr]
                idx_curr += 1
            if idx_curr < len(tokens):
                nf = tokens[idx_curr]
                idx_curr += 1
            if idx_curr < len(tokens):
                peso_raw = tokens[idx_curr]
                idx_curr += 1
            if idx_curr < len(tokens):
                frete_raw = tokens[idx_curr]
                idx_curr += 1
            if idx_curr < len(tokens) and date_regex.match(tokens[idx_curr]):
                pv_entr = tokens[idx_curr]
                idx_curr += 1

            # Restante dos tokens: pode conter [Ocorrência, Atraso, Localização] ou [Atraso, Localização]
            remaining = tokens[idx_curr:]
            # Procurar o token que seja puramente numérico (Atraso)
            atraso_found_idx = -1
            for r_idx, r_tok in enumerate(remaining):
                if re.fullmatch(r'\d{1,3}', r_tok.strip()):
                    atraso_found_idx = r_idx
                    atraso_raw = r_tok
                    break

            if atraso_found_idx != -1:
                # O que vem antes do atraso é ocorrência
                ocorrencia = " ".join(remaining[:atraso_found_idx]).strip()
                # O que vem depois do atraso é localização
                localizacao = " ".join(remaining[atraso_found_idx + 1:]).strip()
            else:
                if remaining:
                    localizacao = " ".join(remaining).strip()

        peso_num = parse_number_br(peso_raw)
        frete_num = parse_number_br(frete_raw)
        dias_atraso, situacao = parse_atraso(atraso_raw)

        records.append({
            'CTRC': ctrc_code,
            'DATA EMISSÃO': emissao,
            'REMETENTE': remetente,
            'DESTINATÁRIO': destinatario,
            'PAGADOR': pagador,
            'DESTINO': destino,
            'N FISCAL': nf,
            'PESO': peso_num,
            'FRETE': frete_num,
            'PREVISÃO ENTREGA': pv_entr,
            'OCORRÊNCIA': ocorrencia,
            'DIAS DE ATRASO': dias_atraso,
            'LOCALIZAÇÃO': localizacao,
            'SITUAÇÃO': situacao,
        })

    return records


def parse_ssw_table(raw_text: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    Recebe o texto copiado do SSW (com barras verticais |, Markdown, tabs ou texto puro)
    e retorna os registros processados.
    """
    if not raw_text or not raw_text.strip():
        return [], "Não foi possível identificar a tabela do SSW.\nVerifique se você copiou a tabela completa do SSW."

    # 1. Tentar primeiro via parser delimitado por barras (|)
    if '|' in raw_text:
        records = parse_pipe_table(raw_text)
        if records:
            return records, ""

    # 2. Tentar via parser por stream de linhas/tabs (copiado direto da página sem markdown)
    records = parse_stream_lines(raw_text)
    if records:
        return records, ""

    # Se há menção a CTRC ou cabeçalho mas nenhum registro encontrado
    upper_text = raw_text.upper()
    if 'CTRC' in upper_text and any(k in upper_text for k in ['EMISSÃO', 'EMISSAO', 'DESTINATÁRIO', 'DESTINATARIO']):
        return [], "A tabela contém apenas o cabeçalho ou os registros não possuem CTRC válido."

    return [], "Não foi possível identificar a tabela do SSW.\nVerifique se você copiou a tabela completa do SSW."
