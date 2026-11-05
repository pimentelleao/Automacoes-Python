# Exemplo de portfólio — automação de atendimento e faturamento em ERP
# CÓDIGO ANONIMIZADO: endpoints, identificadores do portal, filas e documentos foram substituídos.
# NÃO execute contra sistemas reais sem autorização e configuração apropriadas.
# A publicação deste arquivo não substitui a autorização do titular do código.

# PORTAL + SAP FIORI / V01 — TESTE 65 — GLOSA AUTOMÁTICA E RETOMADA LEGADA — versão pública
import os
import re
import sys
import time
import unicodedata
import atexit
import zipfile
import shutil
import json
import hashlib
import tempfile
from decimal import Decimal, InvalidOperation
from datetime import datetime
from urllib.parse import urlsplit

import pandas as pd
from playwright.sync_api import sync_playwright, Error as PlaywrightError


# =========================
# CONFIGURAÇÕES
# =========================
# Identificação impressa no terminal para evitar executar uma cópia antiga.
VERSAO_ROBO = "PORTFOLIO-AUTOMACAO-ERP-V01"

# Compatibilidade com atalhos antigos. A conciliação e a consulta de envios
# pendentes agora são automáticas; esta opção não habilita perguntas no terminal.
REVISAR_CONCILIACOES_PENDENTES = '--revisar-conciliacoes' in sys.argv[1:]

# TESTE 37: limites de confirmação, não pausas fixas nem otimização de velocidade.
TEMPO_VF01_TRANSICAO = 30000
TEMPO_VF01_CONFIRMACAO = 45000
TEMPO_VF01_ESTABILIDADE = 1500
# Se uma escrita SAP ficar incerta, preserva a tela e não fatura outra OV.
# Limpo somente após resultado reconhecido; uma nova execução começa vazio.
FIORI_PENDENCIA_GRAVACAO = {}

# O PORTAL e o SAP Fiori serão executados no mesmo Google Chrome usado para
# gravar os locators no Playwright Codegen.
CANAL_NAVEGADOR = "chrome"
NOME_NAVEGADOR = "Google Chrome"

# Pasta oficial de saída no computador principal.
PASTA_BASE_SAIDAS = os.environ.get(
    "PORTAL_PASTA_SAIDAS",
    os.path.join(os.path.expanduser("~"), "automacao_portfolio", "saidas"),
)
PASTA_REQUISICOES = os.path.join(PASTA_BASE_SAIDAS, "requisições")
PASTA_RESUMO = os.path.join(PASTA_BASE_SAIDAS, "resumo")
ID_RODADA = datetime.now().strftime("%Y%m%d_%H%M%S")

# Endereço fixo do SAP Fiori. O próprio ambiente corporativo realiza o
# redirecionamento para o provedor de identidade quando o login for necessário.
FIORI_URL = os.environ.get("ERP_FIORI_URL", "")
PORTAL_URL = os.environ.get("PORTAL_URL", "")
PORTAL_URL_MENU = os.environ.get("PORTAL_URL_MENU", "")
PORTAL_URL_ATENDIMENTO = os.environ.get("PORTAL_URL_ATENDIMENTO", "")

# Seletores de interface específicos devem ser informados apenas em ambiente autorizado.
PORTAL_SELETOR_CATEGORIA = os.environ.get("PORTAL_SELETOR_CATEGORIA", "#categoria-demo")
PORTAL_SELETOR_ASSUNTO = os.environ.get("PORTAL_SELETOR_ASSUNTO", "#campo-assunto")
PORTAL_SELETOR_DESCRICAO = os.environ.get("PORTAL_SELETOR_DESCRICAO", "#campo-descricao")
PORTAL_SELETOR_CENTRO = os.environ.get("PORTAL_SELETOR_CENTRO", "#campo-centro")
PORTAL_SELETOR_EXPANDIR_ARVORE = os.environ.get("PORTAL_SELETOR_EXPANDIR_ARVORE", "#abrir-arvore")
PORTAL_SELETOR_ARVORE_ALTERNATIVO = os.environ.get("PORTAL_SELETOR_ARVORE_ALTERNATIVO", "#abrir-menu")

FRAME_ABERTURA = os.environ.get("PORTAL_FRAME_ABERTURA", "frame[name='inicio']")
FRAME_PRINCIPAL = os.environ.get("PORTAL_FRAME_PRINCIPAL", "frame[name='principal']")
FRAME_PORTAL_MAIN = os.environ.get("PORTAL_FRAME_MAIN", "iframe[name='painel']")
FRAME_DOCUMENTOS = os.environ.get("PORTAL_FRAME_DOCUMENTOS", "iframe[name='documentos']")
FRAME_ARVORE = os.environ.get("PORTAL_FRAME_ARVORE", "iframe[name='arvore']")
FRAME_RESULTADOS = os.environ.get("PORTAL_FRAME_RESULTADOS", "iframe[name='resultados']")
FRAME_DADOS_ADICIONAIS = os.environ.get("PORTAL_FRAME_DADOS_ADICIONAIS", "iframe[name='dadosExtras']")
FRAME_EDITOR_TEXTO = os.environ.get("PORTAL_FRAME_EDITOR_TEXTO", "iframe[title='Editor de texto']")


# O nome do iframe varia conforme o aplicativo aberto no Launchpad.
# Por isso a aplicação é reconhecida pelo conteúdo da tela, não pelo nome do iframe.


# =========================
# LOGIN PORTAL
# =========================
# V19: login manual no navegador.
# O robô abre a tela do PORTAL e aguarda você fazer login manualmente.
# Depois que o PORTAL carregar, volte no PowerShell/VS Code e pressione ENTER.


def aguardar_login_manual_portal(page):
    """
    Aguarda o usuário fazer login manual no PORTAL e valida se a tela principal carregou.
    """

    print("\n========================================")
    print("LOGIN MANUAL NO PORTAL")
    print("========================================")
    print(f"1. No {NOME_NAVEGADOR} que abriu, faça o login manualmente no PORTAL.")
    print("2. Aguarde aparecer a tela principal com o campo 'Consultar requisição...'.")
    print("3. Depois volte neste terminal e pressione ENTER.")

    for tentativa in range(1, 4):
        input(f"\nQuando o PORTAL estiver logado/carregado, pressione ENTER para continuar (tentativa {tentativa}/3)...")

        try:
            frame = (
                page.frame_locator(FRAME_ABERTURA)
                .frame_locator(FRAME_PRINCIPAL)
            )

            busca = frame.get_by_role("textbox", name="Consultar requisição...")
            busca.wait_for(timeout=15000)

            print("[PORTAL] Login manual validado. Campo 'Consultar requisição...' localizado.")
            return True

        except Exception as e:
            print("[PORTAL] Ainda não localizei a tela principal do PORTAL.")
            print(f"Detalhe: {e}")
            print("Se o login ainda não terminou, conclua no navegador e tente novamente.")

    print("\n[ERRO] Não consegui validar o login manual no PORTAL após 3 tentativas.")
    return False


def localizar_campo_pesquisa_fiori_visivel(page):
    """Localiza a pesquisa aberta mesmo quando o SAP omite o nome acessível."""
    campos = [
        (
            "campo 'Procurar'",
            page.get_by_role("searchbox", name=re.compile(r"Procurar|Pesquisar|Pesquisa", re.I)).first,
        ),
        (
            "campo de pesquisa do Launchpad",
            page.get_by_role("searchbox").first,
        ),
        (
            "input de pesquisa do Launchpad",
            page.locator('input[type="search"]').first,
        ),
        (
            "input identificado como pesquisa",
            page.locator(
                'input[placeholder*="procur" i], '
                'input[placeholder*="pesquis" i], '
                'input[aria-label*="procur" i], '
                'input[aria-label*="pesquis" i]'
            ).first,
        ),
    ]

    for descricao, campo in campos:
        try:
            if campo.is_visible(timeout=400):
                return descricao, campo
        except Exception:
            continue

    return "", None


def aguardar_campo_pesquisa_fiori(page, timeout):
    """Aguarda a pesquisa global aberta e devolve o input real."""
    limite = time.time() + (timeout / 1000)
    while time.time() < limite:
        descricao, campo = localizar_campo_pesquisa_fiori_visivel(page)
        if campo is not None:
            return descricao, campo
        page.wait_for_timeout(250)
    return "", None


def localizar_controle_pesquisa_fiori_visivel(page):
    """
    Reconhece o Fiori tanto com a pesquisa fechada quanto já aberta.

    TESTE 33:
    se o campo de pesquisa já estiver aberto, ele TEM precedência sobre o botão
    do cabeçalho. Isso evita tentar clicar em "Abrir pesquisa" através de um
    popup/UI5 já aberto, cenário que gerou o bloqueio
    #sap-ui-blocklayer-popup na rodada do TESTE 32.
    """
    descricao_campo, campo = localizar_campo_pesquisa_fiori_visivel(page)
    if campo is not None:
        return descricao_campo, campo

    controles = [
        (
            "botão de pesquisa",
            page.get_by_role(
                "button",
                name=re.compile(
                    r"Abrir\s+(?:pesquisa|busca)|Pesquisar|Procurar",
                    re.I,
                ),
            ).first,
        ),
        (
            "botão de pesquisa do cabeçalho",
            page.locator(
                'button[title*="pesquis" i], '
                'button[aria-label*="pesquis" i], '
                '[role="button"][title*="pesquis" i], '
                '[role="button"][aria-label*="pesquis" i]'
            ).first,
        ),
    ]

    for descricao, controle in controles:
        try:
            if controle.is_visible(timeout=400):
                return descricao, controle
        except Exception:
            continue

    return "", None

def aguardar_controle_pesquisa_fiori(page, timeout):
    """Aguarda qualquer um dos dois estados autenticados do Launchpad."""
    limite = time.time() + (timeout / 1000)
    while time.time() < limite:
        descricao, controle = localizar_controle_pesquisa_fiori_visivel(page)
        if controle is not None:
            return descricao, controle
        page.wait_for_timeout(250)
    return "", None


def fiori_aparenta_tela_login(page):
    """Diferencia uma tela real de autenticação de uma falha do Shell-home."""
    try:
        url = str(page.url or "").lower()
    except Exception:
        url = ""

    termos_url_login = (
        "login",
        "logon",
        "saml",
        "oauth",
        "authorize",
        "adfs",
        "microsoftonline",
    )
    if any(termo in url for termo in termos_url_login):
        return True

    marcadores = [
        page.locator('input[type="password"]'),
        page.get_by_role(
            "button",
            name=re.compile(r"Entrar|Sign\s*in|Log\s*on", re.I),
        ),
    ]
    for marcador in marcadores:
        try:
            if marcador.first.is_visible(timeout=400):
                return True
        except Exception:
            continue

    return False


def aguardar_login_fiori_sem_enter(page, timeout=300000):
    """Observa o login/SSO no navegador e continua sem exigir ENTER."""

    print("\n========================================")
    print("AUTENTICAÇÃO NO SAP FIORI")
    print("========================================")
    print("1. Na aba do SAP Fiori, conclua o login corporativo/SSO, se solicitado.")
    print("2. Aguarde a página inicial do Fiori carregar.")
    print("3. O robô continuará automaticamente quando reconhecer o Fiori.")

    limite = time.time() + (timeout / 1000)
    while time.time() < limite:
        try:
            descricao, controle = localizar_controle_pesquisa_fiori_visivel(page)
            if controle is not None and not fiori_aparenta_tela_login(page):
                print(f"[FIORI] Login validado automaticamente pelo {descricao}.")
                return True
        except Exception:
            pass

        page.wait_for_timeout(500)

    print("\n[ERRO] O SAP Fiori não ficou disponível dentro do tempo de login.")
    return False



# Pasta base onde serão criadas as pastas de cada requisição
PASTA_ANEXOS_SAP = PASTA_REQUISICOES



def montar_caminho_excel_resultado():
    """Um relatório por rodada, atualizado após cada OV e requisição."""
    return os.path.join(PASTA_RESUMO, f"resultado_extracao_sap_{ID_RODADA}.xlsx")


class SaidaDuplicada:
    """Replica a saída do terminal em um arquivo de log da rodada."""

    def __init__(self, terminal, arquivo):
        self.terminal = terminal
        self.arquivo = arquivo
        self.encoding = getattr(terminal, "encoding", "utf-8")

    def write(self, texto):
        self.terminal.write(texto)
        self.arquivo.write(texto)
        return len(texto)

    def flush(self):
        self.terminal.flush()
        self.arquivo.flush()

    def isatty(self):
        return self.terminal.isatty()

    def fileno(self):
        return self.terminal.fileno()


def iniciar_log_execucao():
    """Cria o log TXT em resumo e mantém as mensagens visíveis no terminal."""
    caminho_log = os.path.join(PASTA_RESUMO, f"log_execucao_{ID_RODADA}.txt")
    arquivo_log = open(caminho_log, "a", encoding="utf-8", buffering=1)
    stdout_original = sys.stdout
    stderr_original = sys.stderr

    sys.stdout = SaidaDuplicada(stdout_original, arquivo_log)
    sys.stderr = SaidaDuplicada(stderr_original, arquivo_log)

    def finalizar_log():
        try:
            sys.stdout.flush()
            sys.stderr.flush()
        finally:
            sys.stdout = stdout_original
            sys.stderr = stderr_original
            arquivo_log.close()

    atexit.register(finalizar_log)
    return caminho_log

# Tempos de espera
TEMPO_PORTAL_REQUISICAO = 6000
TEMPO_FIORI_NAVEGACAO = 60000
TEMPO_FIORI_ELEMENTO = 30000
TEMPO_FIORI_DOWNLOAD = 60000
TEMPO_FIORI_APOS_SALVAR = 3000

# TESTE 33: recuperação de popup/modal UI5 que bloqueia a pesquisa do Fiori.
TEMPO_FIORI_RECUPERACAO_POPUP = 12000
TEMPO_FIORI_CARREGAMENTO_BLOQUEANTE = 60000

# Status técnico usado quando a linha da fatura no Fluxo de documentos estiver com erro contábil.
STATUS_ERRO_DETERMINACAO_CONTA = "ERRO_DETERMINACAO_CONTA"

# Status específico do novo processo quando a Nota de Crédito continua pendente
# no Fluxo de documentos mesmo após a execução obrigatória da VF01.
STATUS_NOTA_CREDITO_PENDENTE = "NOTA_CREDITO_PENDENTE"

# =========================
# CONFIGURAÇÕES DO NOVO PROCESSO - SUPORTE A VENDA
# =========================

FILA_PORTAL_SUPORTE_VENDA = os.environ.get("PORTAL_FILA", "Equipe de Operações - Faturamento")
SERVICO_CONCILIACAO = os.environ.get("PORTAL_SERVICO_CONCILIACAO", "Conciliação de Crédito")
PREFIXO_DOC_CREDITO = os.environ.get("ERP_PREFIXO_DOC_CREDITO", "DEMO")
ATIVIDADE_NOTA_CREDITO = "Nota de Credito"
ATIVIDADE_DEVOLUCAO_REFATURAMENTO = "Devolução e Refaturamento (Locação)"

# O mesmo motor SAP é usado para as duas atividades. O que muda na VA03 é
# somente a linha de faturamento esperada no Fluxo de documentos.
CONFIG_ATIVIDADES_SUPORTE_VENDA = {
    ATIVIDADE_NOTA_CREDITO: {
        # A atividade Nota de Credito já apresentou mais de uma nomenclatura
        # no Fluxo de documentos da VA03. Casos confirmados:
        #   - "Fat Nota de crédito DEMO DOC_EXEMPLO_01 / 10"
        #   - "Fatura - Nota de Crédito DOC_EXEMPLO_02 / 10"
        # O número do documento/item vem depois do prefixo; por isso a busca
        # aceita ambos os prefixos e não exige igualdade textual completa.
        "documentos_va03": [
            rf"^\s*Fat\s+Nota\s+de\s+cr[eé]dito\s+{re.escape(PREFIXO_DOC_CREDITO)}\b",
            r"^\s*Fatura\s*-\s*Nota\s+de\s+Cr[eé]dito\b",
        ],
        "descricao_documento": (
            "Fat Nota de crédito DEMO / Fatura - Nota de Crédito"
        ),
    },
    ATIVIDADE_DEVOLUCAO_REFATURAMENTO: {
        # A atividade de Devolução e Refaturamento (Locação) apresentou no SAP
        # mais de uma nomenclatura real. O TESTE 43 acrescenta a variação LO Leg
        # observada em produção, preservando os formatos já validados.
        # Casos confirmados:
        #   - "Fat NC Loca Mensal DOC_EXEMPLO_03 / 10"
        #   - "Fat Locação Mensal DOC_EXEMPLO_04 / 10"
        #   - "Fatura - NC Locação (exemplo) DOC_EXEMPLO_05 / 10"
        "documentos_va03": [
            r"^\s*Fat\s+NC\s+Loca\s+Mensal\b",
            r"^\s*Fat\s+Loca[cç][aã]o\s+Mensal\b",
            r"^\s*Fat\s+Loca[cç][aã]o\s+Trimestral\b",
            r"^\s*Fatura\s*-\s*NC\s+Loca[cç][aã]o(?:\s+\w+){0,2}\b",
        ],
        "descricao_documento": (
            "Fat NC Loca Mensal / Fat Locação Mensal / Fat Locação Trimestral / "
            "Fatura - NC Locação (exemplo)"
        ),
    },
}

# Mantém a atividade vinculada ao número da requisição enquanto a fila está
# sendo processada. A seleção manual continua trabalhando somente com números.
ATIVIDADE_POR_REQUISICAO = {}
PAGINA_POR_REQUISICAO = {}

# Limite de segurança para impedir loop infinito em grids paginados do PORTAL.
MAX_PAGINAS_PORTAL = 100

# =========================
# CONFIGURAÇÕES PORTAL - ANEXO / ENCERRAMENTO
# =========================

# Login manual no PORTAL e dialogs inteligentes sem duplicar anexos.
# Operação completa validada: sem erro de determinação de conta e com todos os
# PDFs válidos, o robô assume a requisição, anexa os arquivos e conclui o PORTAL.
ENCERRAR_PORTAL_APOS_PDF = True

# O único impedimento de negócio atual é "Erro na determ. conta". Esta proteção
# adicional não cria uma regra SAP: ela apenas evita concluir o PORTAL se uma falha
# técnica impedir que algum PDF seja efetivamente salvo.

# Campos de Dados Adicionais da fila Equipe de Operações - Faturamento.
# Ambos recebem a quantidade total de documentos apenas quando TODAS as OVs
# da requisição foram concluídas com PDF válido.

TEXTO_ENCERRAMENTO_PADRAO = "Seguem as notas fiscais finalizadas com êxito."

TEMPO_PORTAL_ACAO = 3000
TEMPO_PORTAL_UPLOAD_PRONTO = 60000
TEMPO_PORTAL_CONFIRMACAO_ANEXO = 75000
TEMPO_PORTAL_CONFERENCIA_ANEXO = 30000

# Estado compartilhado do handler de dialogs do PORTAL.
# Usado principalmente para detectar quando o PORTAL informa que o arquivo já existe
# e evitar anexar a mesma nota como nova versão/duplicidade.
PORTAL_DIALOG_STATE = {
    "arquivo_ja_anexado": False,
    "ultimo_dialog": "",
    "historico": [],
    "eventos": [],
}


# =========================
# FUNÇÕES DE ENTRADA
# =========================


def classificar_atividade_suporte_venda(texto):
    """
    Retorna somente uma das duas atividades permitidas.

    Regra crítica V01 TESTE 02:
    - "Devolução e Refaturamento (Produto)" NÃO pertence a este robô;
    - somente "Devolução e Refaturamento (Locação)" pode entrar;
    - a comparação é feita pelo nome completo da atividade, não pelo prefixo
      "Devolução e Refaturamento", porque Produto e Locação são muito parecidos.
    """
    txt = normalizar_texto_busca(texto)

    # Bloqueio explícito da atividade semelhante que está fora do escopo.
    if re.search(r"devolucao\s+e\s+refaturamento\s*\(\s*produto\s*\)", txt, re.I):
        return ""

    if re.search(r"devolucao\s+e\s+refaturamento\s*\(\s*locacao\s*\)", txt, re.I):
        return ATIVIDADE_DEVOLUCAO_REFATURAMENTO

    if re.search(r"(?<![a-z])nota\s+de\s+credito(?![a-z])", txt, re.I):
        return ATIVIDADE_NOTA_CREDITO

    return ""


def extrair_ovs_suporte_venda(texto):
    """Lê listas de 7–10 dígitos após OV/Ordem de Venda; para em outros rótulos."""
    linhas = str(texto or "").splitlines()
    encontrados = []
    rotulo = re.compile(r"\b(?:OV(?:['’´]?s)?|(?:Ordem|Ordens)\s+(?:de\s+)?Venda)\b", re.I)
    numero = r"(?<!\d)\d{7,10}(?!\d)"
    lista = rf"{numero}(?:(?:\s*[/,;]\s*|\s+e\s+|\s+){numero})*"
    prefixo = r"\s*(?:(?:n[º°o.]?|n[uú]mero)\s*)?[:#\-]?\s*"
    for indice, linha in enumerate(linhas):
        for contexto in rotulo.finditer(linha):
            trecho = linha[contexto.end():]
            inline = re.match(prefixo + f"({lista})", trecho, re.I)
            if inline:
                encontrados.extend(re.findall(numero, inline.group(1)))
                continue
            # Lista vertical precisa de cabeçalho explícito e linhas numéricas.
            # Um bloco anterior já encontrado não encerra este bloco.
            if not (re.fullmatch(prefixo, trecho, re.I) or re.search(
                    r"\b(?:abaixo|a\s+seguir|seguintes?|relacionadas?|informadas?)\b", trecho, re.I)):
                continue
            iniciou = False
            for proxima in linhas[indice + 1:]:
                proxima = proxima.strip()
                if not proxima:
                    continue
                if not iniciou and re.fullmatch(r"segue[m]?\s*:?", proxima, re.I):
                    continue
                if not re.fullmatch(lista, proxima, re.I):
                    break
                encontrados.extend(re.findall(numero, proxima))
                iniciou = True
    return list(dict.fromkeys(encontrados))


def extrair_candidatos_ov_descricao(texto, atividade):
    """Layouts adicionais observados; cada candidato exige VA03 antes da VF01.

    Não procura números soltos por toda a descrição. Aceita somente um bloco
    completo: clientes + OV após cabeçalho explícito, NF/OV/centro de Locação,
    ou uma única OV em linha isolada após pedido de devolução/emissão.
    """
    tipo = classificar_atividade_suporte_venda(atividade)
    if tipo not in {ATIVIDADE_NOTA_CREDITO, ATIVIDADE_DEVOLUCAO_REFATURAMENTO}:
        return []
    linhas = [re.sub(r'\s+', ' ', l).strip() for l in str(texto or '').replace('\xa0', ' ').splitlines() if l.strip()]
    if len(linhas) < 2:
        return []
    numeros = [i for i, l in enumerate(linhas) if re.search(r'\d', l)]
    if not numeros:
        return []
    inicio = numeros[0]
    cabecalho = normalizar_texto_busca(' '.join(linhas[:inicio]))
    dados = linhas[inicio:]
    # A saudação/cabeçalho não contém outros números; nenhuma linha numérica
    # inválida do bloco é descartada silenciosamente.
    if inicio == 0:
        return []
    if re.search(r'\bovs?\b.*\b(?:abaixo|seguintes|a seguir)\b', cabecalho):
        ovs = []
        for linha in dados:
            m = re.fullmatch(r'(?P<cliente>[^\d]+?)\s+(?P<ov>60\d{6})', linha)
            if not m or not any(c.isalpha() for c in m['cliente']):
                return []
            if re.search(r'\b(?:bp|cnpj|cpf|nf|fatura|faturamento|documento|centro)\b',
                         normalizar_texto_busca(m['cliente'])):
                return []
            ovs.append(m['ov'])
        return list(dict.fromkeys(ovs))
    if (tipo == ATIVIDADE_DEVOLUCAO_REFATURAMENTO
            and re.search(r'\b(?:devolver|devolucao)\b', cabecalho)
            and re.search(r'\blocacao\b', cabecalho)):
        ovs = []
        for linha in dados:
            m = re.fullmatch(r'0?10\d{7}\s+(60\d{6})\s+\d{4}', linha)
            if not m:
                ovs = []
                break
            ovs.append(m.group(1))
        if ovs:
            return list(dict.fromkeys(ovs))
    if (len(dados) == 1 and re.fullmatch(r'60\d{6}', dados[0])
            and re.search(r'\b(?:gentileza|solicito|seguir|favor)\b', cabecalho)
            and re.search(r'\b(?:devolucao|devolver|emissao|emitir|nota de credito)\b', cabecalho)):
        return dados
    return []


def extrair_ovs_tabela_descricao(texto):
    """Reconhece o layout observado BP / razão social / faturamento / OV.

    Sem cabeçalho OV, exige pelo menos duas linhas completas e o pedido de
    devolução/crédito. Somente o último campo 60xxxxxx é candidato, nunca BP
    ou NF. Cada candidato ainda deve ser confirmado por consulta VA03 antes
    de qualquer VF01. Linhas numéricas fora desse layout não são aceitas.
    """
    normal = normalizar_texto_busca(texto)
    if not re.search(r"\b(?:devolver|devolucao|credito|refaturamento)\b", normal):
        return []
    padrao = re.compile(
        r"^\s*10\d{6}\s+(?P<cliente>[^\r\n]+?)\s+"
        r"0?10\d{7}\s+(?P<ov>60\d{6})\s*$"
    )
    ovs = []
    for linha in str(texto or "").splitlines():
        if not re.search(r"(?<!\d)\d{7,10}(?!\d)", linha):
            continue
        m = padrao.fullmatch(linha)
        # Não aproveita só uma parte de uma tabela mista/ambígua.
        if not m or not any(c.isalpha() for c in m.group("cliente")):
            return []
        ovs.append(m.group("ov"))
    return list(dict.fromkeys(ovs)) if len(ovs) >= 2 else []



# =========================
# FALLBACK DE OV EM PLANILHA ANEXADA AO PORTAL
# =========================

EXTENSOES_PLANILHA_PORTAL = {".xlsx", ".xlsm", ".xlsb", ".xls"}

# Abas tipicamente cadastrais/de apoio. Não devem participar do fallback por
# posição (D/E), pois costumam conter muitos códigos numéricos que não são OVs.
ABAS_EXCEL_IGNORADAS = {
    "base materiais",
    "materiais",
    "emp",
    "base ordens",
    "inicio",
}

PALAVRAS_ABA_OPERACIONAL = (
    "nota",
    "credito",
    "dev",
    "devol",
    "rec",
    "refat",
    "fat",
    "loca",
    "solicit",
)


# TESTE 15:
# Nos modelos observados, a aba abaixo concentra a OV efetivamente usada no
# fluxo Nota de Crédito / Devolução-Recusa. Se ela existir e produzir OVs
# válidas, o robô NÃO mistura números encontrados em outras abas do template.
TOKENS_ABA_PRIORITARIA_PADRAO = {
    "somente",
    "dev",
    "rec",
    "nota",
    "cred",
    "sap",
}

def _normalizar_nome_excel(valor):
    """Normaliza textos de planilha para comparação sem acentos."""
    if valor is None:
        return ""
    return normalizar_texto_busca(str(valor)).strip().lower()


def _valor_excel_para_ov(valor):
    """
    Converte um valor de célula em possível OV.
    Aceita inteiros/float inteiros e strings com 7 a 10 dígitos.
    """
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    if isinstance(valor, bool):
        return ""

    if isinstance(valor, int):
        candidato = str(valor)
    elif isinstance(valor, float):
        if not float(valor).is_integer():
            return ""
        candidato = str(int(valor))
    else:
        candidato = str(valor).strip()
        # Excel às vezes serializa número inteiro como "1234567.0".
        if re.fullmatch(r"\d{7,10}\.0+", candidato):
            candidato = candidato.split(".", 1)[0]

    candidato = candidato.strip()
    if re.fullmatch(r"\d{7,10}", candidato):
        return candidato
    return ""


def _celula_e_rotulo_ov_excel(valor):
    """
    Identifica rótulos de OV com contexto suficiente.
    Evita considerar qualquer frase genérica que apenas mencione "ordens".
    """
    txt = re.sub(r'\s+', ' ', _normalizar_nome_excel(valor)).strip()
    if not txt:
        return False

    # Tipo da ordem e observações mencionam OV, mas NÃO identificam o campo
    # que contém seu número. Não usar esses rótulos para extrair documentos.
    if re.search(r"\b(?:tipo|observacao|observacoes|informacoes adicionais)\b", txt):
        return False

    # Cabeçalho de número, não uma instrução livre como "Conferir OV anterior".
    cabecalho = (r'(?:informar\s+)?(?:(?:n[º°o.]|numero)\s*(?:da\s+|de\s+)?)?'
                 r'(?:ovs?|(?:ordem|ordens)\s+(?:de\s+)?vendas?)'
                 r'(?:\s*\(sap\))?\s*:?')
    if re.fullmatch(cabecalho, txt):
        return True
    # Mantém o formato "OV: 608..." na própria célula; a rotina da aba dá
    # precedência ao cabeçalho de coluna quando existe um formulário tabular.
    return bool(extrair_ovs_suporte_venda(str(valor)))


def _aba_excel_deve_ser_ignorada(nome_aba):
    nome = _normalizar_nome_excel(nome_aba)
    return nome in ABAS_EXCEL_IGNORADAS


def _aba_excel_parece_operacional(nome_aba):
    nome = _normalizar_nome_excel(nome_aba)
    if not nome or _aba_excel_deve_ser_ignorada(nome):
        return False
    return any(p in nome for p in PALAVRAS_ABA_OPERACIONAL)


def _tokens_nome_aba_excel(nome_aba):
    nome = _normalizar_nome_excel(nome_aba)
    return set(re.findall(r"[a-z0-9]+", nome))


def _abas_excel_prioritarias(nomes_abas, atividade=""):
    """
    Escolhe abas com maior probabilidade de conter a OV real do processo.

    Regra principal validada nos arquivos de teste:
      "Somente Dev_Rec_Nota Créd - SAP"

    A mesma aba já apareceu em requisições de Nota de Credito e também em
    Devolução e Refaturamento (Locação), portanto ela tem precedência para
    ambas as atividades.
    """
    nomes_abas = list(nomes_abas or [])
    prioritarias = []

    for aba in nomes_abas:
        if _aba_excel_deve_ser_ignorada(aba):
            continue
        tokens = _tokens_nome_aba_excel(aba)
        if TOKENS_ABA_PRIORITARIA_PADRAO.issubset(tokens):
            prioritarias.append(aba)

    if prioritarias:
        return prioritarias

    # Fallback por afinidade de nome. É usado somente se a aba padrão acima
    # não existir no workbook.
    atividade = str(atividade or "").strip()

    if atividade == ATIVIDADE_NOTA_CREDITO:
        for aba in nomes_abas:
            if _aba_excel_deve_ser_ignorada(aba):
                continue
            tokens = _tokens_nome_aba_excel(aba)
            if "nota" in tokens and ("cred" in tokens or "credito" in tokens):
                prioritarias.append(aba)

    elif atividade == ATIVIDADE_DEVOLUCAO_REFATURAMENTO:
        for aba in nomes_abas:
            if _aba_excel_deve_ser_ignorada(aba):
                continue
            tokens = _tokens_nome_aba_excel(aba)
            if "loca" in tokens or "locacao" in tokens:
                prioritarias.append(aba)

    return list(dict.fromkeys(prioritarias))


def _ler_planilha_dataframe(caminho, aba):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in {".xlsx", ".xlsm"}:
        return pd.read_excel(
            caminho,
            sheet_name=aba,
            header=None,
            dtype=object,
            engine="openpyxl",
        )
    if ext == ".xlsb":
        return pd.read_excel(
            caminho,
            sheet_name=aba,
            header=None,
            dtype=object,
            engine="pyxlsb",
        )
    if ext == ".xls":
        # Pandas utiliza xlrd para .xls antigo. Se não estiver instalado,
        # o erro é tratado no chamador sem derrubar a fila inteira.
        return pd.read_excel(
            caminho,
            sheet_name=aba,
            header=None,
            dtype=object,
            engine="xlrd",
        )
    raise ValueError(f"Extensão de planilha não suportada: {ext}")


def _abrir_excel_file(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in {".xlsx", ".xlsm"}:
        return pd.ExcelFile(caminho, engine="openpyxl")
    if ext == ".xlsb":
        return pd.ExcelFile(caminho, engine="pyxlsb")
    if ext == ".xls":
        return pd.ExcelFile(caminho, engine="xlrd")
    raise ValueError(f"Extensão de planilha não suportada: {ext}")


def extrair_ovs_de_planilha(caminho, atividade=""):
    """
    Extrai OVs de um arquivo Excel anexado ao PORTAL.

    TESTE 15:
    - Se existir a aba padrão "Somente Dev_Rec_Nota Créd - SAP" (ou uma aba
      prioritária equivalente), ela é analisada PRIMEIRO.
    - Se essa aba produzir uma ou mais OVs seguras, somente essas OVs são
      devolvidas. Números de outras abas do mesmo template não são misturados.
    - TESTE 61: um cabeçalho de OV define a coluna exclusiva de leitura.
      Tipo da ordem e observações não fornecem OVs. Se o formulário padrão
      estiver sem OV válida, não usar números de outra coluna/aba.
    - Sem formulário reconhecido na aba prioritária, permanece a busca nas
      demais abas operacionais.

    Retorna (ovs, origens), onde origens mapeia OV -> lista de referências.
    """
    caminho = os.path.abspath(caminho)
    ovs = []
    origens = {}
    abas_com_estrutura_ov = set()

    try:
        excel = _abrir_excel_file(caminho)
    except ImportError as e:
        print(
            f"[PORTAL EXCEL] Dependência ausente para ler {os.path.basename(caminho)}: {e}. "
            "Para .xlsb, instale 'pyxlsb'; para .xls antigo, instale 'xlrd'."
        )
        return [], {}
    except Exception as e:
        print(
            f"[PORTAL EXCEL] Não consegui abrir {os.path.basename(caminho)}: {e}"
        )
        return [], {}

    def registrar(ov, aba, linha, coluna, metodo):
        if not ov:
            return
        ref = (
            f"{os.path.basename(caminho)} | {aba}!R{linha}C{coluna} | {metodo}"
        )
        origens.setdefault(ov, [])
        if ref not in origens[ov]:
            origens[ov].append(ref)
        if ov not in ovs:
            ovs.append(ov)

    def varrer_semantico(abas):
        resultados = []

        for aba in abas:
            if _aba_excel_deve_ser_ignorada(aba):
                print(f"[PORTAL EXCEL] Aba ignorada por segurança: {aba}")
                continue
            try:
                df = _ler_planilha_dataframe(caminho, aba)
            except Exception as e:
                print(f"[PORTAL EXCEL] Não consegui ler a aba '{aba}' de "
                      f"{os.path.basename(caminho)}: {e}")
                continue
            if df is None or df.empty:
                continue

            max_linhas = min(len(df.index), 500)
            max_colunas = min(len(df.columns), 60)
            campos, valores_rotulados = [], []
            formulario = False
            for r in range(max_linhas):
                for c in range(max_colunas):
                    valor = df.iat[r, c]
                    texto = _normalizar_nome_excel(valor)
                    if (re.search(r'\btipo\s+(?:da\s+)?(?:ordem|ordens)\s+(?:de\s+)?vendas?\b', texto)
                            or ('observacao' in texto and 'informacoes adicionais' in texto)):
                        formulario = True
                    if not _celula_e_rotulo_ov_excel(valor):
                        continue
                    numeros = extrair_ovs_suporte_venda(str(valor))
                    if numeros:
                        valores_rotulados.append((r, c, numeros))
                    else:
                        campos.append((r, c, []))

            if campos or formulario:
                abas_com_estrutura_ov.add(aba)
            # Com coluna explícita, não aceitar "OV ..." citado nas observações.
            # Tipo/Observação sem cabeçalho de número não autorizam fallback.
            rotulos = campos if campos else ([] if formulario else valores_rotulados)
            for r, c, numeros_no_rotulo in rotulos:
                for numero in numeros_no_rotulo:
                    resultados.append((numero, aba, r + 1, c + 1, "rótulo/contexto"))
                if numeros_no_rotulo:
                    continue

                # Coluna exclusiva. Em D:E mescladas, o cabeçalho e os valores
                # ficam na âncora D; não coletar valores das colunas vizinhas.
                linhas_vazias_seguidas = 0
                for rr in range(r + 1, max_linhas):
                    valor = df.iat[rr, c]
                    ov = _valor_excel_para_ov(valor)
                    if ov:
                        resultados.append((ov, aba, rr + 1, c + 1,
                                           "campo associado ao rótulo; coluna exclusiva da OV"))
                        linhas_vazias_seguidas = 0
                        continue
                    try:
                        vazio = bool(pd.isna(valor)) or not str(valor).strip()
                    except Exception:
                        vazio = valor is None
                    if vazio:
                        linhas_vazias_seguidas += 1
                        if linhas_vazias_seguidas >= 2:
                            break
                        continue
                    # Outro texto/cabeçalho encerra esta coluna de dados.
                    break

        return resultados

    def varrer_fallback_de(abas):
        candidatos = []

        for aba in abas:
            if aba in abas_com_estrutura_ov:
                print(f"[PORTAL EXCEL] {aba}: formulário de OV identificado, mas sem número válido; "
                      "fallback por posição bloqueado.")
                continue
            if not _aba_excel_parece_operacional(aba):
                continue

            try:
                df = _ler_planilha_dataframe(caminho, aba)
            except Exception:
                continue

            if df is None or df.empty:
                continue

            max_linhas = min(len(df.index), 500)
            for r in range(max_linhas):
                # D e E -> índices 3 e 4.
                for c in (3, 4):
                    if c >= len(df.columns):
                        continue
                    ov = _valor_excel_para_ov(df.iat[r, c])
                    if ov:
                        candidatos.append(
                            (ov, aba, r + 1, c + 1, "fallback D/E")
                        )

        return candidatos

    try:
        abas_todas = list(excel.sheet_names)
        abas_prioritarias = _abas_excel_prioritarias(
            abas_todas,
            atividade=atividade,
        )

        if abas_prioritarias:
            print(
                "[PORTAL EXCEL] Aba(s) prioritária(s) para extração de OV: "
                + ", ".join(abas_prioritarias)
            )

            resultados_prioritarios = varrer_semantico(abas_prioritarias)

            # Se o rótulo semântico não foi detectado, ainda tentamos D/E apenas
            # dentro da(s) aba(s) prioritária(s).
            if not resultados_prioritarios:
                resultados_prioritarios = varrer_fallback_de(abas_prioritarias)

            for item in resultados_prioritarios:
                registrar(*item)

            if ovs:
                print(
                    "[PORTAL EXCEL] OV(s) encontrada(s) na aba prioritária; "
                    "demais abas do template não serão usadas nesta requisição."
                )
                return ovs, origens

            if any(aba in abas_com_estrutura_ov for aba in abas_prioritarias):
                print('[PORTAL EXCEL] Formulário prioritário sem OV válida na coluna própria. '
                      'Números de outras colunas/abas não serão usados.')
                return [], {}

            print(
                "[PORTAL EXCEL] Nenhuma OV segura foi localizada na aba prioritária; "
                "ampliando a busca para as demais abas operacionais."
            )

        # Busca ampliada somente quando a aba prioritária não resolveu.
        abas_restantes = [
            aba for aba in abas_todas
            if aba not in set(abas_prioritarias)
            and not _aba_excel_deve_ser_ignorada(aba)
        ]

        resultados_semanticos = varrer_semantico(abas_restantes)
        for item in resultados_semanticos:
            registrar(*item)

        # Fallback D/E apenas se nenhuma OV semântica foi encontrada.
        if not ovs:
            candidatos_fallback = varrer_fallback_de(abas_restantes)
            unicos_fb = list(dict.fromkeys(x[0] for x in candidatos_fallback))

            if len(unicos_fb) > 30:
                print(
                    f"[PORTAL EXCEL] Fallback D/E encontrou {len(unicos_fb)} números; "
                    "resultado descartado por segurança."
                )
            else:
                for item in candidatos_fallback:
                    registrar(*item)

        return ovs, origens

    finally:
        try:
            excel.close()
        except Exception:
            pass


def extrair_ovs_de_planilhas(caminhos, atividade=""):
    """TESTE 46: examina em ordem e para no primeiro Excel que fornece OVs."""
    for caminho in caminhos:
        try:
            ovs, origens = extrair_ovs_de_planilha(caminho, atividade=atividade)
        except Exception as erro:
            print(f"[PORTAL EXCEL] Falha ao extrair OVs de {os.path.basename(caminho)}: {erro}. Tentando próximo candidato.")
            continue
        if ovs:
            print(f"[PORTAL EXCEL] {os.path.basename(caminho)}: {len(ovs)} OV(s) localizada(s); busca de Excel resolvida.")
            return ovs, origens
        print(f"[PORTAL EXCEL] {os.path.basename(caminho)}: zero OVs extraídas; tentando próximo candidato, se houver.")
    return [], {}

def _nome_arquivo_seguro(nome):
    nome = str(nome or "").strip()
    nome = re.sub(r'[<>:"/\\|?*]+', "_", nome)
    nome = nome.rstrip(" .")
    return nome or f"planilha_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


def _planilha_local_estruturalmente_valida(caminho):
    """Valida rapidamente uma planilha já existente na pasta da requisição."""
    try:
        if not caminho or not os.path.isfile(caminho):
            return False
        ext = os.path.splitext(caminho)[1].lower()
        if ext not in EXTENSOES_PLANILHA_PORTAL or os.path.getsize(caminho) <= 0:
            return False
        if ext in {".xlsx", ".xlsm", ".xlsb"}:
            if not zipfile.is_zipfile(caminho):
                return False
            with zipfile.ZipFile(caminho, "r") as zf:
                nomes = set(zf.namelist())
            if ext in {".xlsx", ".xlsm"}:
                return {"[Content_Types].xml", "xl/workbook.xml"}.issubset(nomes)
            return any(nome.lower().endswith("workbook.bin") for nome in nomes)
        if ext == ".xls":
            with open(caminho, "rb") as arq:
                return arq.read(8) == bytes.fromhex("D0CF11E0A1B11AE1")
    except Exception:
        return False
    return False


def localizar_planilhas_locais_requisicao(pasta_requisicao):
    """Reaproveita Excel válido já salvo na pasta da própria requisição."""
    pasta_requisicao = os.path.abspath(str(pasta_requisicao or "").strip())
    if not os.path.isdir(pasta_requisicao):
        return []

    candidatos = []
    for raiz, _pastas, arquivos in os.walk(pasta_requisicao):
        for nome in arquivos:
            if nome.startswith((".__tmp_excel_", "~$")):
                continue
            ext = os.path.splitext(nome)[1].lower()
            if ext not in EXTENSOES_PLANILHA_PORTAL:
                continue
            caminho = os.path.join(raiz, nome)
            if not _planilha_local_estruturalmente_valida(caminho):
                continue
            try:
                mtime = os.path.getmtime(caminho)
            except Exception:
                mtime = 0
            candidatos.append((mtime, os.path.abspath(caminho)))

    candidatos.sort(key=lambda item: item[0], reverse=True)
    planilhas = [caminho for _mtime, caminho in candidatos]
    if not planilhas:
        return []
    print(
        f"[PORTAL CACHE] {len(planilhas)} Excel(s) estruturalmente válido(s); "
        "extração na ordem do mais recente ao mais antigo, até encontrar OVs."
    )
    return planilhas

def baixar_planilhas_anexas_portal(page, frame, requisicao, pasta_requisicao, atividade="",
                                baixar_todas=False, inspecao=None):
    """Baixa até obter OVs, ou todos os Excels na conferência prévia da NC."""
    if inspecao is not None:
        inspecao.update(completa=False, qtd_excel=None,
                        erro="Não foi possível carregar/identificar todos os anexos Excel do PORTAL.")
    # TESTE 35: garante que nenhum modal residual da navegação anterior esteja
    # cobrindo a aba Documentos antes do fallback por planilha.
    fechar_modal_residual_portal(
        page,
        motivo=f"antes de procurar planilha da requisição {requisicao}",
        timeout=5000,
    )

    frame = frame_principal_portal(page)
    portalmain = frame_portalmain(frame)

    print(
        f"[PORTAL] Requisição {requisicao}: conferindo os anexos disponíveis; "
        "procurando planilha anexada em Documentos..."
    )

    abriu = False
    tentativas = [
        lambda: portalmain.get_by_title("Documentos").click(timeout=12000),
        lambda: portalmain.get_by_text("Documentos", exact=True).click(timeout=12000),
        lambda: portalmain.get_by_text("Documentos").first.click(timeout=12000),
    ]
    for tentativa in tentativas:
        try:
            tentativa()
            abriu = True
            break
        except Exception:
            continue

    if not abriu:
        raise DocumentosPORTALIndisponiveis(
            f'PORTAL {requisicao}: não foi possível abrir Documentos para identificar os anexos.')

    try:
        portalmain.locator(FRAME_DOCUMENTOS).wait_for(
            state="attached", timeout=20000
        )
    except Exception as e:
        raise DocumentosPORTALIndisponiveis(f'PORTAL {requisicao}: a tela de Documentos não abriu: {e}') from e

    # TESTE 24:
    # O PORTAL também pode carregar parcialmente a aba Documentos: o cabeçalho da
    # requisição e a própria aba aparecem, porém a grade de anexos fica vazia.
    # Isso é o mesmo comportamento de frameset incompleto já observado no início
    # do portal de atendimento. Portanto:
    #
    #   1) aguardamos dinamicamente a grade por até 35 s;
    #   2) após ~15 s, reabrimos Documentos uma vez;
    #   3) se ainda ficar vazia, executamos F5 controlado;
    #   4) reabrimos A MESMA requisição pela consulta direta;
    #   5) voltamos para Documentos e aguardamos novamente.
    #
    # O ciclo completo é limitado a 3 tentativas para não entrar em loop infinito.

    def _clicar_documentos(portalmain_atual, timeout_ms=12000):
        """Abre/reabre a aba Documentos e confirma o iframe frameDocumentosDemo."""
        clicou = False
        for tentativa_doc in (
            lambda: portalmain_atual.get_by_title("Documentos").click(timeout=timeout_ms),
            lambda: portalmain_atual.get_by_text(
                "Documentos", exact=True
            ).click(timeout=timeout_ms),
            lambda: portalmain_atual.get_by_text("Documentos").first.click(
                timeout=timeout_ms
            ),
        ):
            try:
                tentativa_doc()
                clicou = True
                break
            except Exception:
                continue

        if not clicou:
            raise RuntimeError("Não foi possível clicar na aba Documentos.")

        portalmain_atual.locator(FRAME_DOCUMENTOS).wait_for(
            state="attached", timeout=20000
        )
        return True

    lista_vazia_confirmada = False

    def _aguardar_materializacao_documentos(portalmain_atual, ciclo):
        """
        Aguarda conteúdo real da grade de documentos.
        Retorna: (materializou, qtd_links, qtd_linhas_anexo, texto_documentos).
        """
        nonlocal lista_vazia_confirmada
        vazio_desde = None
        inicio = time.monotonic()
        limite_s = 35.0
        reclick_realizado = False
        ultimo_log_segundo = -999
        qtd_links_local = 0
        qtd_linhas_local = 0
        texto_local = ""

        while True:
            decorrido = time.monotonic() - inicio

            andoc_local = portalmain_atual.frame_locator(
                FRAME_DOCUMENTOS
            )

            try:
                qtd_links_local = andoc_local.locator("a").count()
            except Exception:
                qtd_links_local = 0

            try:
                qtd_linhas_local = andoc_local.locator(
                    "tr:has(input[type='radio'])"
                ).count()
            except Exception:
                qtd_linhas_local = 0

            try:
                texto_local = (
                    andoc_local.locator("body").inner_text(timeout=1200) or ""
                ).strip()
            except Exception:
                texto_local = ""

            tem_excel_no_texto = bool(
                re.search(
                    r"\.(?:xlsx|xlsm|xlsb|xls)\b",
                    texto_local,
                    re.I,
                )
            )

            lista_confirmada = False
            estado_documentos = None
            try:
                estado_documentos = ler_estado_documentos_portal(andoc_local)
                interpretar_lista_anexos_nc(estado_documentos, requisicao, andoc_local)
                lista_confirmada = True
            except Exception:
                pass

            if lista_confirmada and not documentos_vazios_confirmados(estado_documentos, requisicao) and (
                qtd_linhas_local > 0
                or qtd_links_local > 0
                or tem_excel_no_texto
            ):
                print(
                    f"[PORTAL] Conteúdo de Documentos carregado no ciclo "
                    f"{ciclo}/3 após {decorrido:.1f}s: "
                    f"{qtd_links_local} link(s); "
                    f"{qtd_linhas_local} linha(s) de anexo."
                )
                return (
                    True,
                    qtd_links_local,
                    qtd_linhas_local,
                    texto_local,
                )

            try:
                vazio = documentos_vazios_confirmados(estado_documentos, requisicao)
            except Exception:
                vazio = False
            if vazio:
                if vazio_desde is None:
                    vazio_desde = time.monotonic()
                elif time.monotonic() - vazio_desde >= 1.0:
                    lista_vazia_confirmada = True
                    print(f'[PORTAL] Requisição {requisicao}: Documentos carregado e vazio, confirmado pelo cabeçalho.')
                    return True, 0, 0, texto_local
            else:
                vazio_desde = None

            if decorrido >= 15.0 and not reclick_realizado:
                reclick_realizado = True
                print(
                    f"[PORTAL] Documentos continua sem anexos no ciclo "
                    f"{ciclo}/3 após 15s; reabrindo a aba uma vez..."
                )
                try:
                    _clicar_documentos(portalmain_atual, timeout_ms=8000)
                except Exception as e_reclick:
                    print(
                        "[PORTAL] Reabertura simples de Documentos não foi "
                        f"confirmada. Detalhe: {e_reclick}"
                    )

            if decorrido >= limite_s:
                print(
                    f"[PORTAL] Ciclo {ciclo}/3 atingiu {int(limite_s)}s sem "
                    "materializar nenhum anexo em Documentos."
                )
                return (
                    False,
                    qtd_links_local,
                    qtd_linhas_local,
                    texto_local,
                )

            segundo = int(decorrido)
            if segundo == 0 or segundo - ultimo_log_segundo >= 5:
                ultimo_log_segundo = segundo
                print(
                    f"[PORTAL] Aguardando materialização dos anexos "
                    f"(ciclo {ciclo}/3)... "
                    f"{decorrido:.1f}s/{limite_s:.0f}s"
                )

            page.wait_for_timeout(1000)

    materializou_documentos = False
    qtd_links = 0
    qtd_linhas_anexo = 0
    texto_documentos = ""

    for ciclo_documentos in range(1, 4):
        if ciclo_documentos > 1:
            print(
                f"[PORTAL] A aba Documentos da requisição {requisicao} permaneceu "
                f"vazia. Iniciando recuperação completa "
                f"{ciclo_documentos}/3 (F5 + reabertura da requisição)..."
            )

            recuperou = _recarregar_portal_quando_incompleto(
                page,
                tentativa=ciclo_documentos - 1,
                motivo=(
                    f"Documentos da requisição {requisicao} não materializou "
                    "nenhum anexo após 35 s."
                ),
            )

            if not recuperou:
                print(
                    "[PORTAL] O F5 controlado não confirmou a tela principal, "
                    "mas tentarei reabrir a requisição pela sessão atual."
                )

            try:
                frame = abrir_requisicao_portal(page, requisicao)
                validar_requisicao_portal_aberta(
                    frame,
                    requisicao,
                    timeout=30000,
                )
                portalmain = frame_portalmain(frame)

                print(
                    f"[PORTAL] Requisição {requisicao} reaberta após recuperação; "
                    "abrindo Documentos novamente..."
                )
                _clicar_documentos(portalmain)
            except Exception as e_recuperacao_doc:
                print(
                    f"[PORTAL] Falha ao reabrir a requisição {requisicao} após "
                    f"recuperação de Documentos: {e_recuperacao_doc}"
                )
                if ciclo_documentos >= 3:
                    break
                continue

        materializou_documentos, qtd_links, qtd_linhas_anexo, texto_documentos = (
            _aguardar_materializacao_documentos(
                portalmain,
                ciclo_documentos,
            )
        )

        if materializou_documentos:
            break

    if not materializou_documentos:
        print(
            f"[PORTAL] Requisição {requisicao}: Documentos permaneceu sem anexos "
            "após 3 ciclos completos de espera/recuperação. "
            "Não vou assumir que a requisição realmente não possui arquivo."
        )
        raise DocumentosPORTALIndisponiveis(
            f'PORTAL {requisicao}: Documentos não confirmou o carregamento após as tentativas. '
            'Não foi possível determinar se existem anexos; SAP não executado.')

    if lista_vazia_confirmada:
        if inspecao is not None:
            inspecao.update(completa=True, qtd_excel=0, erro='')
        return [], [], {}

    # Frame final, novamente obtido após toda a sincronização.
    andoc = portalmain.frame_locator(FRAME_DOCUMENTOS)
    print(
        f"[PORTAL] Documentos materializados: {qtd_links} link(s); "
        f"{qtd_linhas_anexo} linha(s) com seleção de anexo."
    )

    candidatos = []
    vistos = set()

    # ---------------------------------------------------------
    # 1) Tentativa principal: links cujo nome/href/title mostra
    #    explicitamente uma extensão Excel.
    # ---------------------------------------------------------
    links = andoc.locator("a")
    try:
        qtd = links.count()
    except Exception as e:
        print(f"[PORTAL] Não consegui enumerar os anexos da requisição: {e}")
        qtd = 0

    for i in range(qtd):
        link = links.nth(i)
        try:
            if not link.is_visible():
                continue
        except Exception:
            pass

        try:
            texto = (link.inner_text(timeout=1500) or "").strip()
        except Exception:
            texto = ""
        try:
            href = (link.get_attribute("href") or "").strip()
        except Exception:
            href = ""
        try:
            titulo = (link.get_attribute("title") or "").strip()
        except Exception:
            titulo = ""

        combinado = " ".join(x for x in (texto, titulo, href) if x)
        m = re.search(
            r"([^\\/:*?\"<>|\r\n]+?\.(?:xlsx|xlsm|xlsb|xls))(?=$|[\s?#&])",
            combinado,
            re.I,
        )
        if not m:
            m = re.search(r"([^\r\n]+?\.(?:xlsx|xlsm|xlsb|xls))$", texto, re.I)

        if not m:
            continue

        nome_visivel = m.group(1).strip()
        chave = ("link", i, texto.lower(), href.lower(), titulo.lower())
        if chave in vistos:
            continue
        vistos.add(chave)
        candidatos.append({
            "modo": "link",
            "indice": i,
            "nome": nome_visivel,
            "rotulo": texto or titulo or nome_visivel,
        })

    # ---------------------------------------------------------
    # 2) Fallback importante: o PORTAL às vezes mostra o nome do
    #    anexo truncado, por exemplo "REF 19710-400 NF SAP -",
    #    embora o download real seja .xlsx. O Codegen já mostrou
    #    esse comportamento. Nesses casos usamos as linhas que
    #    possuem radio + link e classificamos pelo filename real
    #    retornado no download.
    # ---------------------------------------------------------
    if baixar_todas:
        try:
            candidatos = catalogar_excels_credito(andoc)
        except ValidacaoCreditoError as erro:
            if inspecao is not None:
                inspecao["erro"] = str(erro)
            return [], [], {}
        if inspecao is not None:
            inspecao["qtd_excel"] = len(candidatos)
        print(f"[NC PRÉ-VALIDAÇÃO] {len(candidatos)} arquivo(s) Excel identificado(s) na grade do PORTAL.")
        if not candidatos:
            if inspecao is not None:
                inspecao.update(completa=True, erro="")
            return [], [], {}

    if not candidatos and not baixar_todas:
        print(
            "[PORTAL] Nenhum link exibiu extensão Excel explicitamente; "
            "tentando anexos pelas linhas selecionáveis do PORTAL."
        )

        linhas = andoc.locator("tr:has(input[type='radio'])")
        try:
            qtd_rows = linhas.count()
        except Exception:
            qtd_rows = 0

        for r in range(qtd_rows):
            row = linhas.nth(r)
            row_links = row.locator("a")
            try:
                nlinks = row_links.count()
            except Exception:
                nlinks = 0
            for j in range(nlinks):
                lk = row_links.nth(j)
                try:
                    if not lk.is_visible():
                        continue
                except Exception:
                    pass
                try:
                    texto = (lk.inner_text(timeout=1000) or "").strip()
                except Exception:
                    texto = ""
                try:
                    titulo = (lk.get_attribute("title") or "").strip()
                except Exception:
                    titulo = ""
                rotulo = texto or titulo
                if not rotulo:
                    continue
                chave = ("row", r, j, rotulo.lower())
                if chave in vistos:
                    continue
                vistos.add(chave)
                candidatos.append({
                    "modo": "row",
                    "row": r,
                    "link": j,
                    "nome": rotulo,
                    "rotulo": rotulo,
                })

        # Último fallback: links visíveis com texto, caso a marcação por radio
        # mude no PORTAL. Exclui rótulos típicos de comando.
        if not candidatos:
            comandos_ignorar = {
                "salvar", "cancelar", "fechar", "ok", "voltar", "novo",
                "incluir", "excluir", "documentos", "dados adicionais"
            }
            for i in range(qtd):
                lk = links.nth(i)
                try:
                    if not lk.is_visible():
                        continue
                    texto = (lk.inner_text(timeout=1000) or "").strip()
                except Exception:
                    continue
                if not texto or texto.casefold() in comandos_ignorar:
                    continue
                candidatos.append({
                    "modo": "link",
                    "indice": i,
                    "nome": texto,
                    "rotulo": texto,
                })

    if not candidatos:
        print(
            f"[PORTAL] Requisição {requisicao}: nenhum anexo clicável foi "
            "localizado em Documentos."
        )
        return [], [], {}

    print(
        f"[PORTAL] Requisição {requisicao}: {len(candidatos)} candidato(s) "
        "a anexo localizado(s) para validação por download."
    )

    salvos = []
    houve_candidato_excel = bool(candidatos)
    falhas_download = []

    def _reobter_link_e_linha(candidato_atual):
        """Reobtém locators após cada tentativa, pois o PORTAL pode recriar o DOM."""
        andoc_atual = portalmain.frame_locator(FRAME_DOCUMENTOS)
        if candidato_atual["modo"] == "row":
            linha = (
                andoc_atual.locator("tr:has(input[type='radio'])")
                .nth(candidato_atual["row"])
            )
            link_atual = linha.locator("a").nth(candidato_atual["link"])
            return andoc_atual, link_atual, linha

        link_atual = andoc_atual.locator("a").nth(candidato_atual["indice"])
        # Mesmo quando o candidato nasceu como link, tentamos localizar a linha
        # pai. No PORTAL observado, é nela que ficam o radio e o ícone de download.
        linha = link_atual.locator("xpath=ancestor::tr[1]")
        return andoc_atual, link_atual, linha

    def _tentar_download(candidato_atual, metodo, timeout_download=25000):
        """Executa um método de acionamento e retorna Download ou None."""
        andoc_atual, link_atual, linha = _reobter_link_e_linha(candidato_atual)

        try:
            print(
                f"[PORTAL] Download da planilha - método '{metodo}' para: "
                f"{candidato_atual['rotulo']}"
            )

            with page.expect_download(timeout=timeout_download) as info:
                if metodo == "clique_link":
                    link_atual.scroll_into_view_if_needed(timeout=5000)
                    link_atual.click(timeout=12000, no_wait_after=True)

                elif metodo == "duplo_clique_link":
                    link_atual.scroll_into_view_if_needed(timeout=5000)
                    link_atual.dblclick(timeout=12000)

                elif metodo == "radio_e_icone":
                    # Fluxo reproduzido do Codegen do PORTAL: seleciona a linha e
                    # dá duplo clique no ícone do documento.
                    radios = linha.locator("input[type='radio']")
                    if radios.count() > 0:
                        radio = radios.first
                        try:
                            radio.check(timeout=8000, force=True)
                        except Exception:
                            radio.click(timeout=8000, force=True)

                    icones = linha.locator("i")
                    if icones.count() == 0:
                        raise RuntimeError(
                            "Linha do anexo não possui ícone <i> para o fallback."
                        )
                    icone = icones.first
                    icone.scroll_into_view_if_needed(timeout=5000)
                    icone.dblclick(timeout=12000, force=True)

                elif metodo == "radio_e_duplo_link":
                    radios = linha.locator("input[type='radio']")
                    if radios.count() > 0:
                        radio = radios.first
                        try:
                            radio.check(timeout=8000, force=True)
                        except Exception:
                            radio.click(timeout=8000, force=True)
                    link_atual.dblclick(timeout=12000, force=True)

                else:
                    raise RuntimeError(f"Método de download desconhecido: {metodo}")

            download_atual = info.value
            print(
                f"[PORTAL] Evento de download confirmado pelo método '{metodo}': "
                f"{download_atual.suggested_filename}"
            )
            return download_atual

        except Exception as e_metodo:
            falhas_download.append(
                f"{candidato_atual.get('rotulo', '')} | {metodo}: {e_metodo}"
            )
            print(
                f"[PORTAL] Método '{metodo}' não gerou download confirmado. "
                f"Detalhe: {e_metodo}"
            )
            return None

    def _diagnostico_conteudo_curto(caminho, limite=160):
        """Retorna um trecho seguro do conteúdo quando o servidor devolve texto/erro."""
        try:
            with open(caminho, "rb") as arq:
                bruto = arq.read(limite)
            texto = bruto.decode("utf-8", errors="replace")
            texto = re.sub(r"\s+", " ", texto).strip()
            return texto[:limite]
        except Exception:
            return ""

    def _validar_arquivo_excel_real(caminho, extensao):
        """
        TESTE 23: não confia apenas no nome .xlsx retornado pelo navegador.
        Valida a assinatura/conteúdo real antes de considerar o anexo utilizável.
        """
        if not os.path.isfile(caminho):
            return False, "arquivo não foi criado"

        tamanho = os.path.getsize(caminho)
        if tamanho <= 0:
            return False, "arquivo vazio"

        extensao = str(extensao or "").lower()

        # XLSX/XLSM/XLSB são contêineres ZIP. Um download minúsculo com nome
        # .xlsx normalmente é uma mensagem de erro/redirect do servidor PORTAL.
        if extensao in {".xlsx", ".xlsm", ".xlsb"}:
            if not zipfile.is_zipfile(caminho):
                trecho = _diagnostico_conteudo_curto(caminho)
                detalhe = (
                    f"conteúdo não é um ZIP/Excel válido ({tamanho} bytes)"
                )
                if trecho:
                    detalhe += f"; retorno inicial do servidor: {trecho!r}"
                return False, detalhe
            try:
                with zipfile.ZipFile(caminho, "r") as zf:
                    nomes = set(zf.namelist())
                if extensao in {".xlsx", ".xlsm"} and "[Content_Types].xml" not in nomes:
                    return False, "ZIP recebido sem estrutura OOXML de Excel"
                if extensao == ".xlsb" and not any(
                    nome.lower().endswith("workbook.bin") for nome in nomes
                ):
                    return False, "ZIP recebido sem workbook.bin esperado para XLSB"
            except Exception as e_zip:
                return False, f"falha ao validar contêiner Excel: {e_zip}"
            return True, f"Excel ZIP válido ({tamanho} bytes)"

        # XLS legado usa Compound File Binary (OLE2).
        if extensao == ".xls":
            try:
                with open(caminho, "rb") as arq:
                    cabecalho = arq.read(8)
                if cabecalho != bytes.fromhex("D0CF11E0A1B11AE1"):
                    trecho = _diagnostico_conteudo_curto(caminho)
                    detalhe = f"assinatura OLE2 inválida para XLS ({tamanho} bytes)"
                    if trecho:
                        detalhe += f"; retorno inicial do servidor: {trecho!r}"
                    return False, detalhe
            except Exception as e_xls:
                return False, f"falha ao validar XLS legado: {e_xls}"
            return True, f"XLS OLE2 válido ({tamanho} bytes)"

        return False, f"extensão não suportada: {extensao}"

    def _destino_final_unico(nome_real):
        destino_final = os.path.join(pasta_requisicao, nome_real)
        if not os.path.exists(destino_final):
            return destino_final
        base, ext2 = os.path.splitext(nome_real)
        contador = 2
        while True:
            candidato_destino = os.path.join(
                pasta_requisicao, f"{base}_{contador}{ext2}"
            )
            if not os.path.exists(candidato_destino):
                return candidato_destino
            contador += 1

    def _salvar_e_validar_download(download_atual, candidato_atual, metodo):
        """
        Salva primeiro em arquivo temporário, valida o conteúdo real e só então
        promove para o nome definitivo. Se for resposta falsa de poucos bytes,
        remove e permite que o próximo método de clique seja tentado.
        """
        nome_real = str(download_atual.suggested_filename or "").strip()
        if not nome_real:
            nome_real = candidato_atual.get("nome", "anexo.xlsx")
        nome_real = _nome_arquivo_seguro(nome_real)
        ext = os.path.splitext(nome_real)[1].lower()

        if ext not in EXTENSOES_PLANILHA_PORTAL:
            return None, f"download retornou arquivo não Excel: {nome_real}"

        tmp_nome = (
            f".__tmp_excel_{requisicao}_{int(time.time() * 1000)}_"
            f"{re.sub(r'[^A-Za-z0-9_-]+', '_', metodo)}{ext}"
        )
        tmp_path = os.path.join(pasta_requisicao, tmp_nome)

        try:
            download_atual.save_as(tmp_path)
            try:
                falha = download_atual.failure()
            except Exception:
                falha = None
            if falha:
                return None, f"Playwright reportou falha no download: {falha}"

            valido, detalhe = _validar_arquivo_excel_real(tmp_path, ext)
            if not valido:
                return None, detalhe

            destino_final = _destino_final_unico(nome_real)
            os.replace(tmp_path, destino_final)
            return destino_final, detalhe
        finally:
            # Se não foi promovido para o nome final, remove o temporário falso.
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass

    ovs_extraidas, referencias_extraidas = [], {}

    for candidato in candidatos:
        try:
            print(
                f"[PORTAL] Tentando baixar anexo da requisição {requisicao}: "
                f"{candidato['rotulo']}"
            )

            caminho_valido = None

            # TESTE 23: um evento de download não significa que recebemos a
            # planilha. Cada método só é considerado sucesso depois de validar
            # a assinatura/conteúdo real do Excel. Se vier, por exemplo, um
            # falso .xlsx de 47 bytes, removemos e tentamos o próximo método.
            for metodo, timeout_dl in (
                ("clique_link", 25000),
                ("duplo_clique_link", 25000),
                ("radio_e_icone", 30000),
                ("radio_e_duplo_link", 25000),
            ):
                download = _tentar_download(
                    candidato, metodo, timeout_download=timeout_dl
                )
                if download is None:
                    continue

                try:
                    caminho_valido, detalhe_validacao = _salvar_e_validar_download(
                        download, candidato, metodo
                    )
                except Exception as e_validacao:
                    caminho_valido = None
                    detalhe_validacao = str(e_validacao)

                if caminho_valido:
                    tamanho = os.path.getsize(caminho_valido)
                    salvos.append(caminho_valido)
                    print(
                        f"[PORTAL] Planilha real validada pelo método '{metodo}': "
                        f"{caminho_valido} ({tamanho} bytes). "
                        f"Validação: {detalhe_validacao}"
                    )
                    break

                falhas_download.append(
                    f"{candidato.get('rotulo', '')} | {metodo}: "
                    f"download iniciou, mas conteúdo foi rejeitado: {detalhe_validacao}"
                )
                print(
                    f"[PORTAL] Método '{metodo}' iniciou um download, porém o "
                    f"arquivo recebido NÃO é uma planilha válida. "
                    f"Detalhe: {detalhe_validacao}"
                )
                print(
                    "[PORTAL] O arquivo inválido foi removido; tentando o próximo "
                    "método de download..."
                )
                page.wait_for_timeout(1200)

            if caminho_valido is None:
                # Recuperação final: reabre Documentos e reproduz o método do
                # Codegen (radio + ícone), novamente com validação real.
                print(
                    "[PORTAL] Nenhum método entregou uma planilha válida; "
                    "reabrindo Documentos uma vez para tentativa final."
                )
                try:
                    clicou = False
                    for tentativa_reabrir in (
                        lambda: portalmain.get_by_title("Documentos").click(timeout=8000),
                        lambda: portalmain.get_by_text("Documentos", exact=True).click(timeout=8000),
                        lambda: portalmain.get_by_text("Documentos").first.click(timeout=8000),
                    ):
                        try:
                            tentativa_reabrir()
                            clicou = True
                            break
                        except Exception:
                            continue

                    if clicou:
                        portalmain.locator(FRAME_DOCUMENTOS).wait_for(
                            state="attached", timeout=15000
                        )
                        page.wait_for_timeout(2500)
                        download = _tentar_download(
                            candidato, "radio_e_icone", timeout_download=35000
                        )
                        if download is not None:
                            try:
                                caminho_valido, detalhe_validacao = _salvar_e_validar_download(
                                    download, candidato, "radio_e_icone_final"
                                )
                            except Exception as e_validacao:
                                caminho_valido = None
                                detalhe_validacao = str(e_validacao)

                            if caminho_valido:
                                tamanho = os.path.getsize(caminho_valido)
                                salvos.append(caminho_valido)
                                print(
                                    "[PORTAL] Planilha real validada na tentativa "
                                    f"final: {caminho_valido} ({tamanho} bytes). "
                                    f"Validação: {detalhe_validacao}"
                                )
                            else:
                                falhas_download.append(
                                    f"{candidato.get('rotulo', '')} | "
                                    f"radio_e_icone_final: {detalhe_validacao}"
                                )
                except Exception as e_reabrir:
                    falhas_download.append(
                        f"{candidato.get('rotulo', '')} | reabertura final: {e_reabrir}"
                    )
                    print(
                        "[PORTAL] A tentativa final após reabrir Documentos "
                        f"falhou: {e_reabrir}"
                    )

            if caminho_valido is None:
                print(
                    f"[PORTAL] ERRO_DOWNLOAD_PLANILHA: o anexo existe no PORTAL, "
                    f"mas nenhum método retornou um Excel estruturalmente válido: "
                    f"{candidato['rotulo']}"
                )

            # TESTE 35:
            # O fluxo radio + ícone pode deixar #modal-selecao2 aparecer com
            # atraso. O helper aguarda e fecha também a janela branca/overlay
            # que foi observada na rodada 34.
            fechar_modal_residual_portal(
                page,
                motivo=f"após download de planilha da requisição {requisicao}",
                timeout=7000,
            )

            if caminho_valido and not baixar_todas:
                ovs_extraidas, referencias_extraidas = extrair_ovs_de_planilhas(
                    [caminho_valido], atividade=atividade,
                )
                if ovs_extraidas:
                    print(
                        f"[PORTAL] Requisição {requisicao}: Excel válido com "
                        f"{len(ovs_extraidas)} OV(s); não serão tentados outros anexos Excel."
                    )
                    break
                print(
                    f"[PORTAL] Requisição {requisicao}: Excel válido sem OV extraída; "
                    "continuando para o próximo anexo candidato."
                )

        except Exception as e:
            falhas_download.append(
                f"{candidato.get('rotulo', '')} | processamento: {e}"
            )
            print(
                f"[PORTAL] Falha ao processar candidato de anexo "
                f"'{candidato.get('rotulo', '')}': {e}"
            )

    # Última barreira: mesmo que o modal tenha aparecido alguns segundos
    # depois do evento de download, não deixamos a requisição seguir para o SAP
    # com a interface PORTAL bloqueada.
    fechar_modal_residual_portal(
        page,
        motivo=f"ao finalizar leitura de planilhas da requisição {requisicao}",
        timeout=5000,
    )

    salvos = list(dict.fromkeys(os.path.abspath(x) for x in salvos))

    if not salvos:
        if houve_candidato_excel:
            print(
                f"[PORTAL] Requisição {requisicao}: ERRO_DOWNLOAD_PLANILHA - "
                "o anexo foi localizado, mas nenhum download Excel válido "
                "foi concluído."
            )
            if falhas_download:
                print(
                    "[PORTAL] Última falha de download: "
                    f"{falhas_download[-1]}"
                )
        else:
            print(
                f"[PORTAL] Requisição {requisicao}: anexos foram verificados, "
                "mas nenhum download Excel válido foi confirmado."
            )

    if inspecao is not None and baixar_todas:
        inspecao["completa"] = len(salvos) == inspecao.get("qtd_excel")
        inspecao["erro"] = "" if inspecao["completa"] else (
            f"Download incompleto: {len(salvos)} de {inspecao.get('qtd_excel')} Excel(s). "
            + (falhas_download[-1] if falhas_download else "Arquivo ausente.")
        )
    return salvos, ovs_extraidas, referencias_extraidas

# =========================
# PASTAS / PDF
# =========================

def criar_pasta_requisicao(requisicao):
    """
    Cria a pasta da requisição dentro da pasta oficial de saídas.
    Exemplo:
    requisições/1234567
    """
    pasta_req = os.path.join(PASTA_ANEXOS_SAP, str(requisicao).strip())
    os.makedirs(pasta_req, exist_ok=True)
    return pasta_req


def montar_caminho_pdf(requisicao, ordem):
    """
    Monta o caminho final do PDF usando a pasta da requisição e o número da OV.
    Exemplo:
    requisições/1234567/7654321.pdf
    """
    pasta_req = criar_pasta_requisicao(requisicao)
    nome_pdf = f"{str(ordem).strip()}.pdf"
    return os.path.join(pasta_req, nome_pdf)


def validar_associacao_resultado_sap(requisicao, ordem, resultado_sap):
    """Bloqueia qualquer resultado que retorne OV ou caminho de outra iteração."""
    requisicao = str(requisicao).strip()
    ordem = str(ordem).strip()
    resultado_sap = resultado_sap or {}
    ordem_resultado = str(resultado_sap.get("Ordem", "") or "").strip()

    if ordem_resultado != ordem:
        raise RuntimeError(
            f"Resultado SAP cruzado. Requisição {requisicao}: esperada OV "
            f"{ordem}, mas o resultado informou {ordem_resultado or '<vazio>'}."
        )

    caminho = str(resultado_sap.get("ArquivoPDF", "") or "").strip()
    if caminho:
        caminho_esperado = os.path.abspath(montar_caminho_pdf(requisicao, ordem))
        if os.path.abspath(caminho).lower() != caminho_esperado.lower():
            raise RuntimeError(
                f"Caminho PDF cruzado para a OV {ordem}. Esperado: "
                f"{caminho_esperado}; recebido: {caminho}."
            )

    return True


def salvar_dados_requisicao(
    requisicao,
    descricao,
    ordens,
    consultas_sap=None,
    resultado_portal=None,
    erro="",
):
    """Mantém um resumo legível da requisição e de suas OVs na própria pasta."""
    pasta_req = criar_pasta_requisicao(requisicao)
    caminho = os.path.join(pasta_req, f"dados_requisicao_{requisicao}.txt")
    consultas_sap = consultas_sap or []
    resultado_portal = resultado_portal or {}

    linhas = [
        f"Requisição PORTAL: {requisicao}",
        f"Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        "",
        f"OVs localizadas: {len(ordens or [])}",
    ]

    for ordem in ordens or []:
        linhas.append(f"- {ordem}")

    if consultas_sap:
        linhas.extend(["", "RESULTADO DAS OVs"])
        for consulta in consultas_sap:
            linhas.extend([
                f"OV: {consulta.get('Ordem', '')}",
                f"Status SAP: {consulta.get('StatusSAP', '')}",
                f"NF SAP: {consulta.get('NumeroNF', '')}",
                f"Arquivo PDF: {consulta.get('ArquivoPDF', '')}",
                f"Mensagem SAP: {consulta.get('Mensagem', '')}",
                "",
            ])

    if resultado_portal:
        linhas.extend([
            "RESULTADO PORTAL",
            f"Status PORTAL: {resultado_portal.get('StatusPORTAL', '')}",
            f"Qtd. PDFs anexados: {resultado_portal.get('QtdPDFsAnexados', '')}",
            f"Mensagem PORTAL: {resultado_portal.get('MensagemPORTAL', '')}",
            "",
        ])

    if resultado_portal:
        campos = campos_situacao_conciliacao(dict(resultado_portal, Requisicao=str(requisicao)))
        linhas.extend([f"{nome}: {valor or '—'}" for nome, valor in campos.items()])
        linhas.append('')
    if resultado_portal.get('StatusConciliacao'):
        linhas.extend([
            'CONCILIAÇÃO DA NC COM DOIS EXCELS',
            f"Status: {resultado_portal.get('StatusConciliacao', '')}",
            f"Nova requisição: {resultado_portal.get('RequisicaoConciliacao', '')}",
            f"Planilha: {resultado_portal.get('PlanilhaConciliacao', '')}",
            f"Mensagem: {resultado_portal.get('MensagemConciliacao', '')}", '',
        ])

    if erro:
        linhas.extend(["ERRO", str(erro), ""])

    linhas.extend(["DESCRIÇÃO CAPTURADA NO PORTAL", str(descricao or "")])

    with open(caminho, "w", encoding="utf-8-sig") as arquivo:
        arquivo.write("\n".join(linhas))

    return caminho


# =========================
# REGRAS DE STATUS SAP
# =========================


def normalizar_texto_busca(texto):
    """
    Normaliza texto para comparação tolerante:
    - minúsculo;
    - sem acentos;
    - espaços duplicados reduzidos.
    """
    texto = str(texto or "")
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    texto = texto.lower()
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def texto_indica_erro_determinacao_conta(texto):
    """Reconhece somente mensagens reais de erro de determinação de conta."""
    txt = normalizar_texto_busca(texto)
    padroes = (
        r"\berro\s+na\s+determ(?:\.|inacao)?\s*(?:de\s+)?conta\b",
        r"\berro\s+(?:de\s+)?determinacao\s+(?:de\s+)?conta\b",
        r"\bconta\s+nao\s+pode\s+ser\s+determinada\b",
    )
    return any(re.search(p, txt, re.I) for p in padroes)



def consulta_indica_erro_determinacao_conta(consulta):
    """
    Retorna True somente quando a OV foi classificada com o status técnico
    ERRO_DETERMINACAO_CONTA.

    Importante:
    Não usamos a mensagem livre para decidir isso, porque mensagens antigas ou
    outras linhas do Fluxo de documentos poderiam conter o mesmo texto.
    """
    status = str((consulta or {}).get("StatusSAP", "") or "").upper()
    return status == STATUS_ERRO_DETERMINACAO_CONTA


def existe_erro_determinacao_conta_na_requisicao(consultas_sap):
    """Indica se qualquer OV da requisição teve Erro na determ. conta."""
    return any(consulta_indica_erro_determinacao_conta(c) for c in (consultas_sap or []))


def montar_mensagem_erro_determinacao_conta_portal(consultas_sap):
    """Monta mensagem de auditoria para o Excel quando o PORTAL ficar em aberto."""
    linhas = []
    for consulta in consultas_sap or []:
        if consulta_indica_erro_determinacao_conta(consulta):
            ordem = consulta.get("Ordem", "")
            msg = consulta.get("Mensagem", "") or "Erro na determ. conta"
            linhas.append(f"OV {ordem}: {msg}")
    if not linhas:
        return "Erro de determinação de conta identificado. Requisição PORTAL mantida em aberto."
    return "Erro de determinação de conta identificado. Requisição PORTAL mantida em aberto. " + " | ".join(linhas)


# =========================
# SAP FIORI / SAP GUI FOR HTML
# =========================



def pdf_valido(caminho_pdf):
    """Valida existência, tamanho e assinatura básica de um arquivo PDF."""
    try:
        if not caminho_pdf or not os.path.exists(caminho_pdf):
            return False
        if os.path.getsize(caminho_pdf) < 5:
            return False
        with open(caminho_pdf, "rb") as arquivo:
            return arquivo.read(5) == b"%PDF-"
    except Exception:
        return False


def _locators_marcadores_aplicativo_fiori(frame_fiori, tipo_aplicativo):
    """Monta marcadores que identificam inequivocamente cada tela SAP."""
    tipo = str(tipo_aplicativo or "").strip().lower()

    if tipo == "vf01":
        grid = frame_fiori.get_by_role(
            "grid",
            name=re.compile(r"Docs\.\s*a\s+serem\s+processados", re.I),
        )
        return [
            grid.locator('input[name="InputField"]'),
        ]

    if tipo == "va03":
        return [
            frame_fiori.get_by_role("textbox", name="Ordem", exact=True),
            frame_fiori.get_by_role(
                "textbox",
                name=re.compile(r"^\s*Ordem\s*$", re.I),
            ),
            frame_fiori.locator('input[aria-label="Ordem"]'),
            frame_fiori.locator('input[title="Ordem"]'),
        ]

    raise ValueError(f"Aplicativo Fiori sem marcadores configurados: {tipo_aplicativo}")


def localizar_contexto_aplicativo_fiori(
    page_fiori,
    tipo_aplicativo,
    timeout=TEMPO_FIORI_ELEMENTO,
):
    """
    Localiza a página e o frame reais do aplicativo pelo conteúdo carregado.

    Examina somente a aba Fiori recebida. Isso impede que uma tela antiga,
    ainda aberta em outra aba, seja confundida com a OV atual.
    """
    limite = time.time() + (timeout / 1000)
    ultimo_mapa_frames = []

    while time.time() < limite:
        paginas = [page_fiori]

        mapa_frames = []
        for pagina in paginas:
            try:
                if pagina.is_closed():
                    continue
                frames = pagina.frames
            except Exception:
                continue

            # Frames recém-criados costumam estar no final da lista.
            for frame in reversed(frames):
                try:
                    if frame.is_detached():
                        continue
                except Exception:
                    pass

                try:
                    mapa_frames.append(
                        f"nome={frame.name or '<principal>'}; url={frame.url[:160]}"
                    )
                except Exception:
                    pass

                for marcador in _locators_marcadores_aplicativo_fiori(
                    frame,
                    tipo_aplicativo,
                ):
                    try:
                        quantidade = marcador.count()
                        for indice in range(min(quantidade, 10)):
                            if (marcador.nth(indice).is_visible() and
                                    _frame_fiori_realmente_visivel(frame)):
                                print(
                                    f"[FIORI] Tela {str(tipo_aplicativo).upper()} "
                                    f"identificada no frame "
                                    f"'{frame.name or '<principal>'}'."
                                )
                                pagina.bring_to_front()
                                return pagina, frame
                    except Exception:
                        continue

        ultimo_mapa_frames = mapa_frames
        page_fiori.wait_for_timeout(300)

    detalhe = " | ".join(ultimo_mapa_frames[-8:]) or "nenhum frame disponível"
    raise RuntimeError(
        f"A tela {str(tipo_aplicativo).upper()} não foi reconhecida em nenhum "
        f"frame ou janela. Frames observados: {detalhe}"
    )


def preencher_campo_fiori(campo, valor, descricao):
    """Preenche um campo e confirma que o SAP recebeu exatamente o valor."""
    valor = str(valor).strip()
    campo = campo.first
    campo.wait_for(state="visible", timeout=TEMPO_FIORI_ELEMENTO)
    campo.click(timeout=10000)
    campo.fill(valor)

    try:
        valor_lido = str(campo.input_value(timeout=3000) or "").strip()
    except Exception:
        valor_lido = ""

    if valor_lido != valor:
        campo.press("Control+a")
        campo.type(valor, delay=60)
        try:
            valor_lido = str(campo.input_value(timeout=3000) or "").strip()
        except Exception:
            valor_lido = ""

    if valor_lido != valor:
        raise RuntimeError(
            f"O SAP não confirmou o preenchimento de {descricao}. "
            f"Esperado: {valor}; campo retornou: {valor_lido or '<vazio>'}."
        )

    print(f"[FIORI] {descricao} preenchida e validada: {valor}")
    return campo


def limpar_e_preencher_ov_vf01(grid, ordem):
    """Limpa o campo principal da VF01 e insere somente a OV atual."""
    ordem = str(ordem).strip()
    campo = grid.locator('input[name="InputField"]').first
    campo.wait_for(state="visible", timeout=TEMPO_FIORI_ELEMENTO)

    valor_anterior = str(campo.input_value(timeout=3000) or "").strip()
    campo.click(timeout=10000)
    campo.press("Control+a")
    campo.press("Backspace")
    campo.fill("")

    if str(campo.input_value(timeout=3000) or "").strip():
        raise RuntimeError(
            f"O campo da VF01 não ficou vazio antes de inserir a OV {ordem}."
        )

    if valor_anterior:
        print(f"[FIORI VF01] OV anterior removida do campo: {valor_anterior}")
    else:
        print("[FIORI VF01] Campo confirmado vazio antes da nova OV.")

    campo = preencher_campo_fiori(
        campo,
        ordem,
        descricao="OV na VF01",
    )

    # No SAP WebGUI, alterar somente o valor visível do input pode não atualizar
    # o estado interno da transação enquanto o foco continuar no campo. O Tab
    # força o evento de saída/alteração antes do clique em Gravar ou do Ctrl+S.
    campo.press("Tab")
    time.sleep(0.3)

    valor_confirmado = str(campo.input_value(timeout=3000) or "").strip()
    if valor_confirmado != ordem:
        raise RuntimeError(
            f"A OV {ordem} mudou após a confirmação do campo da VF01. "
            f"Valor encontrado: {valor_confirmado or '<vazio>'}."
        )

    print(f"[FIORI VF01] OV {ordem} confirmada no SAP após sair do campo.")
    return campo


def _frame_fiori_realmente_visivel(frame):
    """Desconsidera iframes antigos ocultos, inclusive por ancestrais."""
    try:
        atual = frame
        while atual.parent_frame is not None:
            elemento = atual.frame_element()
            if not elemento.is_visible():
                return False
            if not elemento.evaluate("""el => {
                const r = el.getBoundingClientRect();
                return !el.closest('[aria-hidden="true"]') &&
                    r.bottom > 0 && r.right > 0 &&
                    r.top < innerHeight && r.left < innerWidth;
            }"""):
                return False
            atual = atual.parent_frame
        return not frame.is_detached()
    except Exception:
        return False


def _observar_status_vf01(frame):
    """Marca atualizações reais de status; texto antigo não confirma outro F8."""
    try:
        frame.evaluate("""() => {
            window.__portalVF01Observer?.disconnect();
            window.__portalVF01Status = 0;
            const sel = '[role="status"],[role="alert"],[aria-live="polite"],[aria-live="assertive"]';
            window.__portalVF01Observer = new MutationObserver(ms => {
                if (ms.some(m => {
                    const e = m.target.nodeType === 1 ? m.target : m.target.parentElement;
                    return e?.closest(sel) || [...m.addedNodes].some(n =>
                        n.nodeType === 1 && (n.matches(sel) || n.querySelector(sel)));
                })) window.__portalVF01Status++;
            });
            window.__portalVF01Observer.observe(document.body,
                {subtree:true, childList:true, characterData:true});
        }""")
    except Exception:
        pass


def _status_ov_grade_vf01(frame_fiori, ordem=""):
    """Lê a grade atual numa única avaliação, sem esperar um input antigo.

    Após F8, o SAP pode recriar a grade, trocar o nome do input ou mostrar o
    Documento como texto. A OV esperada e o status precisam estar visíveis na
    mesma linha da grade VF01. O rodapé não é usado como status dessa linha.
    """
    vazio = {"grade": False, "ov": "", "status": "", "evidencias": []}
    try:
        grades = frame_fiori.evaluate(r"""() => {
            const visivel = e => {
                const r = e.getBoundingClientRect(), cs = getComputedStyle(e);
                return r.width > 0 && r.height > 0 && cs.visibility !== 'hidden' &&
                    cs.display !== 'none' && !e.closest('[aria-hidden="true"]') &&
                    r.bottom > 0 && r.top < innerHeight &&
                    r.right > 0 && r.left < innerWidth;
            };
            const saida = [];
            for (const grid of document.querySelectorAll('[role="grid"], [role="table"], table')) {
                if (!visivel(grid)) continue;
                // Não repete tabelas internas de uma grade semântica.
                const pai = grid.parentElement?.closest('[role="grid"], [role="table"]');
                if (pai) continue;
                if (grid.tagName === 'TABLE' && grid.querySelector('[role="grid"], [role="table"]')) continue;
                const nomes = [grid.getAttribute('aria-label') || '',
                    grid.getAttribute('title') || ''];
                for (const id of (grid.getAttribute('aria-labelledby') || '').split(/\s+/)) {
                    const rotulo = document.getElementById(id);
                    if (rotulo) nomes.push(rotulo.innerText || rotulo.textContent || '');
                }
                const dados = {id:grid.id || '', nomes, texto:grid.innerText || '', caixas:[]};
                const coletar = (e, r, texto, entrada) => {
                    if (!e || !texto || !visivel(e) || r.width <= 0 || r.height <= 0 ||
                        r.bottom <= 0 || r.top >= innerHeight || r.right <= 0 || r.left >= innerWidth) return;
                    dados.caixas.push({id:e.id || '', texto, entrada,
                        x:r.left, y:(r.top+r.bottom)/2, direita:r.right, altura:r.height});
                };
                const walker = document.createTreeWalker(grid, NodeFilter.SHOW_TEXT);
                let no;
                while ((no = walker.nextNode())) {
                    const range = document.createRange();
                    range.selectNodeContents(no);
                    coletar(no.parentElement, range.getBoundingClientRect(), no.nodeValue, false);
                }
                for (const e of grid.querySelectorAll('input:not([type="password"]), textarea')) {
                    coletar(e, e.getBoundingClientRect(), e.value, true);
                }
                saida.push(dados);
            }
            return saida;
        }""")
        esperado = _normalizar_numero_documento_sap(ordem)
        candidatas = []
        documentos_visiveis = []
        grades_identificadas = []
        for grade in grades:
            nome = normalizar_texto_busca(" ".join(grade.get("nomes", [])))
            texto_grade = normalizar_texto_busca(grade.get("texto", ""))
            nome_vf01 = bool(re.search(r"\bdocs\.?\s*a\s+serem\s+processados\b", nome))
            cabecalhos_vf01 = (bool(re.search(r"\bdocumento\b", texto_grade)) and
                               "status de processamento" in texto_grade)
            if not (nome_vf01 or cabecalhos_vf01):
                continue
            grades_identificadas.append(grade.get("id", ""))
            caixas = grade.get("caixas", [])
            ancoras = []
            for caixa in caixas:
                valor = str(caixa.get("texto", "") or "").strip()
                if not re.fullmatch(r"\d{7,12}", valor):
                    continue
                if valor not in documentos_visiveis:
                    documentos_visiveis.append(valor)
                if esperado and _normalizar_numero_documento_sap(valor) == esperado:
                    ancoras.append(caixa)
            if not ancoras:
                continue
            # Duplicação input/texto da MESMA célula é tolerada. A mesma OV
            # aparecendo em duas linhas diferentes exige esclarecimento.
            if max(a["y"] for a in ancoras) - min(a["y"] for a in ancoras) > 8:
                return dict(vazio, grade=True, origem="OV repetida em linhas diferentes",
                            documentos_visiveis=documentos_visiveis)
            base = min(ancoras, key=lambda a: a["x"])
            textos = [c for c in caixas if
                      abs(c["y"] - base["y"]) <= 5 and
                      c["altura"] <= max(base["altura"] + 12, 44) and
                      c["x"] >= base["direita"] - 3]
            textos.sort(key=lambda c: c["x"])
            if textos:
                textos.append({"id":"mesma-faixa-da-ov", "texto":" ".join(
                    str(c.get("texto", "") or "") for c in textos)})
            estados = set()
            evidencias = []
            for item in textos:
                texto = str(item.get("texto", "") or "")
                encontrados = re.findall(
                    r"\b(faturamento completo|faturamento incompleto|incompleto|faturamento parcial|nao faturado|"
                    r"nao processado|em processamento|pendente)\b",
                    normalizar_texto_busca(texto),
                )
                if encontrados:
                    estados.update('faturamento incompleto' if e == 'incompleto' else e for e in encontrados)
                    evidencias.append({"id":item.get("id", ""),
                        "texto":re.sub(r"\s+", " ", texto).strip()[:300]})
            candidatas.append({"grade":True, "ov":str(ordem),
                "status":next(iter(estados)) if len(estados) == 1 else "",
                "origem":"campo atual" if base.get("entrada") else "texto da linha atual",
                "id_grade":grade.get("id", ""), "evidencias":evidencias[:10]})
        if len(candidatas) == 1:
            return candidatas[0]
        if len(candidatas) > 1:
            return dict(vazio, grade=True, origem="OV encontrada em mais de uma grade")
        return dict(vazio, grade=bool(grades_identificadas),
                    origem="OV esperada não localizada na grade atual",
                    grades=grades_identificadas, documentos_visiveis=documentos_visiveis[:20])
    except Exception as erro:
        return dict(vazio, erro=str(erro))



def _estado_atual_vf01(page_fiori, ordem=""):
    """Reobtém o frame após F8; identifica também a tela do print da rodada 36."""
    estados = []
    for frame in reversed(page_fiori.frames):
        if not _frame_fiori_realmente_visivel(frame):
            continue
        try:
            inicial = _primeiro_item_visivel(
                _locators_marcadores_aplicativo_fiori(frame, "vf01")
            )
            tipo = _primeiro_item_visivel([frame.get_by_role(
                "textbox", name=re.compile(r"Tp\.doc\.faturamento", re.I)
            )])
            gravar = _primeiro_item_visivel([frame.get_by_role(
                "button", name=re.compile(r"^\s*Gravar\s*(?:\(.*\))?\s*$", re.I)
            )])
            texto = frame.locator("body").inner_text(timeout=1200)
            normal = normalizar_texto_busca(texto)
            faturamento = tipo is not None or (gravar is not None and
                "pagador" in normal and "data do faturamento" in normal)
            linha_ov = _status_ov_grade_vf01(frame, ordem)
            grade_inicial = linha_ov.get("grade", False)
            if inicial is None and not grade_inicial and not faturamento:
                # A tela pós-gravação pode conter só a mensagem e a barra SAP,
                # mas uma notificação antiga do Shell NÃO é resposta WebGUI.
                ancestral = frame
                contexto_webgui = False
                while ancestral is not None:
                    if (str(ancestral.name or "").lower().startswith("application-")
                            or "/sap/bc/gui/" in str(ancestral.url or "").lower()):
                        contexto_webgui = True
                        break
                    ancestral = ancestral.parent_frame
                if not contexto_webgui:
                    continue
                mensagem = capturar_mensagem_fiori(frame)
                if not mensagem:
                    continue
            else:
                mensagem = capturar_mensagem_fiori(frame)
            # O número temporário pode estar em input.value, não em inner_text.
            temporarios = frame.locator("input:not([type='password'])").evaluate_all(
                "els => els.filter(e => e.getClientRects().length).map(e => e.value || '')"
            )
            rascunho = bool(re.search(r"(?<![\w])\$0+\d+\b", texto + " " + " ".join(temporarios)))
            observacao = frame.evaluate("""() => ({
                alteracoes: window.__portalVF01Status || 0,
                ocupado: document.readyState !== 'complete' || [...document.querySelectorAll(
                    '[aria-busy="true"], [role="progressbar"], [aria-modal="true"], ' +
                    '[class*="BusyIndicator"], [class*="busyIndicator"]')].some(e => {
                        const r = e.getBoundingClientRect(), cs = getComputedStyle(e);
                        return r.width > 0 && r.height > 0 && cs.visibility !== 'hidden' &&
                            cs.display !== 'none' && !e.closest('[aria-hidden="true"]');
                    })
            })""")
            estados.append({"frame": frame, "inicial": inicial,
                "tipo": tipo, "gravar": gravar, "faturamento": faturamento,
                "rascunho": rascunho, "mensagem": mensagem,
                "alteracoes": observacao.get("alteracoes", 0),
                "ocupado": observacao.get("ocupado", True),
                "linha_ov": linha_ov, "grade_inicial": grade_inicial})
        except Exception:
            continue
    # A tela atual tem precedência sobre mensagens residuais de outros contextos.
    return next((e for e in estados if e["faturamento"]),
                next((e for e in estados if e["grade_inicial"]),
                     estados[0] if estados else {}))


def _classificar_gravacao_vf01(mensagem):
    """Distingue gravação normal, gravação com erro contábil e resposta incerta.

    VF051 confirma um documento gravado, mas não sua conclusão contábil.
    Só a ressalva conhecida, entre parênteses e após o número, é aceita.
    Negativas, outros erros e mensagens sem número continuam bloqueados.
    A atualidade da resposta é validada pelo chamador, depois de F8/Gravar.
    """
    texto = normalizar_texto_busca(mensagem)
    if not _extrair_documento_gravado_vf01(texto):
        return ""
    ressalva = r"\(\s*erro\s+na\s+determ(?:\.|inacao)\s*(?:de\s+)?conta\s*\)"
    erro_conta = bool(re.search(
        r"\bdocumento\s+\d{8,10}\s+(?:foi\s+)?(?:gravado|salvo)\s*" + ressalva,
        texto,
    ))
    sem_ressalva = re.sub(ressalva, "", texto) if erro_conta else texto
    if re.search(r"\b(?:nao|erro|falha|impossivel)\b", sem_ressalva):
        return ""
    return "FATURAMENTO_GRAVADO_ERRO_CONTA" if erro_conta else "FATURAMENTO_GRAVADO"


def _retorno_vf01(estado, ordem, cenario, metodo, detalhe=""):
    """Registra um desfecho confirmado e libera a próxima etapa."""
    gravado = cenario in {"FATURAMENTO_GRAVADO", "FATURAMENTO_GRAVADO_ERRO_CONTA"}
    mensagem = estado.get("mensagem", "")
    documento = _extrair_documento_gravado_vf01(mensagem) if gravado else ""
    FIORI_PENDENCIA_GRAVACAO.clear()
    print(f"[FIORI VF01] OV {ordem}: {cenario}. {detalhe or mensagem}")
    if estado.get("linha_ov", {}).get("status"):
        print(f"[FIORI VF01] Evidência: {estado['linha_ov']}")
    return {"metodo": metodo, "cenario": cenario, "ctrl_s_ok": metodo == "F8_E_CTRL_S",
            "mensagem": mensagem, "detalhe_salvar": detalhe,
            "gravacao_confirmada": gravado, "documento_faturamento": documento,
            "frame": estado["frame"]}


def mensagem_documento_inexistente_vf01(mensagem, ordem):
    """Somente a negativa explícita para o número atual; nunca só 'não existe'."""
    texto = normalizar_texto_busca(mensagem)
    padrao = r'\b(?:o\s+)?documento\s+(\d{7,12})\s+nao\s+existe\b'
    documentos = re.findall(padrao, texto)
    esperado = _normalizar_numero_documento_sap(ordem)
    # Aceita a repetição da mesma mensagem e os rótulos conhecidos do SAP,
    # mas não uma negativa misturada a outra falha/sucesso de significado incerto.
    restante = re.sub(padrao, '', texto)
    restante = re.sub(r'\b(?:visualizar detalhes|erro)\b', '', restante)
    return bool(esperado and documentos and
                {_normalizar_numero_documento_sap(n) for n in documentos} == {esperado}
                and not re.search(r'\w', restante))


def executar_e_gravar_vf01(frame_fiori, campo_ov):
    """F8; documento novo confirmado ou estado da OV; uma gravação no máximo.

    Completo segue para VA03. Incompleto fica pendente. Mensagem antiga,
    botão Gravar isolado e ausência de input não determinam o resultado.
    """
    if FIORI_PENDENCIA_GRAVACAO:
        raise RuntimeError("Existe resultado VF01 incerto; nenhuma nova escrita foi enviada.")
    pagina = frame_fiori.page
    ordem = str(campo_ov.input_value(timeout=3000) or "").strip()
    if not re.fullmatch(r"\d{7,10}", ordem):
        raise RuntimeError(f"VF01 sem OV válida no campo Documento: {ordem!r}.")
    linha_antes = _status_ov_grade_vf01(frame_fiori, ordem)
    mensagem_antes = capturar_mensagem_fiori(frame_fiori)
    _observar_status_vf01(frame_fiori)
    FIORI_PENDENCIA_GRAVACAO.update({"fase": "F8_SEM_RESULTADO_CONFIRMADO", "ordem": ordem})
    print(f"[FIORI VF01] Executando a OV {ordem} com F8...")
    campo_ov.press("F8")
    limite = time.monotonic() + TEMPO_VF01_TRANSICAO / 1000
    estado_estavel = ""
    estavel_desde = None
    rejeicao_estavel = None
    documento_novo_observado = False
    estado = {}

    while True:
        estado = _estado_atual_vf01(pagina, ordem)
        agora = time.monotonic()
        linha = estado.get("linha_ov", {})
        mensagem = estado.get("mensagem", "")
        bloqueado = (estado.get("ocupado", True) or
                     _sap_busy_indicator_visivel(pagina) or _sap_blocklayer_popup_visivel(pagina))
        mesma_ov = (_normalizar_numero_documento_sap(linha.get("ov")) ==
                    _normalizar_numero_documento_sap(ordem))
        documento = _extrair_documento_gravado_vf01(mensagem)
        cenario_gravacao = _classificar_gravacao_vf01(mensagem)
        if estado.get('faturamento') or estado.get('rascunho'):
            documento_novo_observado = True
        # Um número já presente antes de F8 não é uma nova confirmação.
        if (not bloqueado and not estado.get("rascunho") and documento and
                mensagem != mensagem_antes and
                documento != _extrair_documento_gravado_vf01(mensagem_antes) and
                cenario_gravacao):
            return _retorno_vf01(estado, ordem, cenario_gravacao,
                                "F8_E_RODAPE_DOCUMENTO_GRAVADO")

        # A tela de documento novo prevalece sobre qualquer rodapé residual.
        # Gravar também existe na grade inicial, por isso o botão sozinho não basta.
        if (not bloqueado and estado.get("faturamento") and
                not estado.get("grade_inicial")):
            break

        # TESTE 55: o SAP recusou o número da descrição na própria grade.
        # Só confirma essa negativa ANTES de qualquer Gravar: resposta nova,
        # número exato, grade atual, sem documento/rascunho e estado estável.
        inexistente = (not bloqueado and mesma_ov and estado.get('grade_inicial')
                       and not documento_novo_observado and not linha.get('status')
                       and not linha.get('evidencias')
                       and not documento and mensagem != mensagem_antes
                       and mensagem_documento_inexistente_vf01(mensagem, ordem))
        if inexistente:
            if rejeicao_estavel is None:
                rejeicao_estavel = agora
            if agora - rejeicao_estavel >= TEMPO_VF01_ESTABILIDADE / 1000:
                return _retorno_vf01(estado, ordem, 'DOCUMENTO_INEXISTENTE',
                                    'F8_E_DOCUMENTO_INEXISTENTE',
                                    f'SAP rejeitou o documento {ordem}: não existe. Gravar não foi enviado.')
        else:
            rejeicao_estavel = None

        status = linha.get("status", "")
        terminal = (not bloqueado and not estado.get("rascunho") and mesma_ov and
                    estado.get("grade_inicial") and
                    not mensagem_documento_inexistente_vf01(mensagem, ordem) and
                    status in {"faturamento completo", "faturamento incompleto"})
        if terminal:
            if status != estado_estavel:
                estado_estavel, estavel_desde = status, agora
            mudou_linha = (linha_antes.get("status") != status and
                _normalizar_numero_documento_sap(linha_antes.get("ov")) ==
                _normalizar_numero_documento_sap(ordem))
            estavel = agora - estavel_desde >= TEMPO_VF01_ESTABILIDADE / 1000
            # Se o status já existia antes de F8, aguarda a janela de transição.
            # Recriar o rodapé não autoriza abandonar um documento que ainda abrirá.
            if estavel and (mudou_linha or agora >= limite):
                completo = status == "faturamento completo"
                return _retorno_vf01(
                    estado, ordem,
                    "SEM_NOVO_FATURAMENTO" if completo else "FATURAMENTO_INCOMPLETO",
                    "F8_E_STATUS_COMPLETO" if completo else "F8_E_STATUS_INCOMPLETO",
                    f"OV {ordem}: {status} na própria linha. " +
                    ("Consultando a nota existente na VA03." if completo else
                     "OV mantida pendente, sem nova gravação e sem consulta VA03."),
                )
        else:
            estado_estavel, estavel_desde = "", None
        if agora >= limite:
            raise RuntimeError(
                f"F8 sem resultado confirmado para OV {ordem}. "
                "Tela preservada; novas escritas bloqueadas. "
                f"Linha: {linha}; ocupado: {bloqueado}; mensagem: {mensagem}"
            )
        pagina.wait_for_timeout(250)

    FIORI_PENDENCIA_GRAVACAO["fase"] = "GRAVACAO_SEM_CONFIRMACAO"
    mensagem_antes_salvar = estado.get("mensagem", "")
    if estado.get("gravar") is not None:
        print("[FIORI VF01] Documento novo reconhecido; acionando Gravar uma vez.")
        estado["gravar"].click(timeout=10000)
        metodo = "F8_E_BOTAO_GRAVAR"
    elif estado.get("tipo") is not None:
        print("[FIORI VF01] Documento novo reconhecido; acionando Ctrl+S uma vez.")
        estado["tipo"].press("Control+s", timeout=10000)
        metodo = "F8_E_CTRL_S"
    else:
        raise RuntimeError("Documento novo sem comando de gravação disponível.")
    limite = time.monotonic() + TEMPO_VF01_CONFIRMACAO / 1000
    atual = {}
    while time.monotonic() < limite:
        atual = _estado_atual_vf01(pagina, ordem)
        mensagem = atual.get("mensagem", "")
        cenario_gravacao = _classificar_gravacao_vf01(mensagem)
        if (mensagem and mensagem != mensagem_antes_salvar and
                not atual.get("rascunho") and not atual.get("ocupado", True) and
                not _sap_busy_indicator_visivel(pagina) and not _sap_blocklayer_popup_visivel(pagina) and
                _extrair_documento_gravado_vf01(mensagem) and
                _extrair_documento_gravado_vf01(mensagem) != _extrair_documento_gravado_vf01(mensagem_antes_salvar) and
                cenario_gravacao):
            return _retorno_vf01(atual, ordem, cenario_gravacao, metodo)
        pagina.wait_for_timeout(350)
    raise RuntimeError(
        "Gravar foi enviado uma vez, mas não houve confirmação com número de documento. "
        "Não será repetido nem descartado o documento. "
        f"Mensagem: {atual.get('mensagem', '')}"
    )

def _sap_blocklayer_popup_visivel(page_fiori):
    """
    Retorna True quando a camada modal do SAP UI5 está visível.

    Observação TESTE 34:
    #sap-ui-blocklayer-popup não significa necessariamente "popup de mensagem".
    O SAP também usa essa camada enquanto a Página inicial está carregando.
    """
    try:
        bloqueio = page_fiori.locator("#sap-ui-blocklayer-popup").first
        return bloqueio.is_visible(timeout=250)
    except Exception:
        return False


def _sap_busy_indicator_visivel(page_fiori):
    """
    Detecta indicadores de carregamento do SAP UI5/Fiori.

    O print da rodada TESTE 32 mostrou a Página inicial escurecida com três
    círculos azuis no centro. Nessa situação não devemos pressionar Escape,
    clicar em OK nem reconstruir o Shell imediatamente: devemos aguardar o
    carregamento terminar.
    """
    seletores = [
        ".sapUiLocalBusyIndicator",
        ".sapUiLocalBusyIndicatorAnimation",
        ".sapMBusyIndicator",
        "[role='progressbar']",
        "[aria-busy='true']",
        "[class*='BusyIndicator']",
        "[class*='busyIndicator']",
        "#sap-ui-static .sapUiLocalBusyIndicator",
    ]

    for seletor in seletores:
        try:
            itens = page_fiori.locator(seletor)
            qtd = min(itens.count(), 20)
        except Exception:
            qtd = 0

        for i in range(qtd):
            try:
                if itens.nth(i).is_visible(timeout=150):
                    return True
            except Exception:
                continue

    return False


def _diagnosticar_popup_sap_fiori(page_fiori):
    """Obtém texto/botões de um popup SAP real, sem confundir com busy overlay."""
    diagnostico = {"texto": "", "botoes": []}
    seletores_popup = [
        "#sap-ui-static [role='dialog']",
        "#sap-ui-static .sapMDialog",
        "#sap-ui-static .sapMMessageBox",
        "#sap-ui-static .sapMPopover",
        "#sap-ui-static .sapUiPopupWithPadding",
    ]

    for seletor in seletores_popup:
        try:
            itens = page_fiori.locator(seletor)
            quantidade = min(itens.count(), 20)
        except Exception:
            quantidade = 0

        for indice in range(quantidade):
            try:
                item = itens.nth(indice)
                if not item.is_visible(timeout=200):
                    continue

                texto_popup = re.sub(
                    r"\s+",
                    " ",
                    str(item.inner_text(timeout=1000) or ""),
                ).strip()
                if texto_popup:
                    diagnostico["texto"] = texto_popup[:1200]

                botoes = item.locator(
                    "button, [role='button'], input[type='button'], input[type='submit']"
                )
                try:
                    qtd_botoes = min(botoes.count(), 30)
                except Exception:
                    qtd_botoes = 0

                nomes = []
                for j in range(qtd_botoes):
                    try:
                        botao = botoes.nth(j)
                        if not botao.is_visible(timeout=150):
                            continue
                        nome = (
                            botao.get_attribute("aria-label")
                            or botao.get_attribute("title")
                            or botao.get_attribute("value")
                            or botao.inner_text(timeout=500)
                            or ""
                        )
                        nome = re.sub(r"\s+", " ", str(nome)).strip()
                        if nome and nome not in nomes:
                            nomes.append(nome)
                    except Exception:
                        continue

                diagnostico["botoes"] = nomes[:20]
                return diagnostico
            except Exception:
                continue

    return diagnostico


def _aguardar_carregamento_bloqueante_sap(
    page_fiori,
    motivo="",
    timeout=TEMPO_FIORI_CARREGAMENTO_BLOQUEANTE,
):
    """
    Aguarda passivamente o busy overlay do Fiori desaparecer.

    Não pressiona Escape, não clica e não navega enquanto a tela parece estar
    apenas carregando. Isso evita interromper o carregamento normal do SAP.
    """
    inicio = time.monotonic()
    limite = inicio + max(1.0, timeout / 1000.0)
    ultimo_log = -999

    while time.monotonic() < limite:
        blocklayer = _sap_blocklayer_popup_visivel(page_fiori)
        busy = _sap_busy_indicator_visivel(page_fiori)
        diag = _diagnosticar_popup_sap_fiori(page_fiori)

        # Se apareceu conteúdo/botão de um diálogo real, não é mais "só loading".
        popup_real = bool(diag.get("texto") or diag.get("botoes"))

        if not blocklayer:
            print(
                "[FIORI] Camada de carregamento liberada"
                + (f" durante {motivo}" if motivo else "")
                + "."
            )
            return True

        if popup_real and not busy:
            print(
                "[FIORI] A camada deixou de parecer carregamento e passou a "
                "exibir um popup real; tratamento de popup será usado."
            )
            return False

        decorrido = time.monotonic() - inicio
        segundo = int(decorrido)
        if segundo == 0 or segundo - ultimo_log >= 5:
            ultimo_log = segundo
            print(
                "[FIORI] SAP ainda carregando"
                + (f" durante {motivo}" if motivo else "")
                + f"... {decorrido:.1f}s/{timeout/1000:.0f}s "
                f"(blocklayer={'SIM' if blocklayer else 'NÃO'}; "
                f"busy={'SIM' if busy else 'NÃO'})."
            )

        page_fiori.wait_for_timeout(500)

    print(
        "[FIORI] A camada de carregamento permaneceu ativa além do limite "
        f"de {int(timeout/1000)} s."
    )
    return False


def recuperar_popup_bloqueador_sap_fiori(
    page_fiori,
    motivo="",
    timeout=TEMPO_FIORI_RECUPERACAO_POPUP,
    permitir_reload=True,
):
    """
    Recupera uma camada bloqueadora do SAP sem confundir carregamento com popup.

    TESTE 34:
    1) se houver blocklayer + busy indicator, aguarda até 60 s SEM tocar na tela;
    2) se a camada desaparecer, segue normalmente;
    3) somente se houver diálogo real (texto/botões) aplica Escape/botões neutros;
    4) se o carregamento travar por mais de 60 s, reconstrói Shell-home uma vez;
    5) nunca usa force=True através da camada.
    """
    if not _sap_blocklayer_popup_visivel(page_fiori):
        return True

    diag = _diagnosticar_popup_sap_fiori(page_fiori)
    texto_popup = diag.get("texto", "")
    botoes_popup = diag.get("botoes", [])
    busy = _sap_busy_indicator_visivel(page_fiori)

    # O screenshot da rodada 32 caiu exatamente aqui: tela escurecida + spinner,
    # sem mensagem funcional para o usuário.
    parece_carregamento = busy or (not texto_popup and not botoes_popup)

    if parece_carregamento:
        print(
            "[FIORI] Camada bloqueadora detectada, mas o estado parece ser "
            "CARREGAMENTO do SAP, não popup de mensagem."
        )

        if _aguardar_carregamento_bloqueante_sap(
            page_fiori,
            motivo=motivo or "estabilização do Shell",
            timeout=TEMPO_FIORI_CARREGAMENTO_BLOQUEANTE,
        ):
            return True

        # Depois da espera, reavalia: pode ter virado um diálogo real.
        diag = _diagnosticar_popup_sap_fiori(page_fiori)
        texto_popup = diag.get("texto", "")
        botoes_popup = diag.get("botoes", [])
        busy = _sap_busy_indicator_visivel(page_fiori)

        if _sap_blocklayer_popup_visivel(page_fiori) and not (
            texto_popup or botoes_popup
        ):
            if not permitir_reload:
                return False

            print(
                "[FIORI] Carregamento do SAP parece travado. "
                "Reconstruindo Shell-home uma única vez..."
            )
            try:
                page_fiori.goto(
                    FIORI_URL,
                    wait_until="domcontentloaded",
                    timeout=TEMPO_FIORI_NAVEGACAO,
                )
            except Exception as e:
                print(
                    "[FIORI] Aviso ao reconstruir Shell-home após loading "
                    f"prolongado: {e}"
                )

            # Após navegar, aguarda a camada liberar antes de considerar home pronto.
            _aguardar_carregamento_bloqueante_sap(
                page_fiori,
                motivo="reconstrução do Shell-home",
                timeout=30000,
            )

            limite_home = time.monotonic() + 25
            while time.monotonic() < limite_home:
                if not _sap_blocklayer_popup_visivel(page_fiori):
                    desc, controle = localizar_controle_pesquisa_fiori_visivel(
                        page_fiori
                    )
                    if controle is not None:
                        print(
                            f"[FIORI] Shell-home recuperado pelo {desc} "
                            "após loading prolongado."
                        )
                        return True
                page_fiori.wait_for_timeout(300)

            return False

    # A partir daqui existe indício concreto de popup real.
    print(
        "[FIORI] Popup/modal SAP real bloqueando o Shell detectado"
        + (f" durante {motivo}" if motivo else "")
        + "."
    )
    if texto_popup:
        print(f"[FIORI] Conteúdo do popup: {texto_popup[:700]}")
    if botoes_popup:
        print(f"[FIORI] Botões visíveis no popup: {botoes_popup}")

    limite = time.monotonic() + max(1.0, timeout / 1000.0)

    # Só popup real recebe Escape.
    for tentativa_escape in range(1, 3):
        try:
            page_fiori.keyboard.press("Escape")
            page_fiori.wait_for_timeout(600)
        except Exception:
            pass

        if not _sap_blocklayer_popup_visivel(page_fiori):
            print(
                f"[FIORI] Popup real removido com Escape "
                f"(tentativa {tentativa_escape}/2)."
            )
            return True

    # Só comandos neutros, nunca botões de negócio.
    padrao_neutro = re.compile(
        r"^\s*(?:Fechar|Close|Cancelar|Cancel|OK|Ok)\s*$",
        re.I,
    )

    try:
        area = page_fiori.locator("#sap-ui-static")
        candidatos = [
            area.get_by_role("button", name=padrao_neutro),
            area.get_by_text(padrao_neutro, exact=True),
            area.locator(
                "input[type='button'], input[type='submit'], button, [role='button']"
            ).filter(has_text=padrao_neutro),
        ]

        for candidato in candidatos:
            try:
                qtd = min(candidato.count(), 20)
            except Exception:
                qtd = 0

            for i in range(qtd):
                try:
                    alvo = candidato.nth(i)
                    if not alvo.is_visible(timeout=250):
                        continue

                    alvo.click(timeout=5000)
                    page_fiori.wait_for_timeout(700)

                    if not _sap_blocklayer_popup_visivel(page_fiori):
                        print(
                            "[FIORI] Popup real fechado por botão neutro."
                        )
                        return True
                except Exception:
                    continue
    except Exception:
        pass

    while time.monotonic() < limite:
        if not _sap_blocklayer_popup_visivel(page_fiori):
            print(
                "[FIORI] Popup bloqueador desapareceu durante a espera de "
                "estabilização."
            )
            return True
        page_fiori.wait_for_timeout(300)

    if not permitir_reload:
        return False

    print(
        "[FIORI] Popup real continuou bloqueando o Shell; reconstruindo "
        "Shell-home uma única vez."
    )
    try:
        page_fiori.goto(
            FIORI_URL,
            wait_until="domcontentloaded",
            timeout=TEMPO_FIORI_NAVEGACAO,
        )
    except Exception as e:
        print(f"[FIORI] Aviso ao reconstruir Shell-home após popup: {e}")

    # Depois do goto, trata eventual loading sem pressionar Escape.
    _aguardar_carregamento_bloqueante_sap(
        page_fiori,
        motivo="reconstrução após popup",
        timeout=30000,
    )

    limite_home = time.monotonic() + 20
    while time.monotonic() < limite_home:
        if not _sap_blocklayer_popup_visivel(page_fiori):
            descricao, controle = localizar_controle_pesquisa_fiori_visivel(
                page_fiori
            )
            if controle is not None:
                print(
                    f"[FIORI] Shell-home reconstruído e liberado pelo "
                    f"{descricao}."
                )
                return True
        page_fiori.wait_for_timeout(300)

    print(
        "[FIORI] Não foi possível liberar a camada bloqueadora do SAP "
        "automaticamente."
    )
    return False


def _clicar_primeiro_locator_visivel(locators, descricao, timeout=TEMPO_FIORI_ELEMENTO):
    """Clica no primeiro locator visível dentre alternativas equivalentes."""
    erros = []
    limite = time.time() + (timeout / 1000)

    while time.time() < limite:
        for locator in locators:
            try:
                alvo = locator.first
                if not alvo.is_visible(timeout=350):
                    continue
                alvo.scroll_into_view_if_needed(timeout=3000)
                # O carregamento será validado pelos marcadores da aplicação.
                # Não deixamos o clique preso aguardando uma navegação implícita.
                alvo.click(timeout=10000, no_wait_after=True)
                return alvo
            except Exception as e:
                erros.append(str(e))
        time.sleep(0.25)

    detalhe = " | ".join(erros[-2:])
    raise RuntimeError(f"Não localizei {descricao}. {detalhe}".strip())


def _aplicativo_fiori_visivel(page_fiori):
    """Um iframe WebGUI ainda visível não é a página inicial, mesmo com hash home."""
    for frame in page_fiori.frames:
        if frame.parent_frame is None or not _frame_fiori_realmente_visivel(frame):
            continue
        nome = str(frame.name or "").lower()
        url = str(frame.url or "").lower()
        if (nome.startswith("application-") or "/sap/bc/gui/" in url):
            try:
                if frame.locator("body").inner_text(timeout=500).strip():
                    return True
            except Exception:
                return True
    return False


def _shell_home_fiori_confirmado(page_fiori):
    if "#shell-home" not in str(page_fiori.url or "").lower():
        return False
    if (_sap_blocklayer_popup_visivel(page_fiori) or
            _sap_busy_indicator_visivel(page_fiori) or
            _aplicativo_fiori_visivel(page_fiori)):
        return False
    _, controle = localizar_controle_pesquisa_fiori_visivel(page_fiori)
    return controle is not None


def _impedir_descarte_faturamento_fiori(page_fiori):
    if FIORI_PENDENCIA_GRAVACAO:
        raise RuntimeError(
            "Navegação bloqueada: há execução/gravação VF01 sem confirmação. "
            "Confira a tela preservada e o log antes de reiniciar."
        )
    estado = _estado_atual_vf01(page_fiori)
    if estado.get("faturamento") and estado.get("rascunho"):
        FIORI_PENDENCIA_GRAVACAO["fase"] = "RASCUNHO_ANTES_DA_NAVEGACAO"
        raise RuntimeError(
            "Documento temporário de faturamento ainda aberto. "
            "Não será descartado por navegação/reload."
        )


def _aguardar_shell_home_fiori(page_fiori, timeout):
    """Confirma URL + pesquisa + ausência de app visível e de bloqueador."""
    limite = time.time() + (timeout / 1000)
    while time.time() < limite:
        if _shell_home_fiori_confirmado(page_fiori):
            return True
        page_fiori.wait_for_timeout(250)
    return False


def fechar_pesquisa_global_fiori(page_fiori, timeout=3000):
    """Fecha pesquisa/popup residual do Launchpad e limpa o termo anterior."""
    _, pesquisa = localizar_campo_pesquisa_fiori_visivel(page_fiori)

    if pesquisa is not None:
        try:
            pesquisa.fill("")
        except Exception:
            pass

        try:
            page_fiori.keyboard.press("Escape")
            pesquisa.wait_for(state="hidden", timeout=timeout)
        except Exception:
            # Em algumas versões o campo permanece expandido; termo vazio já
            # elimina a pesquisa residual.
            try:
                if str(pesquisa.input_value(timeout=1000) or "").strip():
                    return False
            except Exception:
                pass

    # TESTE 33: pesquisa fechada não significa tela livre. A rodada 32 deixou
    # #sap-ui-blocklayer-popup ativo mesmo sem searchbox visível.
    if _sap_blocklayer_popup_visivel(page_fiori):
        return recuperar_popup_bloqueador_sap_fiori(
            page_fiori,
            motivo="limpeza da pesquisa global",
            timeout=max(5000, timeout),
            permitir_reload=False,
        )

    return True

def garantir_shell_fiori(page_fiori, forcar_home=False):
    """Disponibiliza a pesquisa do Fiori a partir do Shell-home real."""
    page_fiori.bring_to_front()
    _impedir_descarte_faturamento_fiori(page_fiori)

    # TESTE 36:
    # O campo global de pesquisa continua visível mesmo dentro de aplicativos
    # como VA03/VF01. Isso NÃO significa que estamos no Shell-home. A rodada 35
    # mostrou pesquisa "vf01/Exibir ordens do cliente" sendo feita enquanto a
    # tela de documento anterior ainda estava aberta.
    #
    # Regra: toda troca de transação parte de #Shell-home.
    esta_no_shell_home = _shell_home_fiori_confirmado(page_fiori)
    if not esta_no_shell_home:
        if not forcar_home:
            print(
                "[FIORI] Pesquisa global está disponível, mas o SAP ainda está "
                "dentro de um aplicativo. Voltando ao Shell-home antes da "
                "próxima transação."
            )
        forcar_home = True

    # Se a pesquisa já estiver aberta NO SHELL-HOME, ela é o controle correto.
    # clicar novamente no botão do cabeçalho por trás do popup de pesquisa.
    descricao_campo, campo_aberto = localizar_campo_pesquisa_fiori_visivel(page_fiori)
    if not forcar_home and esta_no_shell_home and campo_aberto is not None:
        if _sap_blocklayer_popup_visivel(page_fiori):
            print(
                "[FIORI] Campo de pesquisa existe, mas o SAP ainda está com "
                "camada bloqueadora; aguardando estabilização antes de usá-lo."
            )
            if not recuperar_popup_bloqueador_sap_fiori(
                page_fiori,
                motivo="campo de pesquisa já aberto",
                timeout=8000,
                permitir_reload=True,
            ):
                raise RuntimeError(
                    "O campo de pesquisa existe, mas a camada bloqueadora do "
                    "SAP não foi liberada."
                )
            descricao_campo, campo_aberto = localizar_campo_pesquisa_fiori_visivel(
                page_fiori
            )

        if campo_aberto is not None:
            print(f"[FIORI] Shell pesquisável já disponível pelo {descricao_campo}.")
            return campo_aberto

    # Se não há campo utilizável e existe blocklayer, primeiro diferencia loading de popup.
    if not forcar_home and _sap_blocklayer_popup_visivel(page_fiori):
        if not recuperar_popup_bloqueador_sap_fiori(
            page_fiori,
            motivo="preparação do Shell",
            timeout=8000,
            permitir_reload=True,
        ):
            raise RuntimeError(
                "O SAP manteve um popup modal bloqueando a pesquisa global."
            )

    if not forcar_home and esta_no_shell_home:
        descricao, controle = localizar_controle_pesquisa_fiori_visivel(page_fiori)
        if controle is not None:
            print(f"[FIORI] Shell pesquisável já disponível pelo {descricao}.")
            return controle

    esperas = (8000, 12000, 15000)
    ultimo_erro = ""

    for tentativa, espera in enumerate(esperas, start=1):
        print(
            f"[FIORI] Recuperando Shell-home automaticamente "
            f"(tentativa {tentativa}/{len(esperas)})..."
        )
        try:
            if tentativa == 2:
                try:
                    page_fiori.keyboard.press("Escape")
                except Exception:
                    pass

            if tentativa < 3:
                page_fiori.goto(
                    FIORI_URL,
                    wait_until="domcontentloaded",
                    timeout=TEMPO_FIORI_NAVEGACAO,
                )
            else:
                page_fiori.reload(
                    wait_until="domcontentloaded",
                    timeout=TEMPO_FIORI_NAVEGACAO,
                )
        except Exception as e:
            ultimo_erro = str(e)
            print(f"[FIORI] Aviso na recuperação automática: {e}")

        if _sap_blocklayer_popup_visivel(page_fiori):
            if not recuperar_popup_bloqueador_sap_fiori(
                page_fiori, motivo=f"carregamento Shell-home tentativa {tentativa}",
                timeout=6000, permitir_reload=False,
            ):
                ultimo_erro = "Camada SAP não liberada; home ainda não confirmado."
                continue

        if _aguardar_shell_home_fiori(page_fiori, timeout=espera):
            # A URL/home pode carregar com um popup residual por cima. Não
            # consideramos o Shell pronto enquanto a camada modal interceptar.
            if _sap_blocklayer_popup_visivel(page_fiori):
                recuperou_popup = recuperar_popup_bloqueador_sap_fiori(
                    page_fiori,
                    motivo=f"recuperação Shell-home tentativa {tentativa}",
                    timeout=6000,
                    permitir_reload=False,
                )
                if not recuperou_popup:
                    ultimo_erro = "Shell-home carregou com popup modal bloqueador."
                    continue

            descricao, controle = localizar_controle_pesquisa_fiori_visivel(page_fiori)
            if controle is not None and _shell_home_fiori_confirmado(page_fiori):
                print(f"[FIORI] Shell-home confirmado pelo {descricao}.")
                return controle

    raise RuntimeError(
        "O Fiori não disponibilizou o Shell-home após três recuperações "
        "automáticas. A autenticação do SAP é validada somente na "
        "inicialização e não será solicitada novamente durante a fila. "
        f"Último detalhe: {ultimo_erro or 'sem detalhe adicional'}."
    )

def preparar_fiori_nova_requisicao(page_fiori, requisicao):
    """Normaliza uma vez o Launchpad antes da primeira OV da requisição."""
    requisicao = str(requisicao).strip()
    print(
        f"[FIORI] Preparando estado limpo para a requisição {requisicao}..."
    )

    page_fiori.bring_to_front()
    try:
        page_fiori.keyboard.press("Escape")
    except Exception:
        pass

    garantir_shell_fiori(page_fiori, forcar_home=True)

    if not fechar_pesquisa_global_fiori(page_fiori):
        print(
            "[FIORI] A pesquisa global permaneceu aberta; "
            "recarregando o Shell-home uma única vez..."
        )
        try:
            page_fiori.reload(
                wait_until="domcontentloaded",
                timeout=TEMPO_FIORI_NAVEGACAO,
            )
        except Exception as e:
            print(f"[FIORI] Aviso durante a limpeza da pesquisa: {e}")

        if not _aguardar_shell_home_fiori(
            page_fiori,
            timeout=TEMPO_FIORI_ELEMENTO,
        ):
            raise RuntimeError(
                "O Fiori não confirmou o Shell-home ao iniciar a nova requisição."
            )

    descricao, controle = aguardar_controle_pesquisa_fiori(
        page_fiori,
        timeout=5000,
    )
    if controle is None:
        raise RuntimeError(
            "O Fiori perdeu o controle de pesquisa ao iniciar a nova requisição."
        )

    print(
        f"[FIORI] Requisição {requisicao}: estado inicial confirmado pelo "
        f"{descricao}."
    )


def abrir_aplicativo_fiori(
    page_fiori,
    termo_pesquisa,
    nomes_aplicativo,
    tipo_aplicativo,
):
    """
    Abre VF01/VA03 pela pesquisa do Launchpad na mesma aba Fiori validada.

    A tela só é aceita quando o marcador específico do aplicativo estiver
    visível nesta mesma aba.
    """
    tipo_aplicativo = str(tipo_aplicativo or "").strip().lower()
    termo_base = str(termo_pesquisa or "").strip()
    termos_por_tentativa = [termo_base, termo_base]

    # Quando a busca curta por VA03 não devolve o cartão do aplicativo, a
    # segunda tentativa usa o título funcional exibido pelo Launchpad.
    if tipo_aplicativo == "va03":
        termos_por_tentativa[1] = "Exibir ordens do cliente"

    ultimo_erro = None

    for tentativa_abertura in range(1, 3):
        termo_atual = termos_por_tentativa[tentativa_abertura - 1]
        try:
            print(
                f"[FIORI] Pesquisando {tipo_aplicativo.upper()} por "
                f"'{termo_atual}' (tentativa {tentativa_abertura}/2)..."
            )
            controle_pesquisa = garantir_shell_fiori(
                page_fiori,
                forcar_home=tentativa_abertura > 1,
            )

            _, pesquisa = localizar_campo_pesquisa_fiori_visivel(page_fiori)
            if pesquisa is None:
                # TESTE 33: antes de clicar no botão de pesquisa, elimina popup
                # residual que possa estar interceptando o ponteiro. Nunca usa
                # force=True através de #sap-ui-blocklayer-popup.
                if _sap_blocklayer_popup_visivel(page_fiori):
                    if not recuperar_popup_bloqueador_sap_fiori(
                        page_fiori,
                        motivo=f"abertura de {tipo_aplicativo.upper()}",
                        timeout=8000,
                        permitir_reload=True,
                    ):
                        raise RuntimeError(
                            f"Popup SAP permaneceu bloqueando a abertura de "
                            f"{tipo_aplicativo.upper()}."
                        )

                    # O DOM do Shell pode ter sido recriado; não reutiliza locator velho.
                    controle_pesquisa = garantir_shell_fiori(
                        page_fiori,
                        forcar_home=False,
                    )
                    _, pesquisa = localizar_campo_pesquisa_fiori_visivel(page_fiori)

                if pesquisa is None:
                    controle_pesquisa.click(timeout=TEMPO_FIORI_ELEMENTO)
                    _, pesquisa = aguardar_campo_pesquisa_fiori(
                        page_fiori,
                        timeout=TEMPO_FIORI_ELEMENTO,
                    )

            if pesquisa is None:
                raise RuntimeError(
                    "O Shell-home foi reconhecido, mas o campo de pesquisa "
                    "não abriu."
                )

            pesquisa.fill("")
            pesquisa.fill(termo_atual)
            pesquisa.press("Enter")
            page_fiori.wait_for_timeout(700)

            candidatos = []
            for nome in nomes_aplicativo:
                candidatos.extend([
                    page_fiori.get_by_role("link", name=nome),
                    page_fiori.get_by_role("button", name=nome),
                    page_fiori.get_by_label(nome),
                    page_fiori.get_by_title(nome),
                    page_fiori.get_by_text(nome),
                ])

            try:
                _clicar_primeiro_locator_visivel(
                    candidatos,
                    descricao=f"o aplicativo pesquisado por '{termo_atual}'",
                    timeout=12000,
                )
            except Exception as primeira_busca:
                # O shell eventualmente mantém apenas o texto consultado e não
                # materializa os resultados. Redigitar dispara todos os eventos
                # do campo antes de reenviar a mesma pesquisa.
                print(
                    f"[FIORI] O resultado de '{termo_atual}' não carregou; "
                    "reenviando a pesquisa uma vez..."
                )
                pesquisa.click(timeout=5000)
                pesquisa.press("Control+a")
                pesquisa.type(termo_atual, delay=60)
                pesquisa.press("Enter")
                page_fiori.wait_for_timeout(1200)
                try:
                    _clicar_primeiro_locator_visivel(
                        candidatos,
                        descricao=(
                            f"o aplicativo pesquisado por '{termo_atual}'"
                        ),
                        timeout=15000,
                    )
                except Exception as segunda_busca:
                    raise RuntimeError(
                        f"Não localizei o aplicativo pesquisado por "
                        f"'{termo_atual}' após reenviar a busca. "
                        f"Primeira: {primeira_busca}; segunda: {segunda_busca}"
                    ) from segunda_busca

            # Uma pesquisa ainda visível significa que o clique não abriu o app.
            # Sem essa trava, o iframe anterior poderia ser aceito por engano.
            pesquisa.wait_for(state="hidden", timeout=8000)

            return localizar_contexto_aplicativo_fiori(
                page_fiori,
                tipo_aplicativo=tipo_aplicativo,
                timeout=TEMPO_FIORI_NAVEGACAO,
            )

        except Exception as e:
            ultimo_erro = e
            if tentativa_abertura < 2:
                print(
                    f"[FIORI] A abertura de {tipo_aplicativo.upper()} não foi "
                    "confirmada; limpando a pesquisa/popup e repetindo uma vez..."
                )

                # Se a tentativa deixou um popup UI5 por cima do Shell, remove-o
                # antes de procurar a pesquisa novamente. Esse é exatamente o
                # cenário registrado na rodada TESTE 32.
                if _sap_blocklayer_popup_visivel(page_fiori):
                    recuperar_popup_bloqueador_sap_fiori(
                        page_fiori,
                        motivo=f"falha ao abrir {tipo_aplicativo.upper()}",
                        timeout=8000,
                        permitir_reload=True,
                    )

                fechar_pesquisa_global_fiori(page_fiori)
                continue

    raise RuntimeError(
        f"Não consegui abrir {tipo_aplicativo.upper()} após duas tentativas. "
        f"Último detalhe: {ultimo_erro}"
    )


def capturar_mensagem_fiori(frame_fiori):
    """Captura, em melhor esforço, mensagens de status/alerta do SAP WebGUI."""
    textos = []
    seletores = [
        "[role='alert']",
        "[role='status']",
        "[aria-live='assertive']",
        "[aria-live='polite']",
    ]

    for seletor in seletores:
        try:
            itens = frame_fiori.locator(seletor)
            for indice in range(min(itens.count(), 20)):
                item = itens.nth(indice)
                if not item.is_visible():
                    continue
                texto = str(item.inner_text(timeout=1500) or "").strip()
                if texto and texto not in textos:
                    textos.append(texto)
        except Exception:
            continue

    return " | ".join(textos[-3:])[:1000]


def salvar_evidencia_fiori(page_fiori, prefixo, requisicao, ordem):
    """Salva print de auditoria sem interromper o processamento em caso de falha."""
    data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome = f"{prefixo}_{requisicao}_{ordem}_{data_hora}.png"
    caminho = os.path.join(criar_pasta_requisicao(requisicao), nome)
    try:
        page_fiori.screenshot(path=caminho, full_page=True)
        return caminho
    except Exception as e:
        print(f"[FIORI] Não consegui salvar a evidência {nome}: {e}")
        return ""


PADRAO_BOTAO_FLUXO_DOCUMENTOS = re.compile(
    r"^\s*Exibir\s+fluxo\s+de\s+documentos\s*$",
    re.I,
)
PADRAO_TITULO_FLUXO_DOCUMENTOS = re.compile(
    r"^\s*Fluxo\s+de\s+documentos\s*$",
    re.I,
)


def _paginas_fiori_priorizadas(page_referencia):
    """Retorna a página atual e as demais abas, priorizando as mais recentes."""
    paginas = []

    try:
        if page_referencia and not page_referencia.is_closed():
            paginas.append(page_referencia)
    except Exception:
        pass

    try:
        outras = [
            pagina
            for pagina in page_referencia.context.pages
            if pagina not in paginas and not pagina.is_closed()
        ]
        paginas.extend(reversed(outras))
    except Exception:
        pass

    return paginas


def _primeiro_item_visivel(locators, limite_itens=10):
    """Retorna o primeiro elemento visível entre localizadores equivalentes."""
    for locator in locators:
        for indice in range(limite_itens):
            try:
                item = locator.nth(indice)
                if item.is_visible():
                    return item
            except Exception:
                break
    return None


def _contexto_tem_titulo_fluxo_documentos(pagina, frame_fiori):
    """Confirma o título real do fluxo na página ou dentro do frame."""
    locators = [
        frame_fiori.get_by_role(
            "heading",
            name=PADRAO_TITULO_FLUXO_DOCUMENTOS,
        ),
        frame_fiori.get_by_text(
            PADRAO_TITULO_FLUXO_DOCUMENTOS,
            exact=True,
        ),
    ]

    try:
        locators.extend([
            pagina.get_by_role(
                "heading",
                name=PADRAO_TITULO_FLUXO_DOCUMENTOS,
            ),
            pagina.get_by_text(
                PADRAO_TITULO_FLUXO_DOCUMENTOS,
                exact=True,
            ),
        ])
    except Exception:
        pass

    return _primeiro_item_visivel(locators, limite_itens=5) is not None


def _frame_contem_fluxo_documentos(frame_fiori):
    """Confirma a tela do fluxo pelas colunas Documento e Status."""
    padrao_documento = re.compile(r"^\s*Documento\s*$", re.I)
    padrao_status = re.compile(r"^\s*Status\s*$", re.I)

    cabecalho_documento = _primeiro_item_visivel([
        frame_fiori.get_by_role("columnheader", name=padrao_documento),
        frame_fiori.locator("div").filter(has_text=padrao_documento),
        frame_fiori.get_by_text(padrao_documento, exact=True),
    ])
    if cabecalho_documento is None:
        return False

    cabecalho_status = _primeiro_item_visivel([
        frame_fiori.get_by_role("columnheader", name=padrao_status),
        frame_fiori.locator("div").filter(has_text=padrao_status),
        frame_fiori.get_by_text(padrao_status, exact=True),
    ])
    return cabecalho_status is not None


def localizar_contexto_fluxo_documentos_fiori(
    page_referencia,
    frame_resumo=None,
    timeout=TEMPO_FIORI_NAVEGACAO,
):
    """
    Localiza o Fluxo de documentos em qualquer aba/frame aberto pelo SAP.

    O clique da VA03 pode trocar o conteúdo do frame atual ou abrir outro
    contexto. A tela correta é reconhecida pelas colunas Documento e Status.
    """
    limite = time.time() + (timeout / 1000)
    ultimo_mapa_frames = []

    while time.time() < limite:
        mapa_frames = []

        for pagina in _paginas_fiori_priorizadas(page_referencia):
            try:
                frames = list(reversed(pagina.frames))
            except Exception:
                continue

            for frame in frames:
                try:
                    if frame.is_detached():
                        continue
                except Exception:
                    pass

                try:
                    mapa_frames.append(
                        f"nome={frame.name or '<principal>'}; url={frame.url[:160]}"
                    )
                except Exception:
                    pass

                try:
                    if not _frame_contem_fluxo_documentos(frame):
                        continue

                    # Impede o falso positivo observado na AV3: a tela-resumo
                    # mantinha elementos Documento/Status no DOM, embora o botão
                    # "Exibir fluxo" ainda estivesse visível e o clique não tivesse
                    # mudado de tela.
                    if frame_resumo is not None and frame == frame_resumo:
                        botao_resumo = _primeiro_item_visivel([
                            frame.get_by_role(
                                "button",
                                name=PADRAO_BOTAO_FLUXO_DOCUMENTOS,
                            )
                        ])
                        if botao_resumo is not None:
                            continue
                    elif not _contexto_tem_titulo_fluxo_documentos(pagina, frame):
                        continue

                    pagina.bring_to_front()
                    print(
                        "[FIORI VA03] Fluxo de documentos confirmado no frame "
                        f"'{frame.name or '<principal>'}'."
                    )
                    return pagina, frame
                except Exception:
                    continue

        ultimo_mapa_frames = mapa_frames
        try:
            page_referencia.wait_for_timeout(300)
        except Exception:
            time.sleep(0.3)

    detalhe = " | ".join(ultimo_mapa_frames[-10:]) or "nenhum frame disponível"
    raise RuntimeError(
        "O clique foi realizado, mas a tela 'Fluxo de documentos' não foi "
        f"reconhecida em nenhuma aba/frame. Frames observados: {detalhe}"
    )


def localizar_botao_fluxo_documentos_fiori(
    page_referencia,
    frame_preferencial=None,
    timeout=TEMPO_FIORI_ELEMENTO,
):
    """Localiza o botão Exibir fluxo de documentos no contexto ainda visível."""
    limite = time.time() + (timeout / 1000)

    while time.time() < limite:
        if frame_preferencial is not None:
            try:
                botao = _primeiro_item_visivel([
                    frame_preferencial.get_by_role(
                        "button",
                        name=PADRAO_BOTAO_FLUXO_DOCUMENTOS,
                    )
                ])
                if botao is not None:
                    return page_referencia, frame_preferencial, botao
            except Exception:
                frame_preferencial = None

        for pagina in _paginas_fiori_priorizadas(page_referencia):
            try:
                frames = list(reversed(pagina.frames))
            except Exception:
                continue

            for frame in frames:
                try:
                    botao = _primeiro_item_visivel([
                        frame.get_by_role(
                            "button",
                            name=PADRAO_BOTAO_FLUXO_DOCUMENTOS,
                        )
                    ])
                    if botao is not None:
                        return pagina, frame, botao
                except Exception:
                    continue

        try:
            page_referencia.wait_for_timeout(250)
        except Exception:
            time.sleep(0.25)

    raise RuntimeError("Botão 'Exibir fluxo de documentos' não localizado.")


def localizar_botao_acionavel_em_frames_fiori(
    page_referencia,
    padrao_nome,
    descricao,
    frame_preferencial=None,
    timeout=TEMPO_FIORI_ELEMENTO,
):
    """Localiza um botão visível e habilitado em qualquer aba/frame do Fiori."""
    limite = time.time() + (timeout / 1000)
    ultimo_mapa_frames = []

    while time.time() < limite:
        contextos = []
        chaves = set()

        if frame_preferencial is not None:
            try:
                pagina_preferencial = frame_preferencial.page
                contextos.append((pagina_preferencial, frame_preferencial))
                chaves.add(id(frame_preferencial))
            except Exception:
                frame_preferencial = None

        mapa_frames = []
        for pagina in _paginas_fiori_priorizadas(page_referencia):
            try:
                frames = list(reversed(pagina.frames))
            except Exception:
                continue

            for frame in frames:
                try:
                    mapa_frames.append(
                        f"nome={frame.name or '<principal>'}; url={frame.url[:120]}"
                    )
                except Exception:
                    pass

                if id(frame) in chaves:
                    continue
                chaves.add(id(frame))
                contextos.append((pagina, frame))

        for pagina, frame in contextos:
            try:
                botoes = frame.get_by_role("button", name=padrao_nome)
                quantidade = min(botoes.count(), 10)
            except Exception:
                continue

            for indice in range(quantidade):
                try:
                    botao = botoes.nth(indice)
                    if not botao.is_visible():
                        continue
                    if not botao.is_enabled():
                        continue
                    pagina.bring_to_front()
                    print(
                        f"[FIORI VA03] {descricao} localizado e habilitado no frame "
                        f"'{frame.name or '<principal>'}'."
                    )
                    return pagina, frame, botao
                except Exception:
                    continue

        ultimo_mapa_frames = mapa_frames
        try:
            page_referencia.wait_for_timeout(250)
        except Exception:
            time.sleep(0.25)

    detalhe = " | ".join(ultimo_mapa_frames[-8:]) or "nenhum frame disponível"
    raise RuntimeError(
        f"{descricao} não ficou visível e habilitado em nenhuma aba/frame. "
        f"Frames observados: {detalhe}"
    )


def _pagina_eh_sap_fiori(pagina, pagina_referencia=None):
    """Heurística conservadora para não misturar frames do PORTAL com o SAP."""
    try:
        if pagina_referencia is not None and pagina is pagina_referencia:
            return True
    except Exception:
        pass

    try:
        url = str(pagina.url or "").lower()
    except Exception:
        url = ""

    if urlsplit(FIORI_URL).hostname and urlsplit(FIORI_URL).hostname.lower() in url or "/sap/bc/" in url:
        return True

    # Pop-ups do SAP podem nascer em about:blank e ganhar URL depois. Se o
    # opener for a aba Fiori de referência, ainda é um contexto SAP válido.
    try:
        opener = pagina.opener()
        if opener is not None:
            if pagina_referencia is not None and opener is pagina_referencia:
                return True
            url_opener = str(opener.url or "").lower()
            if urlsplit(FIORI_URL).hostname and urlsplit(FIORI_URL).hostname.lower() in url_opener or "/sap/bc/" in url_opener:
                return True
    except Exception:
        pass

    return False


def localizar_botao_apos_transicao_fiori(
    page_referencia,
    padrao_nome,
    descricao,
    timeout=45000,
    padrao_erro=None,
    descricao_erro="",
):
    """
    Reaquisição forte do SAP depois de uma navegação que pode recriar o iframe.

    Diferentemente da busca genérica, esta rotina ignora páginas do portal de atendimento
    e reconstrói a lista de frames Fiori em cada ciclo. TESTE 43 também permite
    interromper cedo a espera quando o próprio SAP devolver uma validação conhecida
    da ação recém-executada (ex.: "Selecionar um documento de vendas").
    """
    limite = time.time() + (timeout / 1000)
    proximo_log = time.time()
    ultimo_mapa = []

    while time.time() < limite:
        paginas = []
        try:
            for pagina in page_referencia.context.pages:
                try:
                    if pagina.is_closed():
                        continue
                except Exception:
                    continue
                if _pagina_eh_sap_fiori(pagina, page_referencia):
                    paginas.append(pagina)
        except Exception:
            paginas = [page_referencia]

        if page_referencia not in paginas:
            paginas.insert(0, page_referencia)
        else:
            paginas = [page_referencia] + [p for p in paginas if p is not page_referencia]

        mapa = []
        frames_validos = []
        for pagina in paginas:
            try:
                frames = list(reversed(pagina.frames))
            except Exception:
                continue

            for frame in frames:
                try:
                    if frame.is_detached():
                        continue
                except Exception:
                    pass

                frames_validos.append((pagina, frame))
                try:
                    mapa.append(
                        f"pagina={str(pagina.url or '')[:90]}; "
                        f"frame={frame.name or '<principal>'}; "
                        f"url={str(frame.url or '')[:100]}"
                    )
                except Exception:
                    pass

                locators = []
                try:
                    locators.append(frame.get_by_role("button", name=padrao_nome))
                except Exception:
                    pass
                try:
                    locators.append(frame.get_by_text(padrao_nome, exact=False))
                except Exception:
                    pass

                for locator in locators:
                    try:
                        quantidade = min(locator.count(), 10)
                    except Exception:
                        quantidade = 0
                    for indice in range(quantidade):
                        try:
                            item = locator.nth(indice)
                            if not item.is_visible():
                                continue
                            try:
                                if not item.is_enabled():
                                    continue
                            except Exception:
                                pass
                            pagina.bring_to_front()
                            print(
                                f"[FIORI VA03] {descricao} localizado após "
                                "reaquisição do contexto SAP no frame "
                                f"'{frame.name or '<principal>'}'."
                            )
                            return pagina, frame, item
                        except Exception:
                            continue

        # Só verifica a mensagem de erro depois de procurar o destino em TODOS
        # os frames. Assim uma mensagem antiga no frame anterior não vence uma
        # transição que já criou corretamente a nova tela.
        if padrao_erro is not None:
            for pagina, frame in frames_validos:
                try:
                    erros = frame.get_by_text(padrao_erro, exact=False)
                    quantidade = min(erros.count(), 8)
                except Exception:
                    quantidade = 0
                for indice in range(quantidade):
                    try:
                        item_erro = erros.nth(indice)
                        if not item_erro.is_visible():
                            continue
                        mensagem = (item_erro.inner_text(timeout=700) or "").strip()
                        raise RuntimeError(
                            "SAP_VALIDACAO_TRANSICAO: "
                            + (descricao_erro or mensagem or "validação SAP conhecida")
                        )
                    except RuntimeError:
                        raise
                    except Exception:
                        continue

        ultimo_mapa = mapa
        agora = time.time()
        if agora >= proximo_log:
            restantes = max(0, limite - agora)
            print(
                f"[FIORI VA03] Aguardando a nova tela SAP após a transição... "
                f"{restantes:.0f}s restantes."
            )
            proximo_log = agora + 5

        try:
            page_referencia.wait_for_timeout(350)
        except Exception:
            time.sleep(0.35)

    detalhe = " | ".join(ultimo_mapa[-10:]) or "nenhum frame SAP/Fiori observado"
    raise RuntimeError(
        f"{descricao} não apareceu após a transição do SAP. "
        f"Contextos Fiori observados: {detalhe}"
    )


def selecionar_faturamento_e_localizar_exibir_documento(
    page_referencia,
    frame_fluxo,
    locator_faturamento,
    descricao_faturamento="documento de faturamento",
):
    """
    Seleciona a linha EXATA do faturamento no tree SAP e só devolve o botão
    "Exibir documento" quando houver evidência de que essa linha ficou ativa.

    TESTE 43:
    - botão "Exibir documento" habilitado, sozinho, não comprova seleção;
    - aceita como evidência estrutural aria-selected=true OU a classe ativa já
      observada no WebGUI desta execução (urST4RowActive);
    - se o clique na área da row não ativar o documento, tenta uma vez o texto
      exato da própria FAT dentro da mesma linha, sem procurar outra FAT.
    """
    ultimo_erro = ""

    linha = locator_faturamento
    try:
        id_atual = locator_faturamento.get_attribute("id") or ""
    except Exception:
        id_atual = ""

    if "-Row-" not in id_atual:
        try:
            ancestral = locator_faturamento.locator(
                "xpath=ancestor-or-self::*[contains(@id,'-Row-')][1]"
            )
            if ancestral.count() > 0:
                linha = ancestral.first
        except Exception:
            pass

    try:
        id_linha = linha.get_attribute("id") or ""
    except Exception:
        id_linha = ""

    try:
        texto_linha = (linha.inner_text(timeout=1800) or "").strip()
    except Exception:
        texto_linha = descricao_faturamento

    print(
        f"[FIORI VA03] Linha física usada para seleção: "
        f"{texto_linha or descricao_faturamento}"
        + (f" | id={id_linha}" if id_linha else "")
    )

    def ler_estado_selecao():
        try:
            estado = linha.evaluate(
                """el => ({
                    id: el.id || '',
                    cls: el.className || '',
                    ariaSelected: el.getAttribute('aria-selected') || '',
                    bg: getComputedStyle(el).backgroundColor || ''
                })"""
            )
        except Exception as e_estado:
            return False, {}, str(e_estado)

        classe = str(estado.get("cls", "") or "")
        aria = str(estado.get("ariaSelected", "") or "").strip().lower()
        selecionada = aria == "true" or "urst4rowactive" in classe.lower()
        print(
            "[FIORI VA03] Estado da linha após o clique: "
            f"id={estado.get('id','')}; "
            f"aria-selected={estado.get('ariaSelected','')}; "
            f"background={estado.get('bg','')}; "
            f"class={classe[:180]}; "
            f"seleção_confirmada={'SIM' if selecionada else 'NÃO'}"
        )
        return selecionada, estado, ""

    def clicar_texto_exato_da_fat():
        padrao_texto = re.compile(
            rf"^\s*{re.escape(str(descricao_faturamento or '').strip())}\s*$",
            re.I,
        )
        try:
            candidatos = linha.get_by_text(padrao_texto, exact=False)
            quantidade = min(candidatos.count(), 8)
        except Exception:
            quantidade = 0

        for indice in range(quantidade):
            try:
                alvo = candidatos.nth(indice)
                if not alvo.is_visible(timeout=500):
                    continue
                alvo.scroll_into_view_if_needed(timeout=3000)
                alvo.click(timeout=7000)
                print(
                    "[FIORI VA03] Clique complementar realizado no texto exato "
                    "da FAT dentro da mesma linha."
                )
                return True
            except Exception:
                continue
        return False

    for tentativa in range(1, 4):
        print(
            f"[FIORI VA03] Selecionando a LINHA inteira "
            f"'{descricao_faturamento}' (tentativa {tentativa}/3)..."
        )

        try:
            linha.scroll_into_view_if_needed(timeout=7000)
            page_referencia.wait_for_timeout(250)

            caixa = linha.bounding_box(timeout=4000) or {}
            largura = float(caixa.get("width", 0) or 0)
            altura = float(caixa.get("height", 0) or 0)

            if largura > 40 and altura > 4:
                posicao = {
                    "x": max(30, min(largura - 8, largura * 0.25)),
                    "y": max(2, min(altura - 2, altura * 0.50)),
                }
                linha.click(
                    timeout=TEMPO_FIORI_ELEMENTO,
                    position=posicao,
                    force=tentativa == 3,
                    no_wait_after=tentativa >= 2,
                )
            else:
                linha.click(
                    timeout=TEMPO_FIORI_ELEMENTO,
                    force=tentativa == 3,
                    no_wait_after=tentativa >= 2,
                )

            page_referencia.wait_for_timeout(450)
            selecionada, _estado, detalhe_estado = ler_estado_selecao()

            if not selecionada:
                ultimo_erro = detalhe_estado or "a Row não apresentou evidência de seleção ativa"
                print(
                    "[FIORI VA03] A Row não confirmou a seleção da FAT; "
                    "tentando o texto exato do mesmo documento uma vez."
                )
                if clicar_texto_exato_da_fat():
                    page_referencia.wait_for_timeout(450)
                    selecionada, _estado, detalhe_estado = ler_estado_selecao()
                    if not selecionada:
                        ultimo_erro = detalhe_estado or "texto clicado, mas a Row continuou sem estado ativo"

            if not selecionada:
                if tentativa < 3:
                    continue
                break

        except Exception as e:
            ultimo_erro = str(e)
            if tentativa < 3:
                continue
            break

        try:
            return localizar_botao_acionavel_em_frames_fiori(
                page_referencia,
                padrao_nome=re.compile(r"^\s*Exibir\s+documento\s*$", re.I),
                descricao="Botão 'Exibir documento'",
                frame_preferencial=frame_fluxo,
                timeout=5000 if tentativa < 3 else TEMPO_FIORI_ELEMENTO,
            )
        except Exception as e:
            ultimo_erro = str(e)
            if tentativa < 3:
                print(
                    "[FIORI VA03] A FAT ficou selecionada, mas o botão "
                    "'Exibir documento' ainda não foi localizado; repetindo a seleção."
                )

    raise RuntimeError(
        f"O documento '{descricao_faturamento}' foi localizado, mas a linha "
        "não ficou comprovadamente selecionada para abrir 'Exibir documento'. "
        f"Último detalhe: {ultimo_erro or 'sem evidência adicional'}"
    )


def selecionar_faturamento_nc_por_texto(
    page_referencia,
    frame_fluxo,
    locator_faturamento,
    descricao_faturamento="documento de faturamento",
):
    """Clica no texto da FAT já validada, dentro da mesma linha da OV.

    Nas rodadas de 27 e 29/10, clicar numa fração da largura da Row deixou a classe
    urST4RowActive, mas o SAP rejeitou a seleção. Essa classe indica foco e não
    comprova que o documento foi selecionado. A abertura só é confirmada pela
    tela seguinte em baixar_nota_existente_va03; aqui apenas preparamos o clique.
    A seleção usada pela atividade Devolução e Refaturamento não é alterada.
    """
    normalizar = lambda texto: re.sub(r"\s+", " ", str(texto or "")).strip()
    esperado = normalizar(descricao_faturamento)
    if not re.match(r"^(?:Fat\b|Fatura\b)", esperado, re.I):
        raise RuntimeError("Descrição exata da fatura NC não disponível para seleção.")
    linha = locator_faturamento
    if "-Row-" not in (linha.get_attribute("id") or ""):
        ancestral = linha.locator("xpath=ancestor-or-self::*[contains(@id,'-Row-')][1]")
        if ancestral.count() == 1:
            linha = ancestral.first

    ultimo_erro = "Texto da fatura indisponível."
    for tentativa in (1, 2):
        clique_iniciado = False
        try:
            if normalizar(linha.inner_text(timeout=2000)) != esperado:
                raise RuntimeError("A linha da fatura mudou; não clicar em outro documento.")
            # Não usar coordenadas na Row nem o ícone de expandir/recolher.
            # O número da fatura e o item fazem parte do texto exato.
            buscas = [linha.get_by_text(esperado, exact=True)]
            # No Codegen enviado, o botão tem só o nome do tipo de fatura;
            # número e item podem ser nós separados. O botão continua limitado
            # à linha cujo texto COMPLETO, incluindo número/item, foi validado.
            tipo = re.sub(r"\s+\d{8,10}(?:\s*/\s*\d+)?$", "", esperado).strip()
            if tipo != esperado:
                buscas.append(linha.get_by_role("button", name=tipo, exact=True))
            id_linha = linha.get_attribute("id") or ""
            alvo = None
            for candidatos in buscas:
                visiveis = [candidatos.nth(i) for i in range(candidatos.count())
                            if candidatos.nth(i).is_visible()]
                if len(visiveis) > 1:
                    raise RuntimeError("Mais de um alvo exato na linha da fatura; seleção ambígua.")
                if visiveis and (not id_linha or visiveis[0].get_attribute("id") != id_linha):
                    alvo = visiveis[0]
                    break
            if alvo is None:
                raise RuntimeError("Texto/botão da fatura não isolado dentro da linha validada.")
            alvo.scroll_into_view_if_needed(timeout=5000)
            if normalizar(linha.inner_text(timeout=2000)) != esperado:
                raise RuntimeError("A linha da fatura mudou após a rolagem.")
            clique_iniciado = True
            alvo.click(timeout=TEMPO_FIORI_ELEMENTO)
            print(f"[FIORI VA03 NC] Texto exato clicado: {esperado}. "
                  "A abertura será confirmada na próxima tela do SAP.")
            # Não repetir o clique por ausência de classe CSS. O destino da
            # transição, após Exibir documento, é a confirmação efetiva.
            return localizar_botao_acionavel_em_frames_fiori(
                page_referencia,
                padrao_nome=re.compile(r"^\s*Exibir\s+documento\s*$", re.I),
                descricao="Botão 'Exibir documento'",
                frame_preferencial=frame_fluxo,
                timeout=5000,
            )
        except Exception as erro:
            ultimo_erro = str(erro)
            if clique_iniciado:
                # Mesmo uma seleção pode alternar o estado da árvore. Um
                # timeout após o clique não autoriza clicar de novo às cegas.
                break
            if tentativa < 2:
                page_referencia.wait_for_timeout(350)
    raise RuntimeError(f"Não foi possível selecionar o texto exato de '{esperado}': {ultimo_erro}")


def selecionar_faturamento_por_atividade(
    page_referencia, frame_fluxo, locator_faturamento, descricao_faturamento, atividade,
):
    if classificar_atividade_suporte_venda(atividade) == ATIVIDADE_NOTA_CREDITO:
        selecionar = selecionar_faturamento_nc_por_texto
    else:
        selecionar = selecionar_faturamento_e_localizar_exibir_documento
    return selecionar(page_referencia, frame_fluxo, locator_faturamento, descricao_faturamento)


def aguardar_resumo_ordem_va03_carregado(
    pagina_app,
    frame_va03,
    campo_ordem,
    ordem,
    timeout=TEMPO_FIORI_NAVEGACAO,
):
    """Aguarda a tela de síntese da OV terminar antes de clicar no fluxo."""
    limite = time.time() + (timeout / 1000)
    padrao_ordem = re.compile(rf"\b{re.escape(str(ordem).strip())}\b")

    while time.time() < limite:
        try:
            campo_ainda_visivel = campo_ordem.is_visible()
        except Exception:
            campo_ainda_visivel = False

        botao_fluxo = _primeiro_item_visivel([
            frame_va03.get_by_role(
                "button",
                name=PADRAO_BOTAO_FLUXO_DOCUMENTOS,
            )
        ])

        titulo_com_ordem = _primeiro_item_visivel([
            frame_va03.get_by_text(padrao_ordem),
            pagina_app.get_by_text(padrao_ordem),
        ], limite_itens=10)

        # Na tela inicial da VA03, o botão já existe. A condição adicional
        # evita clicar nele enquanto o Enter da OV ainda está sendo processado.
        if botao_fluxo is not None and (
            not campo_ainda_visivel or titulo_com_ordem is not None
        ):
            # O SAP já expõe o botão antes de terminar de ligar todos os
            # eventos da tela. Esta pequena estabilização evita um clique que
            # parece correto para o Playwright, mas é ignorado pelo WebGUI.
            pagina_app.wait_for_timeout(1800)
            print(f"[FIORI VA03] Síntese da OV {ordem} carregada e confirmada.")
            return

        try:
            pagina_app.wait_for_timeout(250)
        except Exception:
            time.sleep(0.25)

    raise RuntimeError(
        f"A tela de síntese da OV {ordem} não terminou de carregar na VA03."
    )


def abrir_fluxo_documentos_fiori(page_referencia, frame_resumo):
    """Aciona o fluxo e só prossegue depois de confirmar a troca de tela."""
    ultimo_erro = ""

    # O SAP WebGUI pode exibir o botão antes de o handler JavaScript estar
    # pronto. Reobtemos o locator em cada tentativa e alternamos o mecanismo
    # de acionamento. A tela seguinte precisa ser confirmada; não aceitamos
    # apenas o retorno bem-sucedido de click().
    estrategias = [
        ("clique normal", 10000),
        ("tecla Enter", 10000),
        ("clique forçado", 15000),
        ("evento DOM", TEMPO_FIORI_NAVEGACAO),
    ]

    for tentativa, (metodo, tempo_confirmacao) in enumerate(estrategias, start=1):
        # Antes de repetir, verifica se a troca terminou logo após o tempo da
        # tentativa anterior. Assim não procuramos novamente um botão que já
        # saiu da tela.
        if tentativa > 1:
            try:
                return localizar_contexto_fluxo_documentos_fiori(
                    page_referencia,
                    frame_resumo=frame_resumo,
                    timeout=1500,
                )
            except Exception:
                pass

        try:
            pagina_botao, frame_botao, botao_fluxo = (
                localizar_botao_fluxo_documentos_fiori(
                    page_referencia,
                    frame_preferencial=frame_resumo if tentativa == 1 else None,
                    timeout=15000,
                )
            )
        except Exception as e:
            ultimo_erro = str(e)
            # O botão pode ter desaparecido justamente porque a navegação
            # concluiu durante a busca. Confere a tela uma última vez.
            try:
                return localizar_contexto_fluxo_documentos_fiori(
                    page_referencia,
                    frame_resumo=frame_resumo,
                    timeout=5000,
                )
            except Exception:
                if tentativa < len(estrategias):
                    continue
                break

        pagina_botao.bring_to_front()
        botao_fluxo.scroll_into_view_if_needed(timeout=5000)
        print(
            "[FIORI VA03] Acionando 'Exibir fluxo de documentos' por "
            f"{metodo} (tentativa {tentativa}/{len(estrategias)})..."
        )

        try:
            if metodo == "clique normal":
                botao_fluxo.click(timeout=TEMPO_FIORI_ELEMENTO)
            elif metodo == "tecla Enter":
                botao_fluxo.focus(timeout=5000)
                botao_fluxo.press("Enter", timeout=TEMPO_FIORI_ELEMENTO)
            elif metodo == "clique forçado":
                botao_fluxo.click(
                    timeout=TEMPO_FIORI_ELEMENTO,
                    force=True,
                    no_wait_after=True,
                )
            else:
                botao_fluxo.dispatch_event("click", timeout=TEMPO_FIORI_ELEMENTO)
        except Exception as e:
            ultimo_erro = f"Falha no método {metodo}: {e}"
            if tentativa < len(estrategias):
                print(
                    f"[FIORI VA03] {ultimo_erro}. Tentando o próximo método."
                )
                continue
            break

        try:
            return localizar_contexto_fluxo_documentos_fiori(
                pagina_botao,
                frame_resumo=frame_botao,
                timeout=tempo_confirmacao,
            )
        except Exception as e:
            ultimo_erro = str(e)
            if tentativa < len(estrategias):
                print(
                    "[FIORI VA03] O WebGUI ainda permaneceu na síntese; "
                    "tentando outro método de acionamento."
                )

    raise RuntimeError(
        "Não foi possível abrir/localizar o Fluxo de documentos após "
        f"{len(estrategias)} métodos de acionamento. {ultimo_erro}"
    )


def texto_completo_fluxo_fiori(frame_fiori):
    """Obtém o texto visível do fluxo após a coluna Status carregar."""
    cabecalho_status = frame_fiori.locator("div").filter(
        has_text=re.compile(r"^\s*Status\s*$", re.I)
    ).first

    try:
        cabecalho_status.wait_for(state="visible", timeout=TEMPO_FIORI_ELEMENTO)
    except Exception:
        # Fallback caso o cabeçalho não seja um DIV nesta versão do WebGUI.
        frame_fiori.get_by_text(
            re.compile(r"^\s*Status\s*$", re.I),
            exact=True,
        ).first.wait_for(state="visible", timeout=TEMPO_FIORI_ELEMENTO)

    # O cabeçalho pode aparecer ligeiramente antes das linhas. Aguarda o texto
    # estabilizar em duas leituras para não decidir com a árvore ainda vazia.
    corpo = frame_fiori.locator("body")
    texto_anterior = ""
    leituras_estaveis = 0
    limite = time.time() + 10

    while time.time() < limite:
        texto_atual = str(corpo.inner_text(timeout=TEMPO_FIORI_ELEMENTO) or "")
        if texto_atual and texto_atual == texto_anterior:
            leituras_estaveis += 1
            if leituras_estaveis >= 2:
                return texto_atual
        else:
            leituras_estaveis = 0
            texto_anterior = texto_atual
        time.sleep(0.4)

    return texto_anterior


def _linhas_documento_dom_fluxo(frame_fiori):
    """Retorna as linhas do componente SAP tree, sem depender do sufixo C110."""
    return frame_fiori.locator("[id*='-Row-']")


def localizar_botao_download_em_frames(
    page_referencia,
    timeout=TEMPO_FIORI_DOWNLOAD,
):
    """
    Localiza o Download pela hierarquia real gravada no SAP WebGUI.

    Estrutura observada:
    application-SalesOrder-change-iframe (ou SalesDocument-display)
      -> itshtmlvwrfn...
        -> iframe de identificador dinâmico
          -> botão Download

    O identificador do último iframe muda a cada abertura e, por isso, nunca
    é fixado no código. A busca genérica fica somente como contingência.
    """
    limite = time.time() + (timeout / 1000)
    ultimo_mapa_frames = []

    while time.time() < limite:
        mapa_frames = _mapear_frames_visualizador(page_referencia)

        for pagina in _paginas_fiori_priorizadas(page_referencia):
            try:
                frames_intermediarios = [
                    frame
                    for frame in pagina.frames
                    if re.match(r"^itshtmlvwrfn", str(frame.name or ""), re.I)
                ]
            except Exception:
                continue

            # Prioriza o caminho exato fornecido pelo Codegen, mas também
            # aceita SalesDocument-display e pequenas variações do app.
            frames_intermediarios.sort(
                key=lambda frame: _prioridade_hierarquia_visualizador(frame)
            )

            for frame_intermediario in frames_intermediarios:
                try:
                    descendentes = list(_frames_descendentes(frame_intermediario))
                except Exception:
                    descendentes = []

                # O botão está no iframe dinâmico descendente. O próprio
                # itshtml é incluído no fim apenas para tolerar outra versão.
                contextos = list(reversed(descendentes)) + [frame_intermediario]

                for frame_download in contextos:
                    botao = _localizador_download_no_frame(
                        frame_download,
                        timeout_item=600,
                    )
                    if botao is None:
                        continue

                    pagina.bring_to_front()
                    print(
                        "[FIORI PDF] Hierarquia do visualizador localizada: "
                        f"{_cadeia_nomes_frame(frame_download)}."
                    )
                    return pagina, frame_download, botao

        # Contingência: se o fornecedor mudar o nome itshtmlvwrfn, procura o
        # botão em todos os frames, sem chamar count() (que foi o ponto em que
        # a execução anterior ficou presa).
        for pagina in _paginas_fiori_priorizadas(page_referencia):
            try:
                frames = list(reversed(pagina.frames))
            except Exception:
                continue

            for frame in frames:
                botao = _localizador_download_no_frame(frame, timeout_item=120)
                if botao is None:
                    continue

                pagina.bring_to_front()
                print(
                    "[FIORI PDF] Download localizado pela contingência no frame: "
                    f"{_cadeia_nomes_frame(frame)}."
                )
                return pagina, frame, botao

        ultimo_mapa_frames = mapa_frames
        try:
            page_referencia.wait_for_timeout(300)
        except Exception:
            time.sleep(0.3)

    detalhe = " | ".join(ultimo_mapa_frames[-12:]) or "nenhum frame disponível"
    raise RuntimeError(
        "Botão Download não localizado em nenhuma aba/frame do visualizador. "
        f"Frames observados: {detalhe}"
    )


def _frames_descendentes(frame_raiz):
    """Percorre recursivamente os frames abaixo de um frame do Playwright."""
    for frame_filho in frame_raiz.child_frames:
        yield frame_filho
        yield from _frames_descendentes(frame_filho)


def _cadeia_frames(frame_fiori):
    """Retorna a cadeia principal -> ... -> frame atual."""
    cadeia = []
    atual = frame_fiori

    while atual is not None:
        cadeia.append(atual)
        try:
            atual = atual.parent_frame
        except Exception:
            atual = None

    return list(reversed(cadeia))


def _cadeia_nomes_frame(frame_fiori):
    nomes = []
    for frame in _cadeia_frames(frame_fiori):
        try:
            nomes.append(frame.name or "<principal>")
        except Exception:
            nomes.append("<frame sem nome>")
    return " > ".join(nomes)


def _prioridade_hierarquia_visualizador(frame_intermediario):
    """Ordena primeiro a hierarquia exata observada no Codegen."""
    cadeia = _cadeia_nomes_frame(frame_intermediario).lower()
    if "application-salesorder-change-iframe" in cadeia:
        return 0
    if "application-salesdocument-display-iframe" in cadeia:
        return 1
    if "application-sales" in cadeia:
        return 2
    if "application-" in cadeia:
        return 3
    return 4


def _localizador_download_no_frame(frame_fiori, timeout_item=500):
    """
    Obtém o locator sem count() e sem exigir clique manual no visualizador.

    O estado attached é suficiente nesta fase; click() fará a espera de
    acionabilidade e a segunda estratégia poderá forçar o clique.
    """
    padrao_download = re.compile(r"^\s*Download\s*$", re.I)
    locators = [
        frame_fiori.get_by_role("button", name=padrao_download),
        frame_fiori.get_by_role("link", name=padrao_download),
        frame_fiori.locator(
            '[aria-label="Download"], [title="Download"], '
            'button:has-text("Download")'
        ),
    ]

    for locator in locators:
        alvo = locator.first
        try:
            alvo.wait_for(state="attached", timeout=timeout_item)
            return alvo
        except Exception:
            continue

    return None


def _mapear_frames_visualizador(page_referencia):
    """Gera diagnóstico apenas das cadeias ligadas ao Fiori/visualizador."""
    detalhes = []

    for pagina in _paginas_fiori_priorizadas(page_referencia):
        try:
            frames = pagina.frames
        except Exception:
            continue

        for frame in frames:
            cadeia = _cadeia_nomes_frame(frame)
            cadeia_normalizada = cadeia.lower()
            if "application-" not in cadeia_normalizada and "itshtmlvwrfn" not in cadeia_normalizada:
                continue
            try:
                detalhes.append(f"{cadeia}; url={frame.url[:120]}")
            except Exception:
                detalhes.append(cadeia)

    return detalhes


def baixar_pdf_visualizador_fiori(page_fiori, caminho_pdf):
    """Captura o download do visualizador e salva diretamente no destino da OV."""
    pasta_destino = os.path.dirname(caminho_pdf)
    if pasta_destino:
        os.makedirs(pasta_destino, exist_ok=True)

    pagina_download, _frame_download, botao_download = localizar_botao_download_em_frames(
        page_fiori
    )

    download = None
    ultimo_erro = ""

    estrategias = [
        "clique normal",
        "clique forçado",
        "evento DOM",
    ]

    for tentativa, metodo in enumerate(estrategias, start=1):
        try:
            print(
                f"[FIORI PDF] Acionando Download por {metodo} "
                f"(tentativa {tentativa}/{len(estrategias)})..."
            )
            with pagina_download.expect_download(
                timeout=30000,
            ) as download_info:
                if metodo == "clique normal":
                    botao_download.click(timeout=TEMPO_FIORI_ELEMENTO)
                elif metodo == "clique forçado":
                    botao_download.click(
                        timeout=TEMPO_FIORI_ELEMENTO,
                        force=True,
                        no_wait_after=True,
                    )
                else:
                    botao_download.dispatch_event(
                        "click",
                        timeout=TEMPO_FIORI_ELEMENTO,
                    )

            download = download_info.value
            print(
                "[FIORI PDF] Download iniciado pelo navegador: "
                f"{download.suggested_filename}."
            )
            break
        except Exception as e:
            ultimo_erro = str(e)
            if tentativa < len(estrategias):
                print(
                    "[FIORI PDF] O acionamento ainda não gerou o download; "
                    "tentando o próximo método."
                )
                pagina_download, _frame_download, botao_download = (
                    localizar_botao_download_em_frames(
                        pagina_download,
                        timeout=15000,
                    )
                )

    if download is None:
        raise RuntimeError(
            f"O botão Download foi localizado, mas o arquivo não iniciou. {ultimo_erro}"
        )

    print(f"[FIORI PDF] Salvando arquivo em: {caminho_pdf}")
    download.save_as(caminho_pdf)

    if not pdf_valido(caminho_pdf):
        falha = download.failure()
        detalhe = f" Falha informada pelo navegador: {falha}" if falha else ""
        raise RuntimeError(f"Arquivo baixado não possui assinatura PDF válida.{detalhe}")

    # Quando o visualizador abrir em uma nova aba, fecha somente essa aba após
    # salvar o PDF. A aba principal do Fiori permanece disponível para a próxima OV.
    if pagina_download != page_fiori:
        try:
            if not pagina_download.is_closed():
                pagina_download.close()
                print("[FIORI PDF] Aba do visualizador fechada após o download.")
        except Exception as e:
            print(f"[FIORI PDF] Aviso: não consegui fechar a aba do PDF: {e}")

    return caminho_pdf


def _extrair_documento_gravado_vf01(*textos):
    """Extrai o documento criado pela VF01 da mensagem de gravação do SAP."""
    for texto in textos:
        bruto = str(texto or "")
        if not bruto:
            continue
        for padrao in (
            r"\bdocumento\s+0*(\d{8,10})\s+foi\s+gravado\b",
            r"\bdocumento\s+0*(\d{8,10})\s+(?:gravado|salvo)\b",
            r"\b(?:faturamento|fatura|nota)\s+0*(\d{8,10})\s+(?:foi\s+)?(?:gravado|gravada|salvo|salva)\b",
        ):
            m = re.search(padrao, bruto, re.I)
            if m:
                return m.group(1)
    return ""


def _normalizar_numero_documento_sap(valor):
    """Compara 123456789 e 0123456789 como o mesmo documento SAP."""
    digitos = re.sub(r"\D", "", str(valor or ""))
    if not digitos:
        return ""
    return digitos.lstrip("0") or "0"


def _texto_contem_documento_sap(texto, documento):
    esperado = _normalizar_numero_documento_sap(documento)
    if not esperado:
        return False
    for numero in re.findall(r"(?<!\d)\d{8,12}(?!\d)", str(texto or "")):
        if _normalizar_numero_documento_sap(numero) == esperado:
            return True
    return False


def baixar_nota_sap_vf01(page_fiori, requisicao, ordem):
    """Executa a VF01 no Fiori somente para tentar faturar/gravar a OV."""
    requisicao = str(requisicao).strip()
    ordem = str(ordem).strip()
    resultado = {
        "Ordem": ordem,
        "StatusSAP": "VF01_EXECUTADA",
        "NumeroNF": "",
        "ArquivoPDF": "",
        "Mensagem": "",
        "VF01Executada": False,
        "DocumentoFaturamentoVF01": "",
    }

    pagina_app = page_fiori

    try:
        print(f"[FIORI VF01] Faturando requisição {requisicao} / OV {ordem}")
        for tentativa_abertura in range(1, 3):
            try:
                pagina_app, frame = abrir_aplicativo_fiori(
                    page_fiori,
                    termo_pesquisa="vf01",
                    nomes_aplicativo=[
                        re.compile(
                            r"3\s*[°ºo]?\s*passo\s+CRIAR\s+FATURA",
                            re.I,
                        ),
                    ],
                    tipo_aplicativo="vf01",
                )

                grid = frame.get_by_role(
                    "grid",
                    name=re.compile(
                        r"Docs\.\s*a\s+serem\s+processados",
                        re.I,
                    ),
                )
                grid.locator('input[name="InputField"]').first.wait_for(
                    state="visible",
                    timeout=5000,
                )
                break
            except Exception as erro_abertura:
                if tentativa_abertura == 1:
                    print(
                        "[FIORI VF01] A grade real não terminou de carregar; "
                        "reabrindo a VF01 uma única vez..."
                    )
                    continue
                raise RuntimeError(
                    "A VF01 não confirmou a grade e o campo da OV após "
                    f"duas aberturas. Último detalhe: {erro_abertura}"
                ) from erro_abertura

        campo_ov = limpar_e_preencher_ov_vf01(
            grid,
            ordem,
        )
        info_execucao = executar_e_gravar_vf01(frame, campo_ov)
        # Só chega aqui com gravação confirmada ou resposta ATUAL sem novo
        # faturamento. Um comando enviado sem resposta não libera a VA03.
        frame = info_execucao.get("frame", frame)
        resultado["VF01Executada"] = True
        pagina_app.wait_for_timeout(TEMPO_FIORI_APOS_SALVAR)

        mensagem = capturar_mensagem_fiori(frame)
        cenario_vf01 = str(info_execucao.get("cenario", "") or "").strip()
        metodo_gravacao = str(info_execucao.get("metodo", "") or "").strip()
        msg_execucao = str(info_execucao.get("mensagem", "") or "").strip()
        detalhe_salvar = str(info_execucao.get("detalhe_salvar", "") or "").strip()

        if cenario_vf01 == "FATURAMENTO_GRAVADO_ERRO_CONTA":
            resultado["StatusSAP"] = "VF01_GRAVADA_ERRO_DETERMINACAO_CONTA"
        elif cenario_vf01 == "FATURAMENTO_INCOMPLETO":
            resultado["StatusSAP"] = "VF01_FATURAMENTO_INCOMPLETO"
        elif cenario_vf01 == "SEM_NOVO_FATURAMENTO":
            resultado["StatusSAP"] = "VF01_SEM_NOVO_FATURAMENTO"
        elif cenario_vf01 == "DOCUMENTO_INEXISTENTE":
            resultado["StatusSAP"] = "VF01_DOCUMENTO_INEXISTENTE"
            resultado["RejeicaoSemGravacao"] = True
        elif info_execucao.get("gravacao_confirmada"):
            resultado["StatusSAP"] = "VF01_GRAVACAO_CONFIRMADA"
        else:
            resultado["StatusSAP"] = "VF01_FATURAMENTO_TENTADO"

        partes_msg = []
        for parte in (mensagem, msg_execucao):
            if parte and parte not in partes_msg:
                partes_msg.append(parte)
        if detalhe_salvar:
            partes_msg.append(detalhe_salvar)

        resultado["Mensagem"] = " | ".join(partes_msg) or (
            f"VF01 acionada por {metodo_gravacao}. "
            "A VA03 verificará o estado definitivo da OV."
        )

        if info_execucao.get("gravacao_confirmada"):
            documento_vf01 = str(info_execucao.get("documento_faturamento", "") or "").strip()
            if not documento_vf01:
                documento_vf01 = _extrair_documento_gravado_vf01(
                    msg_execucao, mensagem, resultado["Mensagem"]
                )
            if documento_vf01:
                resultado["DocumentoFaturamentoVF01"] = documento_vf01
                print(
                    f"[FIORI VF01] OV {ordem}: documento recém-gravado capturado "
                    f"para conferência exata na VA03: {documento_vf01} "
                    f"(visual esperado também como {documento_vf01.zfill(10)})."
                )
            else:
                print(
                    f"[FIORI VF01] OV {ordem}: gravação confirmada, porém o número "
                    "do documento não pôde ser extraído; VA03 usará a validação "
                    "estrutural da ramificação."
                )

        print(
            f"[FIORI VF01] OV {ordem}: tentativa concluída "
            f"({resultado['StatusSAP']}). {resultado['Mensagem']}"
        )
        return resultado

    except Exception as e:
        evidencia = salvar_evidencia_fiori(
            pagina_app,
            prefixo="erro_vf01_fiori",
            requisicao=requisicao,
            ordem=ordem,
        )
        resultado["StatusSAP"] = "ERRO_VF01"
        if FIORI_PENDENCIA_GRAVACAO:
            FIORI_PENDENCIA_GRAVACAO.update({"requisicao": requisicao, "ordem": ordem})
            resultado["StatusSAP"] = "VF01_RESULTADO_NAO_CONFIRMADO"
        resultado["Mensagem"] = f"ERRO VF01/Fiori: {e}"
        if evidencia:
            resultado["Mensagem"] += f" Evidência: {os.path.basename(evidencia)}"
        print(f"[FIORI VF01] OV {ordem}: {resultado['Mensagem']}")
        return resultado



def coletar_texto_fluxo_generico(frame_fiori):
    """Lê o corpo e as linhas DOM do fluxo sem depender do tipo de faturamento."""
    partes = []
    vistos = set()

    def adicionar(valor):
        valor = str(valor or "").strip()
        if valor and valor not in vistos:
            vistos.add(valor)
            partes.append(valor)

    try:
        adicionar(texto_completo_fluxo_fiori(frame_fiori))
    except Exception:
        pass

    linhas = _linhas_documento_dom_fluxo(frame_fiori)
    try:
        quantidade = min(linhas.count(), 250)
    except Exception:
        quantidade = 0

    for indice in range(quantidade):
        try:
            adicionar(linhas.nth(indice).inner_text(timeout=1500))
        except Exception:
            continue

    return "\n".join(partes)


def diagnosticar_linhas_fluxo_va03(frame_fiori, ordem="", limite=60):
    """
    Imprime linhas úteis do Fluxo de documentos quando o faturamento esperado
    não é localizado. Serve para descobrir variações reais de rótulo no SAP
    sem interromper a fila por falta de evidência.
    """
    ordem = str(ordem or "").strip()
    termos = (
        "fat",
        "nota",
        "cred",
        "loca",
        "documento",
        "br-",
        "pendente",
        "concluido",
        "concluído",
    )

    candidatos = []
    vistos = set()

    linhas = _linhas_documento_dom_fluxo(frame_fiori)
    try:
        quantidade = min(linhas.count(), 300)
    except Exception:
        quantidade = 0

    for indice in range(quantidade):
        try:
            texto_linha = (linhas.nth(indice).inner_text(timeout=1200) or "").strip()
        except Exception:
            continue

        if not texto_linha:
            continue

        normalizado = normalizar_texto_busca(texto_linha)
        contem_ordem = False
        if ordem:
            try:
                contem_ordem = bool(
                    re.search(rf"(?<!\d)0*{re.escape(ordem)}(?:\s*/|\b)", texto_linha)
                )
            except Exception:
                contem_ordem = ordem in texto_linha

        if not contem_ordem and not any(t in normalizado for t in termos):
            continue

        chave = re.sub(r"\s+", " ", texto_linha)
        if chave in vistos:
            continue
        vistos.add(chave)
        candidatos.append(chave)

        if len(candidatos) >= limite:
            break

    if candidatos:
        print(
            f"[FIORI VA03][DIAGNÓSTICO] {len(candidatos)} linha(s) relevante(s) "
            "observada(s) no Fluxo de documentos:"
        )
        for linha in candidatos:
            print(f"[FIORI VA03][DIAGNÓSTICO] {linha}")
    else:
        print(
            "[FIORI VA03][DIAGNÓSTICO] Nenhuma linha textual relevante pôde ser "
            "extraída do tree do Fluxo de documentos."
        )

    return candidatos


def obter_config_atividade_suporte_venda(atividade):
    atividade = str(atividade or "").strip()
    if atividade not in CONFIG_ATIVIDADES_SUPORTE_VENDA:
        raise RuntimeError(
            f"Atividade não suportada para o fluxo SAP: {atividade or '<vazia>'}."
        )
    return CONFIG_ATIVIDADES_SUPORTE_VENDA[atividade]


def _extrair_numero_documento_suporte_venda(locator, atividade):
    """Extrai o número de 10 dígitos da linha de faturamento localizada.

    TESTE 44: o número é estrutural e não depende mais do nome da FAT.
    O SAP já mostrou, entre outros: Fat Nota de crédito, Fat NC Ven Serv,
    Fat NC Loca Mensal e Fatura - NC Locação (exemplo).
    """
    textos = []
    for atributo in ("aria-label", "title"):
        try:
            valor = locator.get_attribute(atributo)
            if valor:
                textos.append(valor)
        except Exception:
            pass
    try:
        valor = locator.inner_text(timeout=1800)
        if valor:
            textos.append(valor)
    except Exception:
        pass

    padrao = re.compile(r"(?<!\d)(\d{10})\s*/", re.I)
    for texto in textos:
        encontrados = padrao.findall(str(texto or ""))
        if encontrados:
            return encontrados[-1]
    return ""


def _padroes_ramificacao_br_por_atividade(atividade):
    """Reconhece a ramificação atual pela estrutura BR + OV exata.

    TESTE 44: o tipo textual da BR deixou de ser critério primário. Em produção
    a mesma atividade PORTAL já abriu NC Locação (exemplo), BR-Nota de Credito.,
    NC de serviços e outras variantes. A trava real é a MESMA linha BR
    conter a OV atual; ramificações históricas com outras OVs continuam fora.
    """
    return [re.compile(r"^\s*BR-", re.I)]

def _linha_contem_ov_atual(texto_linha, ordem):
    """Confirma a OV atual aceitando zeros à esquerda no fluxo SAP."""
    ordem = str(ordem or "").strip()
    if not ordem:
        return False
    try:
        return bool(
            re.search(
                rf"(?<!\d)0*{re.escape(ordem)}(?:\s*/|\b)",
                str(texto_linha or ""),
                re.I,
            )
        )
    except Exception:
        return ordem in str(texto_linha or "")


def _ramificacao_atual_fluxo(frame_fiori, atividade, ordem):
    """
    Localiza a linha BR da OV atual e devolve somente os nós abaixo dela.

    TESTE 36:
    nunca mais usa um faturamento histórico que esteja ACIMA da NC/Nota
    da OV atual. A ramificação termina quando aparece a próxima linha iniciada
    por "BR-".

    Retorno:
      {
        "pai": Locator|None,
        "texto_pai": str,
        "indice_pai": int,
        "filhos": [(indice, locator, texto), ...],
        "linhas": [(indice, locator, texto), ...],
      }
    """
    padroes_br = _padroes_ramificacao_br_por_atividade(atividade)
    linhas = _linhas_documento_dom_fluxo(frame_fiori)

    try:
        quantidade = min(linhas.count(), 300)
    except Exception:
        quantidade = 0

    registros = []
    for indice in range(quantidade):
        linha = linhas.nth(indice)
        try:
            texto_linha = (linha.inner_text(timeout=1600) or "").strip()
        except Exception:
            texto_linha = ""
        if texto_linha:
            registros.append((indice, linha, texto_linha))

    indice_registro_pai = None
    for posicao, (indice_dom, linha, texto_linha) in enumerate(registros):
        if not any(p.search(texto_linha) for p in padroes_br):
            continue
        if not _linha_contem_ov_atual(texto_linha, ordem):
            continue
        # TESTE 44: se a mesma OV possuir mais de uma BR histórica, a operação
        # mais recente tende a ser a última ramificação exibida no fluxo.
        # Mantemos a última ocorrência, em vez de parar na primeira.
        indice_registro_pai = posicao

    if indice_registro_pai is None:
        return {
            "pai": None,
            "texto_pai": "",
            "indice_pai": -1,
            "filhos": [],
            "linhas": registros,
        }

    indice_dom, linha_pai, texto_pai = registros[indice_registro_pai]
    painel_pai = (linha_pai.get_attribute("id") or "").split("-Row-", 1)[0]
    filhos = []

    for posicao in range(indice_registro_pai + 1, len(registros)):
        idx, linha, texto_linha = registros[posicao]
        # O SAP lista o painel direito depois da árvore inteira no DOM.
        # Seus status históricos não são filhos da última OV da árvore.
        painel_linha = (linha.get_attribute("id") or "").split("-Row-", 1)[0]
        if painel_linha != painel_pai:
            continue
        normalizado = normalizar_texto_busca(texto_linha)

        # Nova linha de negócio = fim da ramificação atual.
        if normalizado.startswith("br-"):
            break

        filhos.append((idx, linha, texto_linha))

    return {
        "pai": linha_pai,
        "texto_pai": texto_pai,
        "indice_pai": indice_dom,
        "filhos": filhos,
        "linhas": registros,
    }


def _status_linha_fluxo_fiori(linha):
    """Lê o status da mesma linha, inclusive no painel direito do tree SAP.

    Não associa por ordem global do DOM: a rodada 36 lista primeiro todos os
    documentos, depois as colunas de dados. Exige alinhamento visual único.
    Se houver estados conflitantes na mesma faixa, não faz inferência.
    """
    try:
        linha.scroll_into_view_if_needed(timeout=3000)
        dados = linha.evaluate("""el => {
            const visivel = e => {
                const r = e.getBoundingClientRect(), cs = getComputedStyle(e);
                return r.width > 0 && r.height > 0 && cs.visibility !== 'hidden' &&
                    cs.display !== 'none' && !e.closest('[aria-hidden="true"]') &&
                    r.bottom > 0 && r.top < innerHeight;
            };
            const base = el.getBoundingClientRect();
            if (!visivel(el)) return {textos:[], origem:'linha não visível'};
            const textos = [{id:el.id || '', texto:el.innerText || ''}];
            for (const e of document.querySelectorAll('[id*="-Row-"], [role="row"], tr')) {
                if (e === el || !visivel(e)) continue;
                const r = e.getBoundingClientRect();
                // Mesma faixa vertical, mesma altura e painel à direita, ou
                // uma linha semântica única que contém o próprio documento.
                const mesmaFaixa = Math.abs((r.top+r.bottom-base.top-base.bottom)/2) <= 2 &&
                    Math.abs(r.height-base.height) <= 4;
                if (!mesmaFaixa || (!e.contains(el) && r.left < base.right - 3)) continue;
                textos.push({id:e.id || '', texto:e.innerText || ''});
            }
            return {textos, origem:'mesma linha visual', id:el.id || ''};
        }""")
        # Não depende de String.match/matchAll/normalize do ambiente SAP.
        # A rodada 37 falhou aqui mesmo com a nota visível no fluxo.
        estados = set()
        evidencias = []
        for item in dados.get("textos", []):
            texto = str(item.get("texto", "") or "")
            encontrados = re.findall(
                r"\b(pendente|concluido|faturamento completo)\b",
                normalizar_texto_busca(texto),
            )
            if encontrados:
                estados.update(encontrados)
                evidencias.append({"id": item.get("id", ""),
                    "texto": re.sub(r"\s+", " ", texto).strip()[:250]})
        return {"status": next(iter(estados)) if len(estados) == 1 else "",
                "textos": [str(t.get("texto", "") or "") for t in dados.get("textos", [])],
                "origem": "ambíguo" if len(estados) > 1 else dados.get("origem", ""),
                "id": dados.get("id", ""), "evidencias": evidencias}
    except Exception as erro:
        return {"status": "", "origem": f"status não confirmado: {erro}"}


def _status_br_nota_credito_pendente(frame_fiori, atividade, ordem):
    """
    Verifica APENAS a ramificação BR correspondente à OV atual.

    Regra TESTE 41:
    - verifica somente o status da linha BR correspondente à OV atual;
    - se estiver Pendente, registra a pendência e não repete a VF01 automaticamente;
    - faturamentos históricos acima dessa BR nunca podem ser usados.
    """
    ramo = _ramificacao_atual_fluxo(frame_fiori, atividade, ordem)
    texto_pai = str(ramo.get("texto_pai", "") or "").strip()

    if not texto_pai:
        print(
            f"[FIORI VA03] Ramificação BR da OV atual {ordem} não foi "
            "localizada no Fluxo de documentos."
        )
        return False, ""

    print(
        f"[FIORI VA03] Ramificação atual confirmada para a OV {ordem}: "
        f"{texto_pai}"
    )

    info_status = _status_linha_fluxo_fiori(ramo["pai"])
    try:
        if not _linha_contem_ov_atual(ramo["pai"].inner_text(timeout=1200), ordem):
            return None, "A linha mudou durante a leitura do status; associação não confirmada."
    except Exception:
        return None, "A linha da OV não pôde ser reconfirmada após ler o status."
    status_atual = info_status.get("status", "")
    print(
        f"[FIORI VA03] Status associado à OV {ordem}: "
        f"{status_atual or 'NÃO CONFIRMADO'}; {info_status.get('origem', '')}; "
        f"linha={info_status.get('id', '')}; evidência={info_status.get('evidencias', [])}"
    )
    if status_atual == "pendente":
        print(
            "[FIORI VA03] Ramificação da OV atual ainda está Pendente: "
            f"{texto_pai}"
        )
        return True, f"{texto_pai} | Status: Pendente (mesma linha)"

    if not status_atual:
        return None, f"{texto_pai} | Status não associado com segurança à linha."
    return False, texto_pai


def _locator_faturamento_na_ramificacao_atual(
    frame_fiori,
    atividade,
    ordem,
    documento_faturamento_vf01="",
):
    """Localiza a FAT correta somente dentro da ramificação da OV atual.

    TESTE 44 - regra estrutural:
    - primeiro localiza a BR que contém exatamente a OV da iteração;
    - considera somente filhos dessa BR;
    - a FAT precisa ser uma linha Fat/Fatura com Doc.contábil imediatamente abaixo;
    - se a VF01 devolveu um documento, exige o MESMO número na FAT;
    - sem número novo, usa a última FAT estruturalmente válida da ramificação.

    O nome da atividade PORTAL deixa de restringir a nomenclatura da FAT, porque
    foram observadas variantes reais como Fat Nota de crédito, Fat NC Ven Serv
    e Fatura - NC Locação (exemplo) mesmo em requisições classificadas como
    Nota de Credito.
    """
    ramo = _ramificacao_atual_fluxo(frame_fiori, atividade, ordem)
    texto_pai = str(ramo.get("texto_pai", "") or "").strip()
    if not texto_pai:
        return None

    padrao_doc_contabil = re.compile(r"^\s*Doc\.?\s*cont[aá]bil\b", re.I)
    padrao_faturamento = re.compile(r"^\s*(?:Fat\b|Fatura\b)", re.I)

    filhos = list(ramo.get("filhos") or [])
    candidatos = []
    documento_esperado = _normalizar_numero_documento_sap(documento_faturamento_vf01)

    if documento_esperado:
        print(
            f"[FIORI VA03] Documento alvo recebido da VF01: "
            f"{documento_faturamento_vf01} / {str(documento_faturamento_vf01).zfill(10)}. "
            "Será exigido esse número exato dentro da ramificação da OV atual."
        )
    else:
        print(
            f"[FIORI VA03] VF01 não informou número novo para a OV {ordem}. "
            "Será usada a última linha Fat/Fatura da ramificação atual que tenha "
            "Doc.contábil imediatamente abaixo."
        )

    for posicao, (_, linha, texto_linha) in enumerate(filhos):
        texto_linha = str(texto_linha or "").strip()
        if not padrao_faturamento.search(texto_linha):
            continue

        proximo_texto = filhos[posicao + 1][2] if posicao + 1 < len(filhos) else ""
        if not padrao_doc_contabil.search(str(proximo_texto or "")):
            continue

        corresponde_documento_vf01 = _texto_contem_documento_sap(
            texto_linha, documento_faturamento_vf01
        )
        if documento_esperado and not corresponde_documento_vf01:
            continue

        candidatos.append(
            {
                "posicao": posicao,
                "linha": linha,
                "texto": texto_linha,
                "doc_contabil": str(proximo_texto or "").strip(),
                "exato": corresponde_documento_vf01,
            }
        )

    if not candidatos:
        if documento_esperado:
            print(
                f"[FIORI VA03] Nenhuma FAT/Fatura da ramificação atual correspondeu "
                f"ao documento VF01 {documento_faturamento_vf01} com Doc.contábil "
                "imediatamente abaixo. Não será escolhida outra FAT por aproximação."
            )
        else:
            print(
                f"[FIORI VA03] Nenhuma linha Fat/Fatura com Doc.contábil logo abaixo "
                f"foi encontrada na ramificação atual da OV {ordem}."
            )
        return None

    candidatos.sort(key=lambda item: item["posicao"], reverse=True)
    escolhido = candidatos[0]
    linha = escolhido["linha"]

    try:
        linha.scroll_into_view_if_needed(timeout=7000)
        frame_fiori.page.wait_for_timeout(250)
    except Exception:
        pass

    criterio = (
        f"documento exato VF01={documento_faturamento_vf01}"
        if documento_esperado
        else "última FAT/Fatura da ramificação + Doc.contábil"
    )
    print(
        "[FIORI VA03] Faturamento confirmado estruturalmente na ramificação atual: "
        f"'{escolhido['texto']}' -> '{escolhido['doc_contabil']}' "
        f"(critério={criterio}; posição={escolhido['posicao']})."
    )
    return linha

def localizar_documento_suporte_venda_no_fluxo(frame_fiori, atividade, ordem="", documento_faturamento_vf01=""):
    """
    Valida o status da Nota de Crédito e localiza a linha de faturamento
    específica da atividade atual.
    """
    config = obter_config_atividade_suporte_venda(atividade)

    # O Codegen mostrou um clique no cabeçalho/ramificação "Documento".
    # Ele é tentado somente quando visível; se a árvore já estiver aberta,
    # a ausência desse elemento não bloqueia o processo.
    try:
        documento = frame_fiori.locator("div").filter(
            has_text=re.compile(r"^\s*Documento\s*$", re.I)
        ).first
        if documento.is_visible(timeout=1500):
            documento.click(timeout=5000)
            frame_fiori.page.wait_for_timeout(500)
    except Exception:
        pass

    pendente, detalhe_pendente = _status_br_nota_credito_pendente(
        frame_fiori,
        atividade,
        ordem,
    )
    if pendente is None:
        return {
            "ok": False, "status": "STATUS_OV_NAO_CONFIRMADO",
            "mensagem": "Status da OV atual não confirmado; não será usado status de outra linha.",
            "numero_nf": "", "locator": None, "detalhe": detalhe_pendente,
        }
    if pendente:
        return {
            "ok": False,
            "status": STATUS_NOTA_CREDITO_PENDENTE,
            "mensagem": (
                "A ramificação da Nota de Crédito permanece com status Pendente "
                "após a execução obrigatória da VF01."
            ),
            "numero_nf": "",
            "locator": None,
            "detalhe": detalhe_pendente,
        }

    print(
        "[FIORI VA03] Critério estrutural ativo: BR da OV atual -> Fat/Fatura da "
        "mesma ramificação -> Doc.contábil. O nome textual da FAT é secundário."
    )

    # TESTE 36:
    # O faturamento válido precisa estar DENTRO da ramificação da OV atual.
    # Não existe mais fallback global para "qualquer Fat Locação/Nota" da árvore.
    locator = _locator_faturamento_na_ramificacao_atual(
        frame_fiori,
        atividade,
        ordem,
        documento_faturamento_vf01=documento_faturamento_vf01,
    )

    if locator is None:
        try:
            frame_fiori.page.wait_for_timeout(1200)
        except Exception:
            pass

        diagnosticar_linhas_fluxo_va03(
            frame_fiori,
            ordem=ordem,
        )

        # Uma única reavaliação após estabilização do tree.
        locator = _locator_faturamento_na_ramificacao_atual(
            frame_fiori,
            atividade,
            ordem,
            documento_faturamento_vf01=documento_faturamento_vf01,
        )

    if locator is None:
        return {
            "ok": False,
            "status": "FATURAMENTO_NAO_LOCALIZADO",
            "mensagem": (
                f"Nenhuma FAT/Fatura estruturalmente válida foi localizada dentro "
                f"da ramificação atual da OV {ordem}. Era necessário haver uma "
                "linha Fat/Fatura ligada ao Doc.contábil; quando a VF01 informou "
                "documento novo, o número exato também era obrigatório. "
                "Faturamentos fora da ramificação foram ignorados por segurança. "
                "Consulte as linhas [DIAGNÓSTICO] do log."
            ),
            "numero_nf": "",
            "locator": None,
        }

    numero_nf = _extrair_numero_documento_suporte_venda(locator, atividade)

    # Para o log e para a etapa de seleção, prefere o texto real da linha
    # encontrada no SAP. Isso deixa claro qual dos padrões foi usado.
    descricao_real = ""
    try:
        descricao_real = (locator.inner_text(timeout=1800) or "").strip()
    except Exception:
        pass
    if not descricao_real:
        for atributo in ("aria-label", "title"):
            try:
                valor = locator.get_attribute(atributo)
                if valor:
                    descricao_real = str(valor).strip()
                    break
            except Exception:
                pass
    descricao_real = descricao_real or config["descricao_documento"]

    print(f"[FIORI VA03] Documento correto localizado: {descricao_real}.")

    return {
        "ok": True,
        "status": "FATURAMENTO_LOCALIZADO",
        "mensagem": f"Documento localizado: {descricao_real}.",
        "numero_nf": numero_nf,
        "locator": locator,
        "descricao_documento": descricao_real,
    }

def baixar_nota_existente_va03(page_fiori, requisicao, ordem, atividade, documento_faturamento_vf01=""):
    """Localiza o documento correto na VA03 e baixa o PDF da OV atual."""
    requisicao = str(requisicao).strip()
    ordem = str(ordem).strip()
    atividade = str(atividade or "").strip()
    caminho_pdf = montar_caminho_pdf(requisicao, ordem)
    resultado = {
        "Ordem": ordem,
        "StatusSAP": "ERRO_VA03",
        "NumeroNF": "",
        "ArquivoPDF": "",
        "Mensagem": "",
        "Atividade": atividade,
    }

    pagina_app = page_fiori

    try:
        config = obter_config_atividade_suporte_venda(atividade)
        print(
            f"[FIORI VA03] Requisição {requisicao} / OV {ordem} / "
            f"atividade '{atividade}'. Critério: ramificação BR da OV atual + "
            "Fat/Fatura ligada ao Doc.contábil; número VF01 exato quando disponível."
        )

        pagina_app, frame = abrir_aplicativo_fiori(
            page_fiori,
            termo_pesquisa="va03",
            nomes_aplicativo=[
                re.compile(r"Exibir\s+ordens?\s+do\s+cliente(?:\s+VA03)?", re.I),
                re.compile(r"\bVA03\b", re.I),
            ],
            tipo_aplicativo="va03",
        )

        campo_ordem = preencher_campo_fiori(
            frame.get_by_role("textbox", name="Ordem", exact=True),
            ordem,
            descricao="OV na VA03",
        )
        campo_ordem.press("Enter")

        aguardar_resumo_ordem_va03_carregado(
            pagina_app,
            frame_va03=frame,
            campo_ordem=campo_ordem,
            ordem=ordem,
        )

        pagina_app, frame = abrir_fluxo_documentos_fiori(
            pagina_app,
            frame_resumo=frame,
        )

        texto_fluxo = coletar_texto_fluxo_generico(frame)

        # Trava contra documento cruzado: o fluxo precisa conter a OV da iteração.
        padrao_ov_fluxo = re.compile(rf"(?<!\d)0*{re.escape(ordem)}(?:\s*/|\b)", re.I)
        if not padrao_ov_fluxo.search(texto_fluxo or ""):
            raise RuntimeError(
                f"O Fluxo de documentos aberto não contém a OV atual {ordem}. "
                "Processamento interrompido para impedir PDF de outra OV."
            )

        # Erro contábil só decide esta OV quando estiver dentro da ramificação
        # BR da própria OV; histórico de outras ramificações é ignorado.
        ramo_atual = _ramificacao_atual_fluxo(frame, atividade, ordem)
        textos_ramo = [str(ramo_atual.get("texto_pai", "") or "")]
        if ramo_atual.get("pai") is not None:
            textos_ramo.extend(_status_linha_fluxo_fiori(ramo_atual["pai"]).get("textos", []))
        for _idx, linha, texto_linha in (ramo_atual.get("filhos") or []):
            textos_ramo.append(str(texto_linha or ""))
            textos_ramo.extend(_status_linha_fluxo_fiori(linha).get("textos", []))
        if any(texto_indica_erro_determinacao_conta(t) for t in textos_ramo):
            evidencia = salvar_evidencia_fiori(
                pagina_app,
                prefixo="erro_determ_conta_fiori",
                requisicao=requisicao,
                ordem=ordem,
            )
            resultado["StatusSAP"] = STATUS_ERRO_DETERMINACAO_CONTA
            resultado["Mensagem"] = "Erro na determ. conta identificado dentro da ramificação da OV atual."
            if evidencia:
                resultado["Mensagem"] += f" Evidência: {os.path.basename(evidencia)}"
            print(f'[FIORI VA03] OV {ordem}: {resultado["StatusSAP"]} | {resultado["Mensagem"]}')
            return resultado

        faturamento = localizar_documento_suporte_venda_no_fluxo(
            frame,
            atividade,
            ordem=ordem,
            documento_faturamento_vf01=documento_faturamento_vf01,
        )
        if not faturamento["ok"]:
            evidencia = salvar_evidencia_fiori(
                pagina_app,
                prefixo="documento_nao_disponivel_va03",
                requisicao=requisicao,
                ordem=ordem,
            )
            resultado["StatusSAP"] = faturamento["status"]
            resultado["Mensagem"] = faturamento["mensagem"]
            if evidencia:
                resultado["Mensagem"] += f" Evidência: {os.path.basename(evidencia)}"
            print(f'[FIORI VA03] OV {ordem}: {resultado["StatusSAP"]} | {resultado["Mensagem"]}')
            return resultado

        resultado["NumeroNF"] = faturamento.get("numero_nf", "")

        # TESTE 44: seleção + abertura com uma única recuperação específica.
        # Se o SAP responder "Selecionar um documento de vendas", permanecemos
        # na VA03, readquirimos o fluxo e tentamos selecionar a MESMA FAT uma vez.
        botao_saida = None
        for tentativa_abertura in range(1, 3):
            if tentativa_abertura > 1:
                pagina_app, frame = localizar_contexto_fluxo_documentos_fiori(
                    pagina_app,
                    frame_resumo=None,
                    timeout=12000,
                )
                faturamento_retry = localizar_documento_suporte_venda_no_fluxo(
                    frame,
                    atividade,
                    ordem=ordem,
                    documento_faturamento_vf01=(documento_faturamento_vf01 or resultado['NumeroNF'])
                    if classificar_atividade_suporte_venda(atividade) == ATIVIDADE_NOTA_CREDITO
                    else documento_faturamento_vf01,
                )
                if not faturamento_retry["ok"]:
                    raise RuntimeError(
                        "A VA03 permaneceu no Fluxo de documentos, mas a FAT exata "
                        "não pôde ser readquirida para a segunda seleção. "
                        + faturamento_retry.get("mensagem", "")
                    )
                if (classificar_atividade_suporte_venda(atividade) == ATIVIDADE_NOTA_CREDITO
                        and str(faturamento_retry.get('numero_nf', '')).zfill(10)
                        != str(resultado['NumeroNF']).zfill(10)):
                    raise RuntimeError('A recuperação NC retornou outro faturamento; abertura e PDF bloqueados.')
                faturamento = faturamento_retry

            pagina_app, frame, botao_exibir = selecionar_faturamento_por_atividade(
                pagina_app,
                frame_fluxo=frame,
                locator_faturamento=faturamento["locator"],
                descricao_faturamento=faturamento["descricao_documento"],
                atividade=atividade,
            )
            botao_exibir.click(timeout=TEMPO_FIORI_ELEMENTO)

            try:
                pagina_app, frame, botao_saida = localizar_botao_apos_transicao_fiori(
                    pagina_app,
                    padrao_nome=re.compile(
                        r"^\s*Exibir\s+(?:ordem\s+de\s+)?sa[ií]da\s*$",
                        re.I,
                    ),
                    descricao="Botão 'Exibir ordem de saída'",
                    timeout=45000,
                    padrao_erro=re.compile(r"Selecionar\s+um\s+documento\s+de\s+vendas", re.I),
                    descricao_erro="Selecionar um documento de vendas",
                )
                break
            except Exception as e_transicao:
                erro_transicao = str(e_transicao)
                selecao_rejeitada = "SAP_VALIDACAO_TRANSICAO" in erro_transicao
                if selecao_rejeitada and tentativa_abertura < 2:
                    print(
                        "[FIORI VA03] O SAP pediu 'Selecionar um documento de vendas'. "
                        "A FAT correta será readquirida e selecionada novamente uma única vez, "
                        "sem retornar à VF01."
                    )
                    continue

                ainda_no_fluxo = False
                try:
                    ainda_no_fluxo = _frame_contem_fluxo_documentos(frame)
                except Exception:
                    pass
                if ainda_no_fluxo:
                    print(
                        "[FIORI VA03] O SAP permaneceu no Fluxo de documentos "
                        "depois de 'Exibir documento'. A abertura da FAT não foi confirmada."
                    )
                raise

        if botao_saida is None:
            raise RuntimeError(
                "A FAT foi localizada, porém a transição para 'Exibir ordem de saída' "
                "não foi confirmada após a recuperação limitada de seleção."
            )
        botao_saida.click(timeout=TEMPO_FIORI_ELEMENTO)

        # A abertura da ordem de saída também pode trocar o frame; por isso o
        # segundo Exibir documento usa a mesma estratégia de reaquisição.
        pagina_app, frame, botao_exibir = localizar_botao_apos_transicao_fiori(
            pagina_app,
            padrao_nome=re.compile(r"^\s*Exibir\s+documento\s*$", re.I),
            descricao="Segundo botão 'Exibir documento'",
            timeout=45000,
        )
        botao_exibir.click(timeout=TEMPO_FIORI_ELEMENTO)

        # O visualizador atual possui botão Download. A rotina existente já
        # localiza os iframes dinâmicos (itshtmlvwrfnCxxx + iframe interno) sem
        # fixar IDs do Codegen e salva diretamente o arquivo da OV.
        baixar_pdf_visualizador_fiori(pagina_app, caminho_pdf)

        resultado["StatusSAP"] = "OK_VA03"
        resultado["ArquivoPDF"] = caminho_pdf
        resultado["Mensagem"] = (
            f"PDF baixado pelo Fiori. Documento: {faturamento['descricao_documento']}"
            + (f" {resultado['NumeroNF']}" if resultado["NumeroNF"] else "")
            + "."
        )
        print(f"[OK FIORI VA03] PDF salvo em: {caminho_pdf}")
        return resultado

    except Exception as e:
        evidencia = salvar_evidencia_fiori(
            pagina_app,
            prefixo="erro_va03_fiori",
            requisicao=requisicao,
            ordem=ordem,
        )
        resultado["StatusSAP"] = "ERRO_VA03"
        resultado["ArquivoPDF"] = ""
        resultado["Mensagem"] = f"ERRO VA03/Fiori: {e}"
        if evidencia:
            resultado["Mensagem"] += f" Evidência: {os.path.basename(evidencia)}"
        print(f"[FIORI VA03] OV {ordem}: {resultado['Mensagem']}")
        return resultado


def validar_ov_tabela_na_va03(page_fiori, ordem, atividade):
    """Valida candidato sem rótulo em consulta, antes de permitir faturamento."""
    pagina, frame = abrir_aplicativo_fiori(
        page_fiori, termo_pesquisa="va03",
        nomes_aplicativo=[
            re.compile(r"Exibir\s+ordens?\s+do\s+cliente(?:\s+VA03)?", re.I),
            re.compile(r"\bVA03\b", re.I),
        ],
        tipo_aplicativo="va03",
    )
    campo = preencher_campo_fiori(
        frame.get_by_role("textbox", name="Ordem", exact=True),
        ordem, descricao="OV candidata da tabela na VA03",
    )
    campo.press("Enter")
    aguardar_resumo_ordem_va03_carregado(
        pagina, frame_va03=frame, campo_ordem=campo, ordem=ordem,
    )
    _pagina_fluxo, fluxo = abrir_fluxo_documentos_fiori(pagina, frame_resumo=frame)
    ramo = _ramificacao_atual_fluxo(fluxo, atividade, ordem)
    if ramo.get("pai") is None:
        raise RuntimeError(
            f"OV candidata {ordem} não confirmou a ramificação da atividade {atividade}. "
            "VF01 não será executada."
        )
    print(f"[FIORI VA03] OV da tabela confirmada antes da VF01: {ramo['texto_pai']}")


class ValidacaoCreditoError(ValueError):
    """Pendência verificável da NC; não autoriza VF01 nem conclusão no PORTAL."""


def _texto_credito(valor):
    if valor is None or pd.isna(valor):
        return ""
    return str(valor).strip()


def _codigo_credito(valor, campo="Código"):
    texto = _texto_credito(valor)
    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".")[0]
    if not re.fullmatch(r"\d+", texto):
        raise ValidacaoCreditoError(f"{campo} vazio ou inválido: {texto!r}.")
    return texto.lstrip("0") or "0"


def _referencia_credito(valor):
    texto = _texto_credito(valor)
    if not texto:
        raise ValidacaoCreditoError("Referência vazia.")
    if re.fullmatch(r"\d+(?:\.0+)?", texto):
        return _codigo_credito(texto, "Referência")
    return texto.casefold()


def _valor_credito(valor):
    """Decimal em centavos; zero é válido, ausência/erro nunca vira zero."""
    if isinstance(valor, bool) or not _texto_credito(valor):
        raise ValidacaoCreditoError("Montante vazio ou inválido.")
    if isinstance(valor, str):
        texto = valor.strip().replace("\u00a0", "").replace(" ", "")
        if texto.endswith("-") and not texto.startswith(("-", "+")):
            texto = "-" + texto[:-1]
        if "," in texto:
            if not re.fullmatch(r"[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+),\d{1,2}", texto):
                raise ValidacaoCreditoError(f"Montante inválido: {valor!r}.")
            texto = texto.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})+", texto):
            texto = texto.replace(".", "")
        elif not re.fullmatch(r"[+-]?\d+(?:\.\d{1,2})?", texto):
            raise ValidacaoCreditoError(f"Montante inválido: {valor!r}.")
    else:
        texto = str(valor)
    try:
        numero = Decimal(texto)
        if not numero.is_finite() or numero != numero.quantize(Decimal("0.01")):
            raise ValueError("valor não finito ou com fração de centavo")
        return numero.quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError) as erro:
        raise ValidacaoCreditoError(f"Montante inválido: {valor!r}.") from erro


def _moeda_credito(valor):
    return format(Decimal(valor), ".2f").replace(".", ",")


def _cabecalho_credito(valor):
    """Nomes exatos dos layouts confirmados; nunca depende da posição.

    Conta é o BP do relatório mostrado pelo requisitante e será confrontada
    com o Emissor da ordem. O montante segue com sinal e centavos originais;
    não há conversão de moeda, tolerância ou escolha entre colunas ambíguas.
    """
    nome = re.sub(r"[^a-z0-9]", "", normalizar_texto_busca(_texto_credito(valor)))
    # Centro é a Divisão usada na regra já definida para a conciliação.
    # Se os dois cabeçalhos coexistirem, ler_detalhes_credito mantém o bloqueio
    # por ambiguidade: não escolhe uma coluna pela posição nem pelo conteúdo.
    return {'conta': 'cliente', 'montanteemmoedainterna': 'montanteme',
            'centro': 'divisao'}.get(nome, nome)


def validar_linhas_credito(linhas, indices, primeira_linha=2):
    """Valida todas as linhas de dados antes de consultar ou faturar a OV."""
    unicos = {campo: set() for campo in ("empresa", "divisao", "cliente")}
    referencias = {}
    total = Decimal("0.00")
    quantidade = 0
    total_informado = None
    linha_total = None
    preenchidas = [(n, linha) for n, linha in enumerate(linhas, primeira_linha)
                   if any(_texto_credito(v) for v in linha)]
    for posicao, (numero_linha, linha) in enumerate(preenchidas):
        try:
            valores = {campo: linha[indice] if indice < len(linha) else None
                       for campo, indice in indices.items()}
            # Export do SAP: última linha com apenas o montante total, como no
            # print, ou acompanhada somente de um rótulo de totalização.
            # Linhas com BP/referência/empresa/divisão ou outro conteúdo não
            # são descartadas como total. Um total divergente também bloqueia.
            outros = [_texto_credito(v) for i, v in enumerate(linha)
                      if i != indices['montanteme'] and _texto_credito(v)]
            somente_total = (bool(_texto_credito(valores['montanteme'])) and
                             all(_cabecalho_credito(v) in {'total', 'totalgeral', 'soma'} for v in outros))
            if somente_total:
                if not quantidade or posicao != len(preenchidas) - 1:
                    raise ValidacaoCreditoError('Linha de totalização fora do fim dos dados ou sem notas anteriores.')
                total_informado = _valor_credito(valores['montanteme'])
                if total_informado != total:
                    raise ValidacaoCreditoError(
                        f'Total informado {_moeda_credito(total_informado)} difere da soma das notas '
                        f'{_moeda_credito(total)}. Diferença={_moeda_credito(total_informado - total)}.')
                linha_total = numero_linha
                continue
            quantidade += 1
            for campo in unicos:
                unicos[campo].add(_codigo_credito(valores[campo], campo))
                if len(unicos[campo]) > 1:
                    raise ValidacaoCreditoError(
                        f"{campo}: valores diferentes entre as linhas ({', '.join(sorted(unicos[campo]))})."
                    )
            referencia = _referencia_credito(valores["referencia"])
            if referencia in referencias:
                raise ValidacaoCreditoError(
                    f"Referência repetida: {_texto_credito(valores['referencia'])}; "
                    f"já apareceu na linha {referencias[referencia]}."
                )
            referencias[referencia] = numero_linha
            total += _valor_credito(valores["montanteme"])
        except ValidacaoCreditoError as erro:
            raise ValidacaoCreditoError(f"Linha {numero_linha}: {erro}") from erro
    if not quantidade:
        raise ValidacaoCreditoError("Planilha de detalhes sem linhas de dados.")
    resultado = {
        "Empresa": next(iter(unicos["empresa"])),
        "Divisao": next(iter(unicos["divisao"])),
        "Cliente": next(iter(unicos["cliente"])),
        "Montante": str(total), "QuantidadeLinhas": quantidade,
        "QuantidadeReferencias": len(referencias),
    }
    if total_informado is not None:
        resultado.update(TotalInformado=str(total_informado), LinhaTotal=linha_total)
    return resultado


def ler_detalhes_credito(caminho):
    """Reconhece as cinco informações pelos cabeçalhos em qualquer ordem.

    OV isolada não é planilha de detalhes. Cabeçalho parcial, duplicado ou
    dado inválido gera pendência; não transforma uma NC validada em simples.
    """
    obrigatorios = {"empresa", "divisao", "cliente", "referencia", "montanteme"}
    encontrados = []
    try:
        with pd.ExcelFile(caminho) as arquivo:
            for aba in arquivo.sheet_names:
                tabela = pd.read_excel(arquivo, sheet_name=aba, header=None,
                                       dtype=object, keep_default_na=False)
                linhas = tabela.values.tolist()
                for indice, linha in enumerate(linhas[:60]):
                    nomes = [_cabecalho_credito(v) for v in linha]
                    presentes = obrigatorios.intersection(nomes)
                    if len(presentes) < 3:
                        continue
                    if presentes != obrigatorios:
                        faltantes = ", ".join(sorted(obrigatorios - presentes))
                        cabecalhos = ' | '.join(_texto_credito(v) for v in linha if _texto_credito(v))
                        raise ValidacaoCreditoError(
                            f'{os.path.basename(caminho)} | Aba {aba}, linha {indice + 1}: '
                            f'faltam colunas obrigatórias: {faltantes}. Cabeçalhos lidos: {cabecalhos}.')
                    if any(nomes.count(nome) != 1 for nome in obrigatorios):
                        duplicados = ', '.join(sorted(nome for nome in obrigatorios if nomes.count(nome) > 1))
                        raise ValidacaoCreditoError(
                            f'{os.path.basename(caminho)} | Aba {aba}: coluna obrigatória duplicada '
                            f'ou com mais de um cabeçalho equivalente: {duplicados}. Seleção ambígua.')
                    indices = {nome: nomes.index(nome) for nome in obrigatorios}
                    dados = validar_linhas_credito(
                        linhas[indice + 1:], indices,
                        primeira_linha=indice + 2,
                    )
                    mapa = {nome: {'Cabecalho': _texto_credito(linha[coluna]), 'Coluna': coluna + 1}
                            for nome, coluna in sorted(indices.items())}
                    dados.update(Arquivo=os.path.basename(caminho), Aba=aba, ColunasValidacao=mapa)
                    encontrados.append(dados)
                    break
    except ValidacaoCreditoError:
        raise
    except Exception as erro:
        raise ValidacaoCreditoError(f"Não foi possível ler {os.path.basename(caminho)}: {erro}") from erro
    if len(encontrados) > 1:
        raise ValidacaoCreditoError("Mais de uma aba de detalhes encontrada; seleção ambígua.")
    if encontrados:
        dados = encontrados[0]
        mapeamento = '; '.join(f'{nome} <- {item["Cabecalho"]} (coluna {item["Coluna"]})'
                              for nome, item in dados['ColunasValidacao'].items())
        print(f'[NC PLANILHA] {dados["Arquivo"]} | {dados["Aba"]}: {mapeamento}.')
        if dados.get('LinhaTotal'):
            print(f'[NC PLANILHA] Linha {dados["LinhaTotal"]}: total {_moeda_credito(dados["TotalInformado"])} '
                  f'conferido com {dados["QuantidadeLinhas"]} nota(s); não somado novamente.')
    return encontrados[0] if encontrados else None


def catalogar_excels_credito(andoc):
    """Conta arquivos da grade atual; não confunde extensão de e-mail com Excel."""
    entradas = andoc.locator("body").evaluate(r"""body => {
        const links = [...body.querySelectorAll('a')];
        const visivel = el => !!el.getClientRects().length &&
            getComputedStyle(el).visibility !== 'hidden';
        return [...body.querySelectorAll('tr')]
            .filter(row => row.querySelector('input[type="radio"]') && visivel(row))
            .map(row => [...row.querySelectorAll('a')].filter(visivel).map(link => ({
                indice: links.indexOf(link), texto: (link.innerText || '').trim(),
                titulo: (link.title || '').trim()
            })));
    }""")
    if not entradas:
        raise ValidacaoCreditoError("Grade de anexos sem linhas identificáveis; não foi possível contar os Excels.")
    candidatos = []
    for linha in entradas:
        arquivos = []
        for link in linha:
            for rotulo in (link.get("texto", ""), link.get("titulo", "")):
                if re.search(r"\.[a-zA-Z0-9]{2,5}$", rotulo):
                    arquivos.append((rotulo, link["indice"]))
                    break
        nomes = {nome.casefold() for nome, _ in arquivos}
        if len(nomes) != 1:
            raise ValidacaoCreditoError("Anexo sem extensão identificável ou com títulos ambíguos; contagem de Excel não confirmada.")
        nome, indice = arquivos[0]
        if os.path.splitext(nome)[1].lower() in EXTENSOES_PLANILHA_PORTAL:
            candidatos.append({"modo": "link", "indice": indice, "nome": nome, "rotulo": nome})
    return candidatos


class IdentificacaoModalidadeNCError(ValidacaoCreditoError):
    """Não foi possível distinguir a NC simples da NC com dois Excels."""
    pass


class DocumentosPORTALIndisponiveis(RuntimeError):
    pass


def ler_estado_documentos_portal(andoc):
    """Estado da grade, identidade e carregamento; compartilhado com o leitor NC."""
    return andoc.locator('body').evaluate(r'''body => {
                const vis = e => !!e.getClientRects().length && getComputedStyle(e).visibility !== 'hidden';
                const doc = body.ownerDocument;
                const cab = body.querySelector('#cabecalho');
                const cabecalho = cab && vis(cab) ? cab.innerText : '';
                const controles = [...body.querySelectorAll('button,input[type=button],input[type=submit],[role=button]')]
                    .filter(vis).map(e => ({id:e.id || '', nome:(e.innerText || e.value || e.title || '').trim()}));
                const linhas = [...body.querySelectorAll('tr')]
                    .filter(e => vis(e) && e.querySelector('input[type=radio]'));
                const ocupado = [...body.querySelectorAll('[aria-busy=true],[role=progressbar],.loading,.loader,.spinner')].some(vis)
                    || !!(doc.defaultView.jQuery && doc.defaultView.jQuery.active);
                return {pronto:doc.readyState === 'complete', cabecalho, controles, ocupado,
                    linhas:linhas.map(e => [...e.querySelectorAll('a')].filter(vis)
                        .map(a => ({texto:(a.innerText || '').trim(), titulo:a.title || ''}))),
                    linksArquivos:[...body.querySelectorAll('a')].filter(vis)
                        .map(a => (a.innerText || '') + ' ' + (a.title || ''))
                        .filter(s => /\.(?:xlsx|xlsm|xlsb|xls|pdf|msg|eml|docx|zip)\b/i.test(s)),
                    texto:body.innerText || ''};
            }''', timeout=2000)


def documentos_vazios_confirmados(estado, requisicao):
    if not isinstance(estado, dict) or estado.get('linhas') or estado.get('linksArquivos'):
        return False
    try:
        quantidade, _ = interpretar_lista_anexos_nc(estado, requisicao, None)
        return quantidade == 0
    except IdentificacaoModalidadeNCError:
        return False


def contar_excels_da_requisicao_nc(page, frame, requisicao):
    """Lê somente a lista de anexos; zero documentos é um estado válido.

    Não baixa arquivos, não lê células, não usa o downloader como detector e
    não executa F5. Uma tela vazia ainda carregando não vale como zero anexos.
    """
    validar_requisicao_portal_aberta(frame, requisicao, timeout=15000)
    fechar_modal_residual_portal(page, motivo=f'identificar modalidade NC {requisicao}', timeout=3000)
    portalmain = frame_portalmain(frame)

    def abrir_documentos():
        erros = []
        for localizar in (
            lambda: portalmain.get_by_title('Documentos'),
            lambda: portalmain.get_by_text('Documentos', exact=True),
        ):
            try:
                candidatos = localizar()
                visiveis = [candidatos.nth(i) for i in range(candidatos.count())
                            if candidatos.nth(i).is_visible()]
                if len(visiveis) == 1:
                    visiveis[0].click(timeout=10000)
                    portalmain.locator(FRAME_DOCUMENTOS).wait_for(state='attached', timeout=15000)
                    return
            except PlaywrightError as erro:
                erros.append(str(erro))
        raise IdentificacaoModalidadeNCError('A lista de anexos não pôde ser aberta para identificar a modalidade da NC.')

    abrir_documentos()
    andoc = portalmain.frame_locator(FRAME_DOCUMENTOS)
    limite = time.monotonic() + 20
    ultima = None
    estavel_desde = None
    ultimo_motivo = 'Aguardando a tela de Documentos carregar.'
    while time.monotonic() < limite:
        try:
            estado = ler_estado_documentos_portal(andoc)
            quantidade, assinatura = interpretar_lista_anexos_nc(estado, requisicao, andoc)
            if assinatura != ultima:
                ultima, estavel_desde = assinatura, time.monotonic()
            elif time.monotonic() - estavel_desde >= 1.0:
                print(f'[NC MODALIDADE] PORTAL {requisicao}: {quantidade} arquivo(s) Excel na lista atual de anexos.')
                return quantidade
        except (PlaywrightError, IdentificacaoModalidadeNCError, ValidacaoCreditoError) as erro:
            ultimo_motivo = str(erro)
            ultima = estavel_desde = None
        page.wait_for_timeout(250)
    raise IdentificacaoModalidadeNCError(
        f'Modalidade da NC não confirmada: {ultimo_motivo} Nenhuma planilha foi exigida como validação da NC simples.')


def interpretar_lista_anexos_nc(estado, requisicao, andoc):
    """Só aceita zero após confirmar a tela carregada e a identidade da requisição."""
    if not estado.get('pronto') or estado.get('ocupado'):
        raise IdentificacaoModalidadeNCError('Documentos ainda está carregando.')
    cabecalho = estado.get('cabecalho', '')
    if not re.search(r'(?<!\d)0*' + re.escape(str(requisicao)) + r'(?!\d)', cabecalho):
        raise IdentificacaoModalidadeNCError('Cabeçalho de Documentos não confirmou a requisição atual.')
    # A identificação é somente leitura. Os botões de escrita podem não estar
    # disponíveis enquanto a requisição está "Em andamento" e ainda sem responsável.
    linhas = estado.get('linhas', [])
    if linhas:
        quantidade = len(catalogar_excels_credito(andoc))
    else:
        # Há um arquivo mencionado mas a grade ainda não apareceu: não confundir
        # carregamento parcial com a NC simples sem planilha.
        if estado.get('linksArquivos') or re.search(r'\.(?:xlsx|xlsm|xlsb|xls)\b', estado.get('texto', ''), re.I):
            raise IdentificacaoModalidadeNCError('Planilha indicada na tela, mas sua linha de anexo ainda não carregou.')
        quantidade = 0
    assinatura = json.dumps({'linhas': linhas, 'quantidade': quantidade}, ensure_ascii=False, sort_keys=True)
    return quantidade, assinatura


def extrair_ovs_descricao_nc(descricao):
    """Aceita o sufixo observado 'OV - Centro: 1234' nas listas de NC.

    Remove apenas o Centro de linhas inteiras com uma OV e esse sufixo.
    A extração ainda exige o contexto OV/OVs; não recolhe números soltos,
    BPs, NFs ou documentos. O leitor e o fluxo de Locação ficam intactos.
    """
    texto = str(descricao or '').replace('\xa0', ' ')
    texto = re.sub(
        r'(?im)^([ \t]*\d{7,10})[ \t]*(?:[-–—][ \t]*)?'
        r'Centro[ \t]*:[ \t]*\d{4}[ \t]*\r?$', r'\1', texto)
    return extrair_ovs_suporte_venda(texto)


def obter_ovs_credito_simples(page, frame, requisicao, pasta, descricao, atividade):
    """Fluxo já aprovado: descrição primeiro; Excel apenas para localizar uma OV ausente."""
    ovs = extrair_ovs_descricao_nc(descricao)
    if ovs:
        return ovs, 'PORTAL_DESCRICAO', None
    ovs = extrair_ovs_tabela_descricao(descricao)
    if ovs:
        return ovs, 'PORTAL_DESCRICAO_TABELA', None
    ovs = extrair_candidatos_ov_descricao(descricao, atividade)
    if ovs:
        return ovs, 'PORTAL_DESCRICAO_CANDIDATA', None
    planilhas = localizar_planilhas_locais_requisicao(pasta)
    ovs, _ = extrair_ovs_de_planilhas(planilhas, atividade=atividade)
    origem = 'PORTAL_PLANILHA_LOCAL'
    if not ovs:
        _salvos, ovs, _refs = baixar_planilhas_anexas_portal(
            page, frame, requisicao, pasta, atividade=atividade)
        origem = 'PORTAL_PLANILHA'
    return ovs, origem if ovs else 'SEM_OV_LOCALIZADA', None


def preparar_validacao_credito(page, frame, requisicao, pasta, descricao, atividade):
    """0/1 Excel: NC simples. Exatamente dois: conferência e conciliação da nova modalidade."""
    quantidade = contar_excels_da_requisicao_nc(page, frame, requisicao)
    if quantidade < 2:
        resultado = obter_ovs_credito_simples(page, frame, requisicao, pasta, descricao, atividade)
        print(f'[NC MODALIDADE] PORTAL {requisicao}: NC_SIMPLES; origem={resultado[1]}; '
              'sem conferência de valores em planilha e sem abertura de conciliação.')
        return resultado
    if quantidade > 2:
        raise ValidacaoCreditoError(f'Foram encontrados {quantidade} Excels; não há regra definida para essa combinação.')
    print(f'[NC MODALIDADE] PORTAL {requisicao}: NC_DUAS_PLANILHAS; conferência prévia obrigatória.')
    inspecao = {}
    arquivos, _ovs, _refs = baixar_planilhas_anexas_portal(
        page, frame, requisicao, pasta, atividade=atividade,
        baixar_todas=True, inspecao=inspecao,
    )
    if not inspecao.get("completa"):
        raise ValidacaoCreditoError(inspecao.get("erro") or
                                    "Não foi possível confirmar e baixar todos os Excels da requisição.")
    if len(arquivos) != 2 or inspecao.get('qtd_excel') != 2:
        raise ValidacaoCreditoError('A lista de anexos mudou ou o download dos dois Excels ficou incompleto. Não executar a NC como simples.')
    detalhes = []
    planilhas_ov = []
    for caminho in arquivos:
        hash_antes = hash_arquivo_nc(caminho) if len(arquivos) == 2 else ""
        dados = ler_detalhes_credito(caminho)
        if dados is None:
            planilhas_ov.append(caminho)
        else:
            dados["CaminhoPlanilha"] = os.path.abspath(caminho)
            if len(arquivos) == 2:
                if hash_arquivo_nc(caminho) != hash_antes:
                    raise ValidacaoCreditoError('A planilha foi alterada durante a leitura.')
                dados['HashPlanilhaValidada'] = hash_antes
            detalhes.append(dados)
    ovs_descricao = extrair_ovs_descricao_nc(descricao)
    origem = "PORTAL_DESCRICAO"
    if not ovs_descricao:
        ovs_descricao = extrair_ovs_tabela_descricao(descricao)
        origem = "PORTAL_DESCRICAO_TABELA"
    if len(arquivos) != 2 or len(detalhes) != 1 or len(planilhas_ov) != 1:
        raise ValidacaoCreditoError("A segunda opção exige dois Excels: uma planilha de OV e uma de detalhes identificada pelas cinco colunas.")
    ovs_planilha, _ = extrair_ovs_de_planilhas(planilhas_ov, atividade=atividade)
    normalizar = lambda valores: list(dict.fromkeys(_codigo_credito(v, "OV") for v in valores))
    ovs_descricao, ovs_planilha = normalizar(ovs_descricao), normalizar(ovs_planilha)
    if ovs_descricao and ovs_planilha and set(ovs_descricao) != set(ovs_planilha):
        raise ValidacaoCreditoError("A OV informada na descrição diverge da OV da planilha.")
    ovs = ovs_descricao or ovs_planilha
    if len(ovs) != 1:
        raise ValidacaoCreditoError(f"A segunda opção exige uma OV; foram identificadas {len(ovs)}: {ovs}.")
    dados = detalhes[0]
    dados["Ordem"] = ovs[0]
    print(f"[NC PRÉ-VALIDAÇÃO] PORTAL {requisicao}; OV {ovs[0]}; Empresa {dados['Empresa']}; "
          f"Divisão {dados['Divisao']}; Cliente {dados['Cliente']}; "
          f"Referências únicas {dados['QuantidadeReferencias']}; Montante {_moeda_credito(dados['Montante'])}.")
    return ovs, "PORTAL_NC_DUAS_PLANILHAS", dados


def _campo_credito_va03(pagina, nome, frame_preferencial=None, timeout=15000):
    """Readquire o frame visível por campo; não fixa o nome dinâmico do iframe."""
    limite = time.monotonic() + timeout / 1000
    while time.monotonic() < limite:
        contextos = ([frame_preferencial] if frame_preferencial is not None else [])
        contextos += [fr for fr in pagina.frames if fr not in contextos]
        for contexto in contextos:
            try:
                if not _frame_fiori_realmente_visivel(contexto):
                    continue
                campos = contexto.get_by_role("textbox", name=nome, exact=True)
                visiveis = [campos.nth(i) for i in range(campos.count()) if campos.nth(i).is_visible()]
            except PlaywrightError:
                continue
            if len(visiveis) > 1:
                raise ValidacaoCreditoError(f"Mais de um campo {nome!r} visível na VA03.")
            if visiveis:
                return contexto, visiveis[0]
        pagina.wait_for_timeout(250)
    raise ValidacaoCreditoError(f"Campo {nome!r} não localizado na VA03.")


def conferir_dados_ov_credito(esperado, cliente, liquido, imposto):
    cliente_lido = _codigo_credito(cliente, "Emissor da ordem/BP")
    liquido_lido, imposto_lido = _valor_credito(liquido), _valor_credito(imposto)
    total = liquido_lido + imposto_lido
    previsto = _valor_credito(esperado["Montante"])
    mensagem = (f"BP planilha={esperado['Cliente']}; BP OV={cliente_lido}; "
                f"Montante planilha={_moeda_credito(previsto)}; "
                f"Líquido={_moeda_credito(liquido_lido)}; Imposto={_moeda_credito(imposto_lido)}; "
                f"Total OV={_moeda_credito(total)}; Diferença={_moeda_credito(total - previsto)}.")
    if cliente_lido != esperado["Cliente"] or total != previsto:
        raise ValidacaoCreditoError("Divergência na conferência prévia. " + mensagem)
    return "APROVADA antes da VF01. " + mensagem


def aguardar_condicoes_item_credito(pagina, timeout=15000, frame_preferencial=None):
    """Localiza a aba pelo seletor gravado, no iframe visível após abrir o item."""
    limite = time.monotonic() + timeout / 1000
    diagnostico = "Nenhuma barra de abas visível."
    while time.monotonic() < limite:
        encontrados = []
        contextos = ([frame_preferencial] if frame_preferencial is not None else [])
        contextos += [fr for fr in pagina.frames if fr not in contextos]
        for contexto in contextos:
            try:
                if not _frame_fiori_realmente_visivel(contexto):
                    continue
                # Mesmo seletor da gravação do usuário: busca por texto contido.
                # Não exige que o título Dados do item esteja no mesmo iframe.
                barras = contexto.get_by_role("tablist")
                abas = barras.get_by_text("Condições")
                visiveis = [abas.nth(i) for i in range(abas.count()) if abas.nth(i).is_visible()]
                if not visiveis:
                    abas = contexto.get_by_role("tab", name="Condições")
                    visiveis = [abas.nth(i) for i in range(abas.count()) if abas.nth(i).is_visible()]
                if not visiveis:
                    rotulos = []
                    for i in range(min(barras.count(), 3)):
                        barra = barras.nth(i)
                        if barra.is_visible():
                            rotulos.append(re.sub(r"\s+", " ", barra.inner_text(timeout=1000)).strip()[:300])
                    if rotulos:
                        diagnostico = "Abas visíveis: " + " | ".join(rotulos)
            except PlaywrightError:
                # O SAP pode substituir o iframe durante a abertura do item.
                continue
            if len(visiveis) > 1:
                raise ValidacaoCreditoError("Mais de uma aba Condições visível nos dados do item da VA03.")
            if visiveis:
                encontrados.append((contexto, visiveis[0]))
        if len(encontrados) > 1:
            raise ValidacaoCreditoError("Mais de uma tela de item da VA03 visível; conferência interrompida.")
        if encontrados:
            return encontrados[0]
        pagina.wait_for_timeout(250)
    raise ValidacaoCreditoError(
        "Aba Condições não localizada após o duplo clique no Material; VF01 não autorizada. " + diagnostico)


def validar_credito_na_va03(page_fiori, requisicao, ordem, esperado):
    if _codigo_credito(ordem, "OV") != esperado.get("Ordem"):
        raise ValidacaoCreditoError("A validação pertence a outra OV.")
    pagina, frame = abrir_aplicativo_fiori(
        page_fiori, termo_pesquisa="va03",
        nomes_aplicativo=[re.compile(r"Exibir\s+ordens?\s+do\s+cliente(?:\s+VA03)?", re.I),
                         re.compile(r"\bVA03\b", re.I)], tipo_aplicativo="va03",
    )
    campo = preencher_campo_fiori(frame.get_by_role("textbox", name="Ordem", exact=True),
                                 ordem, descricao="OV da nota de crédito na conferência prévia")
    campo.press("Enter")
    aguardar_resumo_ordem_va03_carregado(pagina, frame_va03=frame, campo_ordem=campo, ordem=ordem)
    # Confirma a identificação exibida, além da leitura do valor recém-preenchido.
    padrao = re.compile(rf"\b{re.escape(str(ordem))}\b")
    if _primeiro_item_visivel([frame.get_by_text(padrao), pagina.get_by_text(padrao)], limite_itens=20) is None:
        raise ValidacaoCreditoError(f"O cabeçalho da VA03 não confirmou a OV {ordem}.")
    frame, campo_bp = _campo_credito_va03(pagina, "Emissor da ordem", frame)
    cliente = campo_bp.input_value(timeout=5000)
    # O usuário confirmou uma única linha de serviço neste fluxo.
    frame, material = _campo_credito_va03(pagina, "Material", frame)
    print(f"[NC PRÉ-VALIDAÇÃO] PORTAL {requisicao}; OV {ordem}: abrindo o item com duplo clique no Material.")
    material.dblclick(timeout=10000)
    frame, aba_condicoes = aguardar_condicoes_item_credito(pagina, frame_preferencial=frame)
    print(f"[NC PRÉ-VALIDAÇÃO] PORTAL {requisicao}; OV {ordem}: aba Condições localizada; executando clique.")
    aba_condicoes.click(timeout=10000)
    print(f"[NC PRÉ-VALIDAÇÃO] PORTAL {requisicao}; OV {ordem}: clique em Condições executado; aguardando Líquido e Imposto.")
    frame, liquido = _campo_credito_va03(pagina, "Líquido", frame)
    frame, imposto = _campo_credito_va03(pagina, "Imposto", frame)
    print(f"[NC PRÉ-VALIDAÇÃO] PORTAL {requisicao}; OV {ordem}: campos Líquido e Imposto localizados; conferindo valores e BP.")
    # Duas leituras iguais evitam conferir números durante uma atualização de tela.
    leitura_anterior = None
    limite = time.monotonic() + 10
    while time.monotonic() < limite:
        leitura = (liquido.input_value(timeout=3000), imposto.input_value(timeout=3000))
        if leitura == leitura_anterior and all(str(v).strip() for v in leitura):
            mensagem = conferir_dados_ov_credito(esperado, cliente, *leitura)
            print(f"[NC PRÉ-VALIDAÇÃO] PORTAL {requisicao}; OV {ordem}: {mensagem}")
            return mensagem
        leitura_anterior = leitura
        pagina.wait_for_timeout(500)
    raise ValidacaoCreditoError("Líquido e Imposto não estabilizaram; VF01 não autorizada.")


def preparar_pdf_credito_para_portal(requisicao, consulta, pdf):
    """Identifica o PDF confirmado por requisição, OV e documento SAP."""
    req = _codigo_credito(requisicao, "Requisição")
    ov = _codigo_credito(consulta.get("Ordem"), "OV")
    doc = _codigo_credito(consulta.get("NumeroNF"), "Documento SAP").zfill(10)
    vf01 = consulta.get("DocumentoFaturamentoVF01")
    if vf01 and _codigo_credito(vf01) != _codigo_credito(doc):
        raise ValidacaoCreditoError("Documento da VF01 diverge do PDF localizado na VA03.")
    if consulta.get("StatusSAP") != "OK_VA03" or not pdf_valido(pdf):
        raise ValidacaoCreditoError("PDF da NC sem confirmação válida da VA03.")
    pasta = os.path.join(os.path.dirname(pdf), "anexos_portal")
    os.makedirs(pasta, exist_ok=True)
    destino = os.path.join(pasta, f"NC_PORTAL{req}_OV{ov}_DOC{doc}.pdf")
    shutil.copyfile(pdf, destino)
    consulta["ArquivoAnexoPORTAL"] = destino
    return destino


def processar_ov_sap(page_fiori, requisicao, ordem, atividade, validar_tabela=False, validacao_credito=None,
                     registrar_etapa=None):
    """VF01 -> classifica resultado -> VA03 somente quando houver documento utilizável.

    TESTE 41:
    - documento novo: VA03 procura o número exato retornado pela VF01;
    - completo sem número: VA03 usa a última FAT válida + Doc.contábil da ramificação;
    - incompleto: pula a OV, sem segunda VF01 e sem VA03;
    - pendência na VA03: não refaz VF01 automaticamente; mantém a OV em aberto.
    """
    requisicao = str(requisicao).strip()
    ordem = str(ordem).strip()
    atividade = str(atividade or "").strip()

    obter_config_atividade_suporte_venda(atividade)

    if FIORI_PENDENCIA_GRAVACAO:
        mensagem = (
            "OV não processada: existe resultado VF01 incerto da requisição "
            f"{FIORI_PENDENCIA_GRAVACAO.get('requisicao', '')} / OV "
            f"{FIORI_PENDENCIA_GRAVACAO.get('ordem', '')}. "
            "Tela preservada para conferência; não houve nova tentativa de faturamento."
        )
        print(f"[FIORI] OV {ordem}: {mensagem}")
        return {
            "Ordem": ordem,
            "Atividade": atividade,
            "StatusSAP": "NAO_PROCESSADA_ESTADO_SAP_INCERTO",
            "NumeroNF": "",
            "ArquivoPDF": "",
            "Mensagem": mensagem,
        }

    mensagem_validacao_nc = ""
    if validacao_credito:
        try:
            mensagem_validacao_nc = validar_credito_na_va03(
                page_fiori, requisicao, ordem, validacao_credito)
            validacao_credito["ResultadoVA03"] = mensagem_validacao_nc
        except Exception as erro:
            evidencia = salvar_evidencia_fiori(
                page_fiori, prefixo="nc_validacao_previa",
                requisicao=requisicao, ordem=ordem)
            return {
                "Ordem": ordem, "Atividade": atividade,
                "StatusSAP": "NC_VALIDACAO_REPROVADA", "NumeroNF": "", "ArquivoPDF": "",
                "ValidacaoNC": str(erro),
                "Mensagem": f"VF01 não executada: {erro}. Evidência: "
                            f"{os.path.basename(evidencia) if evidencia else 'indisponível'}",
            }

    if validar_tabela:
        try:
            validar_ov_tabela_na_va03(page_fiori, ordem, atividade)
        except Exception as erro:
            evidencia = salvar_evidencia_fiori(
                page_fiori, prefixo="ov_tabela_nao_confirmada",
                requisicao=requisicao, ordem=ordem,
            )
            return {
                "Ordem": ordem, "Atividade": atividade,
                "StatusSAP": "OV_TABELA_NAO_CONFIRMADA", "NumeroNF": "",
                "ArquivoPDF": "",
                "Mensagem": f"Consulta prévia VA03 não validou a OV; sem VF01: {erro}. "
                            f"Evidência: {os.path.basename(evidencia) if evidencia else 'indisponível'}",
            }

    if registrar_etapa:
        registrar_etapa('VF01_INICIADA', None)
    resultado_vf01 = baixar_nota_sap_vf01(page_fiori, requisicao, ordem)
    if registrar_etapa:
        registrar_etapa('VF01_FINALIZADA', resultado_vf01)
    status_vf01 = str((resultado_vf01 or {}).get("StatusSAP", "") or "ERRO_VF01").strip()
    mensagem_vf01 = str((resultado_vf01 or {}).get("Mensagem", "") or "").strip()

    if status_vf01 == 'VF01_DOCUMENTO_INEXISTENTE':
        return {'Ordem': ordem, 'Atividade': atividade, 'StatusSAP': 'DOCUMENTO_INEXISTENTE',
                'NumeroNF': '', 'ArquivoPDF': '', 'DocumentoFaturamentoVF01': '',
                'RejeicaoSemGravacao': bool(resultado_vf01.get('RejeicaoSemGravacao')),
                'Mensagem': 'VF01: número rejeitado pelo SAP; Gravar não enviado. ' + mensagem_vf01}

    if status_vf01 == "VF01_GRAVADA_ERRO_DETERMINACAO_CONTA":
        documento = str(resultado_vf01.get("DocumentoFaturamentoVF01", "") or "")
        print(
            f"[FIORI] OV {ordem}: documento {documento} gravado com erro contábil. "
            "Requisição permanecerá aberta; seguindo a fila sem repetir a gravação."
        )
        return {
            "Ordem": ordem, "Atividade": atividade,
            "StatusSAP": STATUS_ERRO_DETERMINACAO_CONTA,
            "NumeroNF": documento.zfill(10) if documento else "",
            "DocumentoFaturamentoVF01": documento,
            "ArquivoPDF": "",
            "Mensagem": (
                f"VF01[{status_vf01}]: {mensagem_vf01} | "
                "Documento gravado com erro de determinação de conta. "
                "Sem nova gravação e sem conclusão do PORTAL; requer tratamento no SAP."
            ),
        }

    if not resultado_vf01 or not resultado_vf01.get("VF01Executada"):
        return {
            "Ordem": ordem,
            "StatusSAP": status_vf01,
            "NumeroNF": "",
            "ArquivoPDF": "",
            "Atividade": atividade,
            "Mensagem": "VF01 não liberou consulta segura da VA03. " + mensagem_vf01,
        }

    if status_vf01 == "VF01_FATURAMENTO_INCOMPLETO":
        mensagem = (
            f"VF01[{status_vf01}]: {mensagem_vf01} | "
            "OV pulada conforme regra de negócio; não houve nova gravação nem consulta VA03."
        )
        print(f"[FIORI] OV {ordem}: faturamento incompleto; pulando a OV.")
        return {
            "Ordem": ordem,
            "StatusSAP": "FATURAMENTO_INCOMPLETO",
            "NumeroNF": "",
            "ArquivoPDF": "",
            "Atividade": atividade,
            "Mensagem": mensagem,
        }

    documento_vf01 = str(
        (resultado_vf01 or {}).get("DocumentoFaturamentoVF01", "") or ""
    ).strip()

    if documento_vf01:
        print(
            f"[FIORI] OV {ordem}: VF01 informou documento {documento_vf01}; "
            f"VA03 procurará exatamente {documento_vf01.zfill(10)}."
        )
    else:
        print(
            f"[FIORI] OV {ordem}: VF01 confirmou estado completo sem número novo; "
            "VA03 procurará a última FAT válida da ramificação ligada ao Doc.contábil."
        )

    resultado_va03 = baixar_nota_existente_va03(
        page_fiori,
        requisicao,
        ordem,
        atividade,
        documento_faturamento_vf01=documento_vf01,
    )

    # Não refaz VF01 automaticamente quando a VA03 ainda estiver pendente.
    # A OV fica registrada para tratamento posterior e o PORTAL não será concluído sem PDF.
    historico = [
        f"VF01[{status_vf01}]: {mensagem_vf01}",
        "VA03[{}]: {}".format(
            resultado_va03.get("StatusSAP", ""),
            resultado_va03.get("Mensagem", ""),
        ),
    ]
    if mensagem_validacao_nc:
        historico.insert(0, "NC[" + mensagem_validacao_nc + "]")
        resultado_va03["ValidacaoNC"] = mensagem_validacao_nc
    resultado_va03["Mensagem"] = " | ".join(historico)
    resultado_va03["Atividade"] = atividade
    resultado_va03["DocumentoFaturamentoVF01"] = documento_vf01
    return resultado_va03

class RequisicaoEmExecucaoError(RuntimeError):
    pass


class ControleRetomadaOVError(RuntimeError):
    pass


class TravaRequisicao:
    """Trava de processo para todas as atividades, mantida até salvar o resultado."""
    def __init__(self, pasta):
        self.pasta, self.arquivo = pasta, None

    def __enter__(self):
        os.makedirs(self.pasta, exist_ok=True)
        self.arquivo = open(os.path.join(self.pasta, 'controle_execucao.lock'), 'a+b')
        try:
            self.arquivo.seek(0, os.SEEK_END)
            if self.arquivo.tell() == 0:
                self.arquivo.write(b'0')
                self.arquivo.flush()
            self.arquivo.seek(0)
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self.arquivo.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.arquivo, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as erro:
            self.arquivo.close()
            self.arquivo = None
            raise RequisicaoEmExecucaoError(
                'Outra execução usa esta pasta de requisição, ou a trava não pôde ser obtida. '
                'Nenhuma ação SAP/PORTAL foi iniciada nesta execução.') from erro
        return self

    def __exit__(self, *_):
        if self.arquivo is not None:
            self.arquivo.close()
            self.arquivo = None


def carregar_controle_ovs(pasta, requisicao, atividade):
    """Não importa PDFs avulsos nem infere faturamento pelo nome de arquivo."""
    caminho = os.path.join(pasta, 'controle_ovs.json')
    tipo = classificar_atividade_suporte_venda(atividade)
    if not os.path.exists(caminho):
        return {'VersaoControle': 1, 'Requisicao': str(requisicao), 'Atividade': tipo, 'OVs': {}}
    try:
        with open(caminho, encoding='utf-8') as arquivo:
            estado = json.load(arquivo)
        if (estado.get('VersaoControle') != 1 or estado.get('Requisicao') != str(requisicao)
                or estado.get('Atividade') != tipo or not isinstance(estado.get('OVs'), dict)):
            raise ValueError('Identidade/atividade/versão incompatível.')
        for ov, entrada in estado['OVs'].items():
            if not re.fullmatch(r'\d{7,10}', ov) or not isinstance(entrada, dict):
                raise ValueError('Registro de OV inválido.')
            if entrada.get('Fase') not in {'VF01_INICIADA', 'VF01_FINALIZADA', 'CONSULTA_FINALIZADA'}:
                raise ValueError('Etapa de OV desconhecida.')
            for campo in ('ConsultaSAP', 'ResultadoVF01'):
                resultado = entrada.get(campo)
                if resultado:
                    validar_associacao_resultado_sap(requisicao, ov, resultado)
            consulta = entrada.get('ConsultaSAP') or {}
            vf01 = entrada.get('ResultadoVF01') or {}
            documento = str(vf01.get('DocumentoFaturamentoVF01') or '').strip()
            numero = str(consulta.get('NumeroNF') or '').strip()
            if documento and not re.fullmatch(r'\d{8,10}', documento):
                raise ValueError('Documento VF01 inválido no controle.')
            if (documento and numero and _normalizar_numero_documento_sap(documento)
                    != _normalizar_numero_documento_sap(numero)):
                raise ValueError('Documento da consulta diverge do faturamento registrado.')
        return estado
    except Exception as erro:
        raise ControleRetomadaOVError(
            f'Controle de OVs inválido: {erro}. Conferir o registro antes de repetir SAP.') from erro


def salvar_controle_ovs(pasta, estado, ordem, entrada):
    novo = dict(estado, OVs=dict(estado['OVs']))
    novo['OVs'][str(ordem)] = dict(entrada, AtualizadoEm=datetime.now().isoformat(timespec='seconds'))
    destino = os.path.join(pasta, 'controle_ovs.json')
    with open(destino + '.tmp', 'w', encoding='utf-8') as arquivo:
        json.dump(novo, arquivo, ensure_ascii=False, indent=2)
        arquivo.flush()
        os.fsync(arquivo.fileno())
    os.replace(destino + '.tmp', destino)
    estado.clear()
    estado.update(novo)


def vf01_confirmada_para_retomada(resultado):
    return (isinstance(resultado, dict) and resultado.get('VF01Executada') is True
            and resultado.get('StatusSAP') in {'VF01_GRAVACAO_CONFIRMADA',
                'VF01_SEM_NOVO_FATURAMENTO', 'VF01_GRAVADA_ERRO_DETERMINACAO_CONTA'})


def processar_ov_com_retomada(page_fiori, requisicao, ordem, atividade, pasta_requisicao,
                            validar_tabela=False, validacao_credito=None):
    """Retoma apenas a OV atual: PDF íntegro ou consulta VA03 após VF01 confirmada."""
    if validacao_credito:
        raise ControleRetomadaOVError('NC com validação deve usar o controle exclusivo de duas planilhas.')
    requisicao, ordem = str(requisicao).strip(), str(ordem).strip()
    estado = carregar_controle_ovs(pasta_requisicao, requisicao, atividade)
    entrada = dict(estado['OVs'].get(ordem) or {})
    anterior = entrada.get('ConsultaSAP') or {}
    vf01 = entrada.get('ResultadoVF01') or {}

    if FIORI_PENDENCIA_GRAVACAO:
        return dict(Ordem=ordem, Atividade=atividade, StatusSAP='NAO_PROCESSADA_ESTADO_SAP_INCERTO',
                    NumeroNF='', ArquivoPDF='', Mensagem='Existe resultado SAP incerto; nenhuma nova ação executada.')

    if entrada:
        # Só documentos emitidos e registrados por esta versão podem ser reutilizados.
        numero = str(anterior.get('NumeroNF') or '')
        pdf = anterior.get('ArquivoPDF') or ''
        hash_pdf = entrada.get('HashPDF') or ''
        if (anterior.get('StatusSAP') == 'OK_VA03' and re.fullmatch(r'\d{8,10}', numero)
                and vf01_confirmada_para_retomada(vf01) and pdf_valido(pdf) and hash_pdf
                and hash_arquivo_nc(pdf) == hash_pdf):
            print(f'[SAP RETOMADA] PORTAL {requisicao} / OV {ordem}: documento {numero} e PDF íntegro '
                  'confirmados no controle; seguindo para conferência dos anexos no PORTAL.')
            return dict(anterior, RetomadaSAP=True, Mensagem='Resultado SAP/PDF retomado do controle. | '
                        + str(anterior.get('Mensagem', '')))

        if vf01_confirmada_para_retomada(vf01):
            documento = str(vf01.get('DocumentoFaturamentoVF01') or numero or '').strip()
            print(f'[SAP RETOMADA] PORTAL {requisicao} / OV {ordem}: somente VA03/PDF; VF01 não será repetida.')
            resultado = baixar_nota_existente_va03(page_fiori, requisicao, ordem, atividade,
                                                  documento_faturamento_vf01=documento)
            resultado.update(Atividade=atividade, RetomadaSAP=True,
                             DocumentoFaturamentoVF01=vf01.get('DocumentoFaturamentoVF01', ''))
            resultado['Mensagem'] = 'Retomada somente pela VA03. | ' + str(resultado.get('Mensagem', ''))
        elif (anterior.get('StatusSAP') == 'DOCUMENTO_INEXISTENTE'
              and anterior.get('RejeicaoSemGravacao') is True
              and vf01.get('StatusSAP') == 'VF01_DOCUMENTO_INEXISTENTE'
              and vf01.get('RejeicaoSemGravacao') is True):
            # Preserva a recuperação da OV correta na planilha sem repetir o número rejeitado.
            return dict(anterior, RetomadaSAP=True)
        elif (entrada.get('Fase') == 'CONSULTA_FINALIZADA' and not vf01
              and anterior.get('StatusSAP') == 'OV_TABELA_NAO_CONFIRMADA'):
            entrada = {}  # Falha de consulta prévia; VF01 não foi chamada.
        else:
            return dict(Ordem=ordem, Atividade=atividade, StatusSAP='PENDENTE_CONFERENCIA_SAP',
                        NumeroNF='', ArquivoPDF='', RetomadaSAP=True,
                        Mensagem='Etapa SAP anterior sem liberação segura para retomar. VF01 não repetida. '
                        + str(anterior.get('Mensagem') or vf01.get('Mensagem') or entrada.get('Fase')))

    if not entrada:
        def registrar_etapa(fase, resultado_vf01):
            entrada['Fase'] = fase
            if resultado_vf01 is not None:
                validar_associacao_resultado_sap(requisicao, ordem, resultado_vf01)
                entrada['ResultadoVF01'] = dict(resultado_vf01)
            salvar_controle_ovs(pasta_requisicao, estado, ordem, entrada)

        try:
            resultado = processar_ov_sap(
                page_fiori=page_fiori, requisicao=requisicao, ordem=ordem, atividade=atividade,
                validar_tabela=validar_tabela, validacao_credito=None, registrar_etapa=registrar_etapa)
        except BaseException:
            # A interrupção entre o comando e a confirmação mantém a intenção em disco.
            if entrada.get('Fase') == 'VF01_INICIADA':
                FIORI_PENDENCIA_GRAVACAO.update(requisicao=requisicao, ordem=ordem,
                                              motivo='Interrupção antes de registrar o retorno da VF01.')
            raise

    validar_associacao_resultado_sap(requisicao, ordem, resultado)
    pdf = resultado.get('ArquivoPDF') or ''
    entrada.update(Fase='CONSULTA_FINALIZADA', ConsultaSAP=dict(resultado),
                   HashPDF=hash_arquivo_nc(pdf) if resultado.get('StatusSAP') == 'OK_VA03' and pdf_valido(pdf) else '')
    salvar_controle_ovs(pasta_requisicao, estado, ordem, entrada)
    return resultado


# TESTE 51: funções exclusivas da NC com dois arquivos Excel.
class ConciliacaoNCError(RuntimeError):
    pass


class TravaNC:
    """Uma execução por requisição, inclusive entre dois processos no Windows."""
    def __init__(self, pasta):
        self.pasta = pasta
        self.arquivo = None

    def __enter__(self):
        os.makedirs(self.pasta, exist_ok=True)
        self.arquivo = open(os.path.join(self.pasta, 'controle_nc51.lock'), 'a+b')
        try:
            self.arquivo.seek(0, os.SEEK_END)
            if self.arquivo.tell() == 0:
                self.arquivo.write(b'0')
                self.arquivo.flush()
            self.arquivo.seek(0)
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self.arquivo.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.arquivo, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except Exception as erro:
            self.arquivo.close()
            self.arquivo = None
            raise ConciliacaoNCError('Outra execução está usando o controle desta NC; não houve nova ação.') from erro
        return self

    def __exit__(self, *_):
        if self.arquivo is not None:
            # Fechar libera a trava também em caso de exceção/interrupção.
            self.arquivo.close()


def caminho_controle_nc(pasta):
    return os.path.join(pasta, 'controle_nc_duas_planilhas.json')


def hash_arquivo_nc(caminho):
    resultado = hashlib.sha256()
    with open(caminho, 'rb') as arquivo:
        for bloco in iter(lambda: arquivo.read(1024 * 1024), b''):
            resultado.update(bloco)
    return resultado.hexdigest()


def carregar_controle_nc(pasta, requisicao):
    caminho = caminho_controle_nc(pasta)
    if not os.path.exists(caminho):
        return None
    try:
        with open(caminho, encoding='utf-8') as arquivo:
            estado = json.load(arquivo)
        fases = {'SAP_EM_EXECUCAO', 'VALIDACAO_REPROVADA', 'SAP_PENDENTE',
                 'SAP_CONFIRMADO', 'PORTAL_FINALIZACAO', 'ORIGINAL_ENCERRADA',
                 'ENVIO_INICIADO', 'CRIADA_ANEXO_PENDENTE', 'CONCLUIDA'}
        if (estado.get('VersaoControle') != 1 or estado.get('Requisicao') != str(requisicao)
                or estado.get('OrigemOVs') != 'PORTAL_NC_DUAS_PLANILHAS'
                or classificar_atividade_suporte_venda(estado.get('Atividade', '')) != ATIVIDADE_NOTA_CREDITO
                or estado.get('Fase') not in fases):
            raise ValueError('Identidade, versão ou etapa incompatível.')
        return estado
    except Exception as erro:
        raise ConciliacaoNCError(f'Controle NC inválido: {erro}. Conferir antes de repetir qualquer etapa.') from erro


def salvar_controle_nc(pasta, estado, **alteracoes):
    novo = dict(estado, **alteracoes)
    novo['AtualizadoEm'] = datetime.now().isoformat(timespec='seconds')
    destino = caminho_controle_nc(pasta)
    temporario = destino + '.tmp'
    with open(temporario, 'w', encoding='utf-8') as arquivo:
        json.dump(novo, arquivo, ensure_ascii=False, indent=2)
        arquivo.flush()
        os.fsync(arquivo.fileno())
    os.replace(temporario, destino)
    estado.clear()
    estado.update(novo)


def incluir_pendencias_conciliacao_nc(requisicoes):
    """Requisições já encerradas somem da fila; recupera somente controles desta versão."""
    retorno = list(requisicoes)
    if not os.path.isdir(PASTA_REQUISICOES):
        return retorno
    for entrada in sorted(os.scandir(PASTA_REQUISICOES), key=lambda e: e.name):
        if not entrada.is_dir() or not re.fullmatch(r'\d{6,10}', entrada.name):
            continue
        if not os.path.isfile(caminho_controle_nc(entrada.path)):
            continue
        try:
            estado = carregar_controle_nc(entrada.path, entrada.name)
            if estado['Fase'] in {'CONCLUIDA', 'VALIDACAO_REPROVADA'}:
                continue
        except ConciliacaoNCError as erro:
            print(f'[NC CONTROLE] PORTAL {entrada.name}: {erro}')
        if entrada.name not in retorno:
            retorno.append(entrada.name)
            ATIVIDADE_POR_REQUISICAO[entrada.name] = 'Nota de Credito'
            print(f'[NC RETOMADA] PORTAL {entrada.name} incluída na seleção por ter uma etapa pendente.')
    return retorno


def validar_documentos_conciliacao_nc(estado):
    """Exige validação anterior à VF01, resultado SAP associado e arquivos intactos."""
    dados, consulta = estado['Validacao'], estado['ConsultaSAP']
    if not str(dados.get('ResultadoVA03', '')).startswith('APROVADA antes da VF01.'):
        raise ConciliacaoNCError('Não há confirmação da validação anterior ao faturamento.')
    if consulta.get('StatusSAP') != 'OK_VA03' or consulta.get('ValidacaoNC') != dados['ResultadoVA03']:
        raise ConciliacaoNCError('SAP/PDF ou validação da OV ainda não confirmados.')
    validar_associacao_resultado_sap(estado['Requisicao'], dados['Ordem'], consulta)
    documento = str(consulta.get('NumeroNF', '')).strip()
    if not re.fullmatch(r'\d{8,10}', documento) or not int(documento):
        raise ConciliacaoNCError('Número de faturamento ausente ou inválido.')
    vf01 = str(consulta.get('DocumentoFaturamentoVF01', '')).strip()
    if vf01 and (not vf01.isdigit() or vf01.zfill(10) != documento.zfill(10)):
        raise ConciliacaoNCError('Documento da VF01 difere do documento confirmado na VA03.')
    planilha = dados['CaminhoPlanilha']
    if hash_arquivo_nc(planilha) != estado['HashPlanilha']:
        raise ConciliacaoNCError('A planilha mudou depois da validação. Conciliação não enviada.')
    if hash_arquivo_nc(consulta['ArquivoPDF']) != estado['HashPDF']:
        raise ConciliacaoNCError('O PDF mudou depois da confirmação SAP.')
    return documento.lstrip('0'), planilha


def montar_solicitacao_conciliacao_nc(estado):
    documento, planilha = validar_documentos_conciliacao_nc(estado)
    dados = estado['Validacao']
    req = estado['Requisicao']
    return {
        'Assunto': f'Requisição {req} - Conciliação',
        'Descricao': (
            'Prezados, favor seguir com a conciliação da nota de crédito faturada.\n\n'
            f'Requisição de origem: {req}\nDocumento de faturamento: {documento}\n'
            f'OV: {dados["Ordem"]}\nCliente: {dados["Cliente"]}\n'
            f'Empresa: {dados["Empresa"]}\nDivisão/Centro: {dados["Divisao"]}\n'
            f'Montante validado: {_moeda_credito(dados["Montante"])}\n\n'
            'Planilha de verificação original em anexo, sem alterações.\n\n'
            'Descrição da requisição de origem:\n' + estado.get('Descricao', '')
        ),
        'Centro': dados['Divisao'], 'Planilha': planilha,
    }


def resolver_centro_conciliacao_nc(opcoes, divisao):
    """O value pode ser um ID interno. Prefere o código exibido no rótulo."""
    codigo = _codigo_credito(divisao, 'Divisão').lstrip('0') or '0'
    disponiveis = [o for o in opcoes if not o.get('disabled') and str(o.get('value', '')).strip()]
    por_rotulo = [o for o in disponiveis if codigo in
                  [(n.lstrip('0') or '0') for n in re.findall(r'(?<!\d)\d+(?!\d)', o.get('text', ''))]]
    if len(por_rotulo) == 1:
        return str(por_rotulo[0]['value'])
    if len(por_rotulo) > 1:
        raise ConciliacaoNCError(f'Centro {divisao} aparece em mais de uma opção.')
    por_valor = [o for o in disponiveis if str(o['value']).lstrip('0') == codigo
                 and not re.search(r'\d', o.get('text', ''))]
    if len(por_valor) != 1:
        raise ConciliacaoNCError(f'Divisão {divisao} sem correspondência inequívoca no campo Centro.')
    return str(por_valor[0]['value'])


def ler_controles_portal_nc(page):
    return page.locator('body').evaluate(r'''body => {
        const vis = e => !!e.getClientRects().length && getComputedStyle(e).visibility !== 'hidden';
        const nav = body.querySelector('#btnNavegacao');
        return {
            texto: body.innerText || '',
            links: [...body.querySelectorAll('a')].filter(vis).map(a => ({texto: a.innerText.trim(), href: a.getAttribute('href') || ''})),
            botao: nav && vis(nav) ? (nav.innerText || nav.value || nav.title || '').trim() : '',
            campos: [...body.querySelectorAll('input:not([type=hidden]), textarea, select')].filter(vis)
                .map(e => [e.id || e.name, e.type, e.disabled]),
            erros: [...body.querySelectorAll('.field-validation-error, .validation-summary-errors, .alert-danger, .dz-error-message')]
                .filter(vis).map(e => e.innerText.trim()).filter(Boolean),
            uploads: [...body.querySelectorAll('tr, li, .dz-preview, .file-preview-frame, .uploadify-queue-item')]
                .filter(vis).map(e => ({texto: e.innerText || '',
                    concluido: e.matches('.dz-success, .template-download, .file-upload-success, .upload-complete') ||
                        !!e.querySelector('.dz-success, .file-upload-success, .upload-complete') ||
                        /(?:\bcarregado\s*!|\bupload conclu[ií]do\b|\barquivo (?:enviado|carregado) com sucesso\b)/i.test(e.innerText || '')})),
            pendente: [...body.querySelectorAll('[role=progressbar], .dz-processing:not(.dz-success), .file-uploading')]
                .filter(vis).some(e => e.getAttribute('aria-valuenow') !== '100')
        };
    }''')


def separar_avisos_upload_portal_nc(tela):
    """Só a mensagem de espera observada é transitória; outros erros bloqueiam."""
    esperas, erros = [], []
    for erro in tela.get('erros', []):
        texto = re.sub(r'\s+', ' ', normalizar_texto_busca(erro)).strip()
        if re.fullmatch(r'aguarde enquanto (?:os )?anexos sao carregados[.!]*', texto):
            esperas.append(erro)
        else:
            erros.append(erro)
    return esperas, erros


class MonitorUploadPortalNC:
    """Observa a transferência deste arquivo, sem interceptar/repetir requests.

    input.files prova seleção local, não conclusão do upload. Uma transferência
    HTTP finalizada é evidência adicional aos indicadores do componente.
    URLs, corpos e dados de autenticação não são registrados no log.
    """
    def __init__(self, page, caminho):
        self.page = page
        self.nome = os.path.basename(caminho)
        self.ativos = set()
        self.conhecidos = set()
        self.concluidos = 0
        self.erros = []
        self.handlers = [('request', self.iniciar), ('requestfinished', self.finalizar),
                         ('requestfailed', self.falhar)]
        for evento, handler in self.handlers:
            page.on(evento, handler)

    def iniciar(self, request):
        try:
            anterior = request.redirected_from
            corresponde = anterior in self.conhecidos if anterior is not None else False
            if not corresponde and request.method.upper() in {'POST', 'PUT', 'PATCH'}:
                corpo = request.post_data_buffer or b''
                # Marcador multipart exato: não confunde o nome citado numa
                # descrição de chamado com um arquivo efetivamente enviado.
                marcador = b'filename="' + self.nome.encode('utf-8') + b'"'
                corresponde = marcador in corpo
            if corresponde:
                self.conhecidos.add(request)
                self.ativos.add(request)
        except Exception:
            # Se o browser não expuser o multipart, só a confirmação visual
            # do componente poderá liberar o avanço.
            return

    def finalizar(self, request):
        if request not in self.conhecidos:
            return
        self.ativos.discard(request)
        try:
            resposta = request.response()
            if resposta is None:
                self.erros.append('Upload sem resposta HTTP confirmada.')
                return
            if 300 <= resposta.status < 400:
                return  # A resposta final do redirecionamento ainda é necessária.
            if not 200 <= resposta.status < 300:
                self.erros.append(f'Upload recusado pelo servidor (HTTP {resposta.status}).')
                return
            if 'json' in resposta.headers.get('content-type', '').lower():
                dados = resposta.json()
                itens = [dados] if isinstance(dados, dict) else dados if isinstance(dados, list) else []
                if isinstance(dados, dict) and isinstance(dados.get('files'), list):
                    itens += dados['files']
                if any(isinstance(d, dict) and (d.get('error') or d.get('errors') or
                       d.get('success') is False or d.get('sucesso') is False) for d in itens):
                    self.erros.append('O servidor informou falha ao receber a planilha.')
                    return
            self.concluidos += 1
        except Exception:
            self.erros.append('Não foi possível confirmar a resposta da transferência da planilha.')

    def falhar(self, request):
        if request in self.conhecidos:
            self.ativos.discard(request)
            self.erros.append('A transferência da planilha falhou antes da resposta do servidor.')

    def fechar(self):
        for evento, handler in self.handlers:
            self.page.remove_listener(evento, handler)


def aguardar_upload_planilha_portal_nc(page, campo, caminho, monitor, timeout=90000):
    nome = os.path.basename(caminho)
    arquivo = {'nome': nome, 'tamanho': os.path.getsize(caminho)}
    limite = time.monotonic() + timeout / 1000
    estavel_desde = None
    ultimo_log = -10
    inicio = time.monotonic()
    while time.monotonic() < limite:
        tela = ler_controles_portal_nc(page)
        esperas, erros = separar_avisos_upload_portal_nc(tela)
        if erros or monitor.erros:
            raise ConciliacaoNCError('Upload: ' + ' | '.join(erros + monitor.erros))
        try:
            selecionados = campo.evaluate('el => [...(el.files || [])].map(f => ({nome: f.name, tamanho: f.size}))')
        except PlaywrightError:
            selecionados = []  # O componente pode remover o input após transferir.
        if selecionados and selecionados != [arquivo]:
            raise ConciliacaoNCError('O seletor contém arquivo diferente da planilha de verificação.')
        visual = any(u.get('concluido') and nome in u.get('texto', '') for u in tela.get('uploads', []))
        identificado = selecionados == [arquivo] or nome in tela.get('texto', '')
        transferido = visual or monitor.concluidos > 0
        pronto = (identificado and transferido and not monitor.ativos
                  and not tela.get('pendente') and not esperas)
        agora = time.monotonic()
        if pronto:
            if estavel_desde is None:
                estavel_desde = agora
            if agora - estavel_desde >= 2:
                print(f'[NC CONCILIAÇÃO] Upload concluído e estável: {nome}. Avanço liberado; '
                      'a planilha gravada será conferida por download.')
                return
        else:
            estavel_desde = None
        if agora - inicio - ultimo_log >= 10:
            ultimo_log = agora - inicio
            print(f'[NC CONCILIAÇÃO] Aguardando concluir o upload da planilha '
                  f'({int(agora - inicio)}s/{timeout // 1000}s); avanço ainda bloqueado.')
        page.wait_for_timeout(300)
    raise ConciliacaoNCError('Upload da planilha não confirmado no prazo; nenhum avanço/envio foi acionado. '
                           'Conciliação pendente; a retomada usará o faturamento já registrado.')


def numero_conciliacao_criada_nc(tela, origem):
    texto = normalizar_texto_busca(tela.get('texto', ''))
    esperas, erros = separar_avisos_upload_portal_nc(tela)
    if erros:
        raise ConciliacaoNCError('Portal: ' + ' | '.join(erros))
    if esperas:
        return ''
    if re.search(r'\b(?:nao (?:foi possivel|foi|pode)|erro|falha)\b', texto):
        # O formulário pode conter a descrição original com a palavra erro;
        # isso apenas impede considerar uma página ambígua como confirmação.
        return ''
    sucesso = re.search(r'requisicao[\s\S]{0,100}(?:criada|registrada|aberta|gerada|cadastrada|sucesso)', texto)
    sucesso = sucesso or re.search(r'(?:sucesso|criada|registrada)[\s\S]{0,100}requisicao', texto)
    if not sucesso:
        return ''
    # TESTE 63: o número pode estar no recibo sem ser um link. Exige rótulo
    # explícito de requisição; BP, OV e faturamento não são candidatos.
    numeros = set()
    for padrao in (
        r'\brequisicao\s*(?:n[º°o.]?\s*)?[:#-]?\s*(\d{6,10})\s*[-:,.]?\s*(?:(?:foi|foi\s+criada)\s+)?(?:criada|registrada|aberta|gerada|cadastrada|com\s+sucesso)',
        r'\brequisicao\s+(?:criada|registrada|aberta|gerada|cadastrada)\s+(?:com\s+sucesso[.!]?\s*)?(?:sob\s+(?:o\s+)?)?(?:numero|n[º°o.]?)\s*[:#-]?\s*(\d{6,10})(?!\d)',
        r'\b(?:numero|n[º°o.]?)\s+da\s+requisicao\s*[:#-]?\s*(\d{6,10})(?!\d)',
    ):
        numeros.update(re.findall(padrao, texto))
    # Só o recibo genérico sem ID explícito usa o único link numérico da tela.
    if not numeros:
        numeros = {str(link.get('texto', '')).strip() for link in tela.get('links', [])
                   if re.fullmatch(r'\d{6,10}', str(link.get('texto', '')).strip())}
    # Número igual à origem, inclusive com zeros à esquerda, nunca é válido.
    if any(not int(n) or n.lstrip('0') == str(origem).lstrip('0') for n in numeros):
        return ''
    if len(numeros) != 1:
        return ''
    return next(iter(numeros))


def assinatura_etapa_portal_nc(tela):
    # Não usa relógios, texto da descrição ou indicadores voláteis como prova de avanço.
    return (tela.get('botao', ''), json.dumps(tela.get('campos', []), ensure_ascii=False))


def selecionar_planilha_portal_nc(page, caminho):
    """Seletor gravado pelo usuário; escolha única, sem arrastar e sem reenvio."""
    botao = page.get_by_role('button', name='Selecionar arquivos', exact=True)
    campo = None
    if botao.count() == 1 and botao.evaluate("el => el.tagName === 'INPUT' && el.type === 'file'"):
        campo = botao
    else:
        inputs = page.locator('input[type="file"]')
        habilitados = [inputs.nth(i) for i in range(inputs.count()) if inputs.nth(i).is_enabled()]
        if len(habilitados) == 1:
            campo = habilitados[0]
        elif len(habilitados) > 1:
            raise ConciliacaoNCError('Mais de um seletor de arquivo no formulário da conciliação.')
        else:
            if botao.count() != 1 or not botao.is_visible():
                raise ConciliacaoNCError('Selecionar arquivos não localizado no formulário.')
            with page.expect_file_chooser(timeout=15000) as escolha:
                botao.click(timeout=10000)
            campo = escolha.value.element
    monitor = MonitorUploadPortalNC(page, caminho)
    try:
        campo.set_input_files(caminho, timeout=20000)
        aguardar_upload_planilha_portal_nc(page, campo, caminho, monitor)
    finally:
        monitor.fechar()


def confirmar_planilha_no_formulario_nc(page, pedido):
    nome = os.path.basename(pedido['Planilha'])
    tela = ler_controles_portal_nc(page)
    arquivos = page.locator('input[type="file"]').evaluate_all(
        "els => els.flatMap(e => [...(e.files || [])].map(f => ({nome: f.name, tamanho: f.size})))")
    esperado = {'nome': nome, 'tamanho': os.path.getsize(pedido['Planilha'])}
    if tela['erros'] or tela['pendente']:
        raise ConciliacaoNCError('O formulário indica erro ou upload ainda em andamento.')
    if arquivos != [esperado] and nome not in tela['texto']:
        raise ConciliacaoNCError('A planilha não consta na etapa atual do formulário. Avanço bloqueado.')


def configurar_dialogs_conciliacao_nc(page):
    """Handler só da aba nova; não muda a política dos fluxos já aprovados."""
    eventos = []
    def tratar(dialog):
        texto = normalizar_texto_busca(dialog.message or '')
        conhecido = bool(re.fullmatch(
            r'(?:(?:voce deseja|deseja|confirma|tem certeza que deseja)\s+)?'
            r'(?:realmente\s+)?(?:enviar|criar|registrar|abrir|concluir|salvar)'
            r'(?:\s+(?:esta|a|essa))?\s+(?:requisicao|solicitacao)\s*[?.!]?\s*', texto))
        aceito = dialog.type == 'alert' or (dialog.type == 'confirm' and conhecido)
        eventos.append({'Mensagem': dialog.message, 'Aceito': aceito})
        dialog.accept() if aceito else dialog.dismiss()
    page.on('dialog', tratar)
    return eventos


def abrir_servico_glosa_nc(portal):
    """Segue o menu gravado, sem reaproveitar o ID de outro serviço."""
    portal.goto(PORTAL_URL_MENU,
                wait_until='domcontentloaded', timeout=60000)
    categoria = portal.locator(PORTAL_SELETOR_CATEGORIA)
    categoria.wait_for(state='visible', timeout=30000)
    categoria.click(timeout=15000)
    servico = portal.get_by_role('link', name=SERVICO_CONCILIACAO, exact=True)
    servico.wait_for(state='visible', timeout=15000)
    servico.click(timeout=15000)
    portal.locator(PORTAL_SELETOR_ASSUNTO).wait_for(state='visible', timeout=30000)
    texto = normalizar_texto_busca(portal.locator('body').inner_text())
    texto = re.sub(r'\s+', ' ', texto).replace('–', '-').replace('—', '-')
    if normalizar_texto_busca(SERVICO_CONCILIACAO) not in texto:
        raise ConciliacaoNCError('O serviço Conciliação de Crédito não foi confirmado.')
    print('[NC CONCILIAÇÃO] Serviço confirmado: Conciliação de Crédito.')


def selecionar_centro_glosa_nc(portal, divisao):
    """Centro sempre da Divisão validada; suporta select nativo e busca Select2."""
    codigo = _codigo_credito(divisao, 'Divisão')
    id_campo = PORTAL_SELETOR_CENTRO
    centro = portal.locator(id_campo)
    centro.wait_for(state='attached', timeout=15000)

    def opcoes_atuais():
        return centro.locator('option').evaluate_all(
            'els => els.map(e => ({value: e.value, text: e.textContent.trim(), disabled: e.disabled}))')

    if centro.is_visible():
        valor = resolver_centro_conciliacao_nc(opcoes_atuais(), codigo)
        centro.select_option(value=valor)
    else:
        # O Codegen da Glosa abre a seta <b>, pesquisa e clica em treeitem.
        # Limita a abertura ao componente associado ao campo Centro, evitando
        # um clique global em qualquer <b> da página.
        seletores = [
            portal.locator('#select2-' + id_campo + '-container'),
            portal.locator('#s2id_' + id_campo),
            centro.locator("xpath=following-sibling::*[contains(concat(' ',normalize-space(@class),' '),' select2 ')][1]"),
        ]
        abrir = None
        for candidatos in seletores:
            visiveis = [candidatos.nth(i) for i in range(candidatos.count())
                        if candidatos.nth(i).is_visible()]
            if len(visiveis) > 1:
                raise ConciliacaoNCError('Mais de um componente associado ao Centro.')
            if visiveis:
                abrir = visiveis[0]
                break
        if abrir is None:
            raise ConciliacaoNCError('Seletor pesquisável do Centro não localizado junto ao campo mapeado.')
        abrir.click(timeout=15000)
        pesquisas = portal.locator('input[type="search"]:visible')
        pesquisas.wait_for(state='visible', timeout=10000)
        pesquisas.fill(codigo, timeout=10000)
        limite = time.monotonic() + 15
        escolhido = None
        padrao = re.compile(rf'^\s*0*{re.escape(codigo)}(?!\d)(?:\s|$|[-–—:])', re.I)
        while time.monotonic() < limite:
            for papel in ('treeitem', 'option'):
                candidatos = portal.get_by_role(papel, name=padrao)
                visiveis = [candidatos.nth(i) for i in range(candidatos.count())
                            if candidatos.nth(i).is_visible()]
                if len(visiveis) > 1:
                    raise ConciliacaoNCError(f'Centro {codigo} aparece em mais de uma opção de pesquisa.')
                if visiveis:
                    escolhido = visiveis[0]
                    break
            if escolhido is not None:
                break
            portal.wait_for_timeout(300)
        if escolhido is None:
            raise ConciliacaoNCError(f'Divisão {codigo} não encontrada na pesquisa do Centro.')
        escolhido.click(timeout=15000)
        # A busca pode carregar opções remotamente. Só depois da escolha lemos
        # o value interno e o rótulo efetivamente selecionado pelo componente.
        valor = resolver_centro_conciliacao_nc(opcoes_atuais(), codigo)

    if centro.input_value() != valor:
        raise ConciliacaoNCError(f'O Centro não conservou a Divisão {codigo} selecionada.')
    print(f'[NC CONCILIAÇÃO] Centro confirmado pela Divisão da planilha: {codigo}.')
    return valor


def enviar_conciliacao_portal_nc(page_portal, pasta, estado):
    pedido = montar_solicitacao_conciliacao_nc(estado)
    portal = page_portal.context.new_page()
    globals().setdefault('PORTAIS_ENVIO_NC', {})[estado['Requisicao']] = portal
    eventos = configurar_dialogs_conciliacao_nc(portal)
    abrir_servico_glosa_nc(portal)
    assunto = portal.locator(PORTAL_SELETOR_ASSUNTO)
    descricao = portal.locator(PORTAL_SELETOR_DESCRICAO)
    centro = portal.locator(PORTAL_SELETOR_CENTRO)
    assunto.wait_for(state='visible', timeout=30000)
    assunto.fill(pedido['Assunto'])
    descricao.fill(pedido['Descricao'])
    valor_centro = selecionar_centro_glosa_nc(portal, pedido['Centro'])
    if (assunto.input_value() != pedido['Assunto'] or descricao.input_value() != pedido['Descricao']
            or centro.input_value() != valor_centro):
        raise ConciliacaoNCError('Os campos do formulário não conservaram os valores preenchidos.')
    selecionar_planilha_portal_nc(portal, pedido['Planilha'])
    validar_documentos_conciliacao_nc(estado)
    salvar_controle_nc(pasta, estado, Pedido=pedido, ServicoConciliacao=SERVICO_CONCILIACAO)
    # Pode concluir no primeiro avanço, como na gravação da Glosa. O segundo
    # só ocorre se houver outra etapa confirmada; não há clique duplo de envio.
    # O registro anterior ao clique distingue envios novos dos controles legados.
    for etapa in (1, 2):
        tela = ler_controles_portal_nc(portal)
        if tela['erros'] or tela['pendente']:
            raise ConciliacaoNCError('Há erro de formulário ou transferência de arquivo em andamento.')
        confirmar_planilha_no_formulario_nc(portal, pedido)
        validar_documentos_conciliacao_nc(estado)
        antes = assinatura_etapa_portal_nc(tela)
        botao = portal.locator('#btnNavegacao')
        if not botao.is_visible() or not botao.is_enabled():
            raise ConciliacaoNCError('Botão de avanço/envio indisponível.')
        instante = datetime.now().isoformat(timespec='microseconds')
        alteracoes = {'Fase': 'ENVIO_INICIADO', 'EtapaPortal': etapa,
                      'EnvioIniciadoEm': instante, 'UltimoRetornoEnvio': ''}
        if etapa == 1:
            alteracoes['TentativaEnvioConciliacao'] = {
                'Versao': 1, 'IniciadaEm': instante, 'Rodada': globals().get('ID_RODADA', ''),
                'RequisicaoOrigem': estado['Requisicao'],
                'HashPlanilha': estado['HashPlanilha'], 'HashPDF': estado['HashPDF']}
            recuperacao = estado.get('RecuperacaoLegadaSemNumero')
            if recuperacao and recuperacao.get('Situacao') == 'PREPARADA':
                alteracoes['RecuperacaoLegadaSemNumero'] = {
                    **recuperacao, 'Situacao': 'ENVIO_INICIADO', 'EnvioIniciadoEm': instante}
        # Persistência atômica antes do clique: queda do processo não permite
        # reutilizar a exceção legada em uma execução posterior.
        salvar_controle_nc(pasta, estado, **alteracoes)
        print(f'[NC CONCILIAÇÃO] Avanço {etapa}/2: {tela["botao"]}.')
        botao.click(timeout=15000)
        limite = time.monotonic() + 90
        assinatura_estavel = None
        leituras = 0
        espera_upload_registrada = False
        while time.monotonic() < limite:
            if any(not e['Aceito'] for e in eventos):
                raise ConciliacaoNCError('Confirmação desconhecida no portal; conferir a tela antes de reenviar.')
            tela = ler_controles_portal_nc(portal)
            numero = numero_conciliacao_criada_nc(tela, estado['Requisicao'])
            if numero:
                salvar_controle_nc(pasta, estado, Fase='CRIADA_ANEXO_PENDENTE', RequisicaoConciliacao=numero)
                print(f'[NC CONCILIAÇÃO] PORTAL {numero} criada; conferindo a planilha gravada.')
                portal.close()
                globals().get('PORTAIS_ENVIO_NC', {}).pop(estado['Requisicao'], None)
                return numero
            esperas, _erros = separar_avisos_upload_portal_nc(tela)
            if esperas or tela.get('pendente'):
                if not espera_upload_registrada:
                    print('[NC CONCILIAÇÃO] Portal ainda aguarda os anexos; acompanhando a resposta sem repetir o clique.')
                    espera_upload_registrada = True
                leituras = 0
                portal.wait_for_timeout(400)
                continue
            atual = assinatura_etapa_portal_nc(tela)
            if atual != antes and tela['botao'] and not tela['pendente']:
                leituras = leituras + 1 if atual == assinatura_estavel else 1
                assinatura_estavel = atual
                if leituras >= 3:
                    break
            else:
                leituras = 0
            portal.wait_for_timeout(400)
        else:
            motivo = 'Portal continuou aguardando os anexos. ' if espera_upload_registrada else ''
            raise ConciliacaoNCError(motivo + 'Envio/avanço sem confirmação. '
                                    'Não reenviar antes de conferir se a conciliação foi criada.')
    raise ConciliacaoNCError('Dois avanços executados sem número confirmado. Conferir o PORTAL; não haverá reenvio automático.')


def confirmar_anexo_gravado_nc(page, frame, caminho, hash_esperado, pasta):
    """Nome exato + download + SHA-256, tanto para o PDF quanto para o Excel."""
    portalmain = frame_portalmain(frame)
    portalmain.get_by_text('Documentos', exact=True).click(timeout=15000)
    andoc = portalmain.frame_locator(FRAME_DOCUMENTOS)
    andoc.locator('body').wait_for(state='visible', timeout=20000)
    nome = os.path.basename(caminho)
    limite = time.monotonic() + 20
    indices = []
    while time.monotonic() < limite:
        indices = andoc.locator('body').evaluate(r'''(body, nome) => {
            const norm = s => String(s || '').trim().toLowerCase();
            const vis = e => !!e.getClientRects().length && getComputedStyle(e).visibility !== 'hidden';
            return [...body.querySelectorAll('a')].map((e, i) => ({e, i})).filter(({e}) =>
                vis(e) && [e.innerText, e.title].some(s => norm(s) === norm(nome)) &&
                ![...(e.closest('tr') || e).querySelectorAll('input[type=text], textarea')].some(vis)
            ).map(({i}) => i);
        }''', nome)
        if indices:
            break
        page.wait_for_timeout(400)
    if len(indices) != 1:
        raise ConciliacaoNCError(f'Anexo {nome}: esperado um link gravado; encontrados {len(indices)}.')
    with page.expect_download(timeout=30000) as transferencia:
        andoc.locator('a').nth(indices[0]).click(timeout=15000)
    download = transferencia.value
    if download.failure():
        raise ConciliacaoNCError('Não foi possível baixar o anexo para confirmar a gravação.')
    fd, temporario = tempfile.mkstemp(prefix='conferencia_nc_', suffix=os.path.splitext(nome)[1], dir=pasta)
    os.close(fd)
    try:
        download.save_as(temporario)
        if hash_arquivo_nc(temporario) != hash_esperado:
            raise ConciliacaoNCError(f'O conteúdo do anexo {nome} difere do arquivo validado.')
    finally:
        os.remove(temporario)


def confirmar_original_encerrada_nc(page, pasta, estado):
    req = estado['Requisicao']
    frame = abrir_requisicao_portal(page, req)
    status = confirmar_status_cabecalho_portal(page, frame, req)
    if normalizar_texto_busca(status).strip() != 'encerrada':
        raise ConciliacaoNCError(f'A requisição original não está encerrada: {status}. Conciliação não enviada.')
    consulta = estado['ConsultaSAP']
    pdf = preparar_pdf_credito_para_portal(req, consulta, consulta['ArquivoPDF'])
    confirmar_anexo_gravado_nc(page, frame, pdf, estado['HashPDF'], pasta)
    salvar_controle_nc(pasta, estado, OriginalEncerrada=True,
                      ResultadoPORTAL={**(estado.get('ResultadoPORTAL') or {}), 'StatusPORTAL': 'CONCLUIDO', 'QtdPDFsAnexados': 1,
                                    'MensagemPORTAL': 'Encerramento e PDF da requisição original conferidos.'})


def confirmar_conciliacao_gravada_nc(page, pasta, estado):
    numero = estado['RequisicaoConciliacao']
    frame = abrir_requisicao_portal(page, numero)
    validar_requisicao_portal_aberta(frame, numero, timeout=15000)
    descricao = frame_portalmain(frame).frame_locator(
        FRAME_EDITOR_TEXTO).locator('body').inner_text(timeout=15000)
    norm = lambda s: re.sub(r'\s+', ' ', str(s)).strip()
    if norm(estado['Pedido']['Descricao']) not in norm(descricao):
        raise ConciliacaoNCError(f'PORTAL {numero} criada, mas a descrição gravada não corresponde ao pedido.')
    confirmar_anexo_gravado_nc(page, frame, estado['Validacao']['CaminhoPlanilha'], estado['HashPlanilha'], pasta)
    salvar_controle_nc(pasta, estado, Fase='CONCLUIDA', MensagemConciliacao='Nova requisição e planilha original confirmadas.')


def resultado_controle_nc(estado, retomada=False):
    resultado = dict(estado.get('ResultadoPORTAL') or {
        'StatusPORTAL': 'PENDENTE_NC', 'QtdPDFsAnexados': 0,
        'MensagemPORTAL': 'Etapa pendente; consultar a mensagem da conciliação.'})
    resultado.update(
        OrigemOVs='PORTAL_NC_DUAS_PLANILHAS',
        StatusConciliacao=estado['Fase'],
        RequisicaoConciliacao=estado.get('RequisicaoConciliacao', ''),
        MensagemConciliacao=estado.get('MensagemConciliacao', ''),
        PlanilhaConciliacao=estado['Validacao'].get('CaminhoPlanilha', ''),
        ValidacaoNC=(estado['Validacao'].get('ResultadoVA03') or
                     (estado.get('ConsultaSAP') or {}).get('ValidacaoNC', '')), RetomadaNC=retomada,
        Etapa='NC_' + estado['Fase'],
    )
    if estado['Fase'] in {'SAP_EM_EXECUCAO', 'SAP_PENDENTE'}:
        consulta = estado.get('ConsultaSAP') or {}
        status_sap = consulta.get('StatusSAP') or 'SEM_RESULTADO_REGISTRADO'
        documento = consulta.get('DocumentoFaturamentoVF01') or consulta.get('NumeroNF') or ''
        detalhe = (f'OV {estado["Validacao"].get("Ordem", "")}: SAP {status_sap}'
                   + (f'; faturamento {documento}' if documento else '; sem faturamento confirmado')
                   + '. ' + str(consulta.get('Mensagem') or 'Execução interrompida antes de registrar o resultado SAP.'))
        resultado['MensagemPORTAL'] = ('Conciliação não iniciada: etapa SAP/PDF pendente. ' + detalhe)
        resultado['MensagemConciliacao'] = (str(resultado.get('MensagemConciliacao') or '')
                                           + ' | ' + detalhe).lstrip(' |')
    if resultado.get('StatusPORTAL') == 'CONCLUIDO':
        numero = estado.get('RequisicaoConciliacao', '')
        if estado['Fase'] == 'CONCLUIDA':
            mensagem = f'Conciliação {numero} criada e conferida com a planilha original.'
        else:
            mensagem = ('PORTAL original encerrado; processo ainda pendente da conciliação '
                        f'({estado["Fase"]}' + (f'; PORTAL {numero}' if numero else '; sem número confirmado') + '). '
                        + str(estado.get('MensagemConciliacao', '')))
        resultado['MensagemPORTAL'] = str(resultado.get('MensagemPORTAL', '')) + ' | ' + mensagem
    return resultado


def conferir_vinculo_conciliacao_existente_nc(page, estado, numero, pedido):
    """Consulta sem escrita: número, serviço Glosa e descrição do pedido atual."""
    if (not re.fullmatch(r'\d{6,10}', numero) or not int(numero)
            or numero.lstrip('0') == estado['Requisicao'].lstrip('0')):
        raise ConciliacaoNCError('Número de conciliação inválido ou igual ao PORTAL original; vínculo não alterado.')
    frame = abrir_requisicao_portal(page, numero)
    validar_requisicao_portal_aberta(frame, numero, timeout=15000)
    portalmain = frame_portalmain(frame)
    texto_servico = normalizar_texto_busca(portalmain.locator('body').inner_text(timeout=15000))
    servicos = re.findall(r'\bconciliacao\s*(?:[-–—]\s*)?credito(?:\s+(glosa|autorizado))?\b', texto_servico)
    # Rejeitar uma requisição com outro tipo de crédito (ex.: autorizado) ou ambígua.
    if not servicos or 'autorizado' in servicos or len(servicos) != 1:
        raise ConciliacaoNCError(f'PORTAL {numero}: não foi possível confirmar o serviço Conciliação de Crédito. '
                                'Vínculo não alterado; conferir o cabeçalho do chamado.')
    descricao = portalmain.frame_locator(
        FRAME_EDITOR_TEXTO).locator('body').inner_text(timeout=15000)
    norm = lambda s: re.sub(r'\s+', ' ', str(s)).strip()
    if norm(pedido['Descricao']) not in norm(descricao):
        raise ConciliacaoNCError(f'PORTAL {numero}: descrição incompatível com a origem, OV, faturamento e demais '
                                'dados do pedido registrado. Vínculo não alterado.')


def localizar_controle_consulta_nc(page, papel, nome):
    """Localiza apenas controles visíveis pelo rótulo; não inventa endpoints."""
    encontrados = []
    for frame in page.frames:
        controles = frame.get_by_role(papel, name=nome, exact=True)
        for i in range(controles.count()):
            controle = controles.nth(i)
            if controle.is_visible():
                encontrados.append(controle)
    if len(encontrados) != 1:
        raise ConciliacaoNCError(f'Consulta automática: esperado um controle {papel} {nome}; '
                                f'encontrados {len(encontrados)}.')
    return encontrados[0]


def preparar_lista_consulta_conciliacao_nc(consulta):
    """Abre a lista do solicitante e inclui encerradas usando os rótulos da UI.

    O layout desta consulta deve ser reconhecido em execução. Tela desconhecida,
    filtros não verificáveis e falta de total NUNCA significam lista vazia.
    """
    consulta.goto(PORTAL_URL_MENU,
                  wait_until='domcontentloaded', timeout=30000)
    nome = re.compile(r'^\s*(?:Minhas (?:requisições|solicitações)|'
                      r'Consultar requisições|Consulta de requisições|Acompanhar requisições)\s*$', re.I)
    limite = time.monotonic() + 12
    while True:
        try:
            link = localizar_controle_consulta_nc(consulta, 'link', nome)
            break
        except ConciliacaoNCError:
            if time.monotonic() >= limite:
                raise ConciliacaoNCError('Consulta automática: menu de requisições do solicitante '
                                        'não reconhecido. Não foi possível verificar envios anteriores.')
            consulta.wait_for_timeout(300)
    link.click(timeout=10000)
    limite = time.monotonic() + 12
    while time.monotonic() < limite:
        # Status é escolhido explicitamente. Não usa só a fila de chamados abertos.
        for frame in consulta.frames:
            selects = frame.locator('select:visible')
            for i in range(selects.count()):
                campo = selects.nth(i)
                meta = campo.evaluate(r'''e => ({rotulo: [e.getAttribute('aria-label'), e.name, e.id,
                    ...[...(e.labels || [])].map(l => l.innerText)].filter(Boolean).join(' '),
                    opcoes: [...e.options].map(o => ({value:o.value, texto:o.text, disabled:o.disabled}))})''')
                if not re.search(r'status|situacao', normalizar_texto_busca(meta['rotulo'])):
                    continue
                todas = [o for o in meta['opcoes'] if not o['disabled'] and re.fullmatch(
                    r'(?:todos|todas)(?: os status| as situacoes| as requisicoes)?',
                    normalizar_texto_busca(o['texto']).strip())]
                if len(todas) != 1:
                    raise ConciliacaoNCError('Consulta automática: opção para incluir todos os status não confirmada.')
                campo.select_option(value=todas[0]['value'])
                if campo.input_value() != todas[0]['value']:
                    raise ConciliacaoNCError('Consulta automática: filtro de status não conservou Todos.')
                # Não limpa nem interpreta filtros desconhecidos. Se houver algum,
                # o resultado não é suficiente para provar inexistência.
                filtros = frame.locator('input:visible, select:visible, textarea:visible').evaluate_all(r'''els =>
                    els.filter(e => !['button','submit','reset','file','hidden'].includes(e.type))
                    .map(e => ({tag:e.tagName, tipo:e.type, valor:e.value || '',
                        rotulo:[e.getAttribute('aria-label'), e.placeholder, e.id, e.name,
                            ...[...(e.labels || [])].map(l=>l.innerText)].filter(Boolean).join(' '),
                        texto:e.tagName==='SELECT' ? [...e.selectedOptions].map(o=>o.text).join(' ') : '',
                        marcado:!!e.checked}))''')
                for filtro in filtros:
                    rotulo = normalizar_texto_busca(filtro['rotulo'])
                    if re.search(r'status|situacao', rotulo):
                        continue
                    if filtro['tag'] == 'SELECT':
                        # Tamanho da página não restringe o universo; será paginado.
                        if re.search(r'length|por pagina|registros|exibir|mostrar', rotulo):
                            continue
                        texto = normalizar_texto_busca(filtro['texto']).strip()
                        if filtro['valor'] and not re.fullmatch(r'(?:todos|todas)(?: .*)?', texto):
                            raise ConciliacaoNCError('Consulta automática: há outro filtro ativo na listagem.')
                    elif filtro['tipo'] in ('checkbox', 'radio'):
                        if filtro['marcado']:
                            raise ConciliacaoNCError('Consulta automática: opção marcada restringe ou torna incerto o alcance da lista.')
                    elif filtro['valor'].strip():
                        raise ConciliacaoNCError('Consulta automática: há filtro de texto/data preenchido na listagem.')
                # Alguns layouts aplicam o select imediatamente; outros usam botão.
                botoes = frame.get_by_role('button', name=re.compile(r'^\s*(?:Pesquisar|Buscar|Consultar|Filtrar)\s*$', re.I))
                visiveis = [botoes.nth(j) for j in range(botoes.count()) if botoes.nth(j).is_visible()]
                if len(visiveis) > 1:
                    raise ConciliacaoNCError('Consulta automática: mais de um botão de pesquisa visível.')
                if visiveis:
                    visiveis[0].click(timeout=10000)
                consulta.wait_for_load_state('networkidle', timeout=15000)
                consulta.wait_for_timeout(800)
                return
        consulta.wait_for_timeout(300)
    raise ConciliacaoNCError('Consulta automática: filtro que inclui chamados encerrados não localizado.')


def ler_lista_consulta_conciliacao_nc(consulta):
    """Lê tabelas de requisições e totais explícitos; não usa números do corpo inteiro."""
    telas = []
    for frame in consulta.frames:
        tela = frame.locator('body').evaluate(r'''body => {
            const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
            const vis = e => !!e.getClientRects().length && getComputedStyle(e).visibility !== 'hidden';
            const tabelas = [...body.querySelectorAll('table')].filter(vis).map(t => {
                const trs = [...t.rows].filter(vis);
                const cab = trs.length ? [...trs[0].cells].map(e=>norm(e.innerText)) : [];
                const nr = cab.findIndex(s=>/^(?:n[º°o.]?\s*(?:da\s+)?|numero\s+(?:da\s+)?)?requisicao$/.test(s));
                const assunto = cab.findIndex(s=>s==='assunto');
                if(nr<0 || assunto<0) return null;
                const linhas = trs.slice(1).filter(tr=>tr.querySelector('td')).map(tr=>({
                    numero: (tr.cells[nr]?.innerText || '').trim(),
                    assunto: (tr.cells[assunto]?.innerText || '').trim(),
                    texto: tr.innerText || ''}));
                return {linhas};
            }).filter(Boolean);
            return {tabelas, texto:body.innerText || '',
                ocupado:[...body.querySelectorAll('[aria-busy=true], .dataTables_processing, .loading, .loader')].some(vis),
                erros:[...body.querySelectorAll('.alert-danger,[role=alert].error')].filter(vis).map(e=>e.innerText),
                status:[...body.querySelectorAll('select')].filter(vis)
                    .filter(e=>/status|situacao/.test(norm([e.id,e.name,e.getAttribute('aria-label'),
                        ...[...(e.labels||[])].map(l=>l.innerText)].join(' '))))
                    .map(e=>[...e.selectedOptions].map(o=>o.text).join(' '))};
        }''')
        if tela['tabelas']:
            telas.append(tela)
    if len(telas) != 1 or len(telas[0]['tabelas']) != 1:
        raise ConciliacaoNCError('Consulta automática: tabela de requisições/assuntos não identificada de forma única.')
    tela = telas[0]
    if tela['ocupado'] or tela['erros']:
        raise ConciliacaoNCError('Consulta automática: listagem carregando ou com erro.')
    if not tela['status'] or any(not re.fullmatch(
            r'(?:todos|todas)(?: os status| as situacoes| as requisicoes)?',
            normalizar_texto_busca(s).strip()) for s in tela['status']):
        raise ConciliacaoNCError('Consulta automática: a lista não confirmou todos os status.')
    texto = normalizar_texto_busca(tela['texto'])
    total = None
    for padrao in (r'(?:mostrando|exibindo)\s+(?:de\s+)?[\d.,]+\s+(?:a|ate|-)\s*[\d.,]+\s+de\s+([\d.,]+)\s+(?:registros|requisicoes|resultados)',
                   r'total\s*(?:de\s+)?(?:registros|requisicoes|resultados)\s*:\s*([\d.,]+)'):
        m = re.search(padrao, texto)
        if m:
            total = int(re.sub(r'\D', '', m.group(1)))
            break
    vazia = bool(re.search(r'\b(?:nenhum (?:registro|resultado)|nenhuma requisicao)\s+(?:encontrad[oa]|localizad[oa])', texto))
    linhas, invalidas = [], []
    for linha in tela['tabelas'][0]['linhas']:
        if re.fullmatch(r'\d{6,10}', linha['numero']) and int(linha['numero']):
            linhas.append(linha)
        elif not (vazia and re.search(r'nenhum|nenhuma', normalizar_texto_busca(linha['texto']))):
            invalidas.append(linha)
    if invalidas:
        raise ConciliacaoNCError('Consulta automática: linhas da lista não puderam ser lidas integralmente.')
    # Zero só vale quando acompanhado da mensagem explícita na tabela reconhecida.
    if total == 0 and not vazia:
        raise ConciliacaoNCError('Consulta automática: total zero sem confirmação de lista vazia.')
    if vazia and not linhas and total is None:
        total = 0
    if total is not None and total < len(linhas):
        raise ConciliacaoNCError('Consulta automática: contagem da lista inconsistente.')
    return {'Linhas': linhas, 'Total': total}


def aguardar_lista_consulta_conciliacao_nc(consulta, anterior=None):
    limite, assinatura, estavel = time.monotonic() + 12, None, 0
    ultimo_erro = 'Listagem sem confirmação.'
    while time.monotonic() < limite:
        try:
            tela = ler_lista_consulta_conciliacao_nc(consulta)
            atual = json.dumps(tela, ensure_ascii=False, sort_keys=True)
            if atual != anterior:
                estavel = estavel + 1 if atual == assinatura else 1
                assinatura = atual
                if estavel >= 3:
                    return tela, atual
        except Exception as erro:
            ultimo_erro = str(erro)
            estavel = 0
        consulta.wait_for_timeout(400)
    raise ConciliacaoNCError('Consulta automática: leitura/paginação não estabilizou. ' + ultimo_erro)


def consultar_conciliacoes_existentes_nc(page, estado, pedido):
    """Consulta somente leitura. Ausência exige lista completa com todos os status.

    O assunto é o mesmo que as versões anteriores do robô enviavam. Cada número
    encontrado ainda será aberto e conferido por serviço e descrição completa.
    """
    consulta = None
    candidatos, vistos = set(), set()
    try:
        consulta = page.context.new_page()
        preparar_lista_consulta_conciliacao_nc(consulta)
        total_inicial, anterior = None, None
        prazo = time.monotonic() + 60
        for pagina in range(1, 51):
            if time.monotonic() >= prazo:
                raise ConciliacaoNCError('Consulta automática: limite de tempo da listagem atingido.')
            tela, assinatura = aguardar_lista_consulta_conciliacao_nc(consulta, anterior)
            total = tela['Total']
            if pagina == 1:
                total_inicial = total
            elif total != total_inicial:
                raise ConciliacaoNCError('Consulta automática: total mudou durante a paginação.')
            novos = {linha['numero'] for linha in tela['Linhas']}
            if vistos.intersection(novos) or len(novos) != len(tela['Linhas']):
                raise ConciliacaoNCError('Consulta automática: paginação repetiu requisições.')
            vistos.update(novos)
            for linha in tela['Linhas']:
                assunto = normalizar_texto_busca(linha['assunto']).replace('–', '-').replace('—', '-')
                # Marca também assuntos parcialmente alterados para conferência;
                # evita declarar inexistente um chamado com a mesma origem.
                if (linha['numero'].lstrip('0') != estado['Requisicao'].lstrip('0') and
                        re.search(rf'(?<!\d)0*{re.escape(estado["Requisicao"].lstrip("0"))}(?!\d)', assunto)):
                    candidatos.add(linha['numero'])
            if total is not None and len(vistos) == total:
                return {'Numeros': sorted(candidatos), 'Completa': True, 'TotalLido': len(vistos),
                        'Motivo': 'Todas as páginas conferidas, incluindo encerradas.'}
            if total is not None and len(vistos) > total:
                raise ConciliacaoNCError('Consulta automática: total inferior às linhas consultadas.')
            # Paginação por rótulo visível; falha de navegação não prova fim da lista.
            proxima = localizar_controle_consulta_nc(consulta, 'link', re.compile(r'^\s*(?:Próximo|Próxima|Next)\s*$', re.I))
            if not proxima.is_enabled() or proxima.get_attribute('aria-disabled') == 'true':
                raise ConciliacaoNCError('Consulta automática: fim da lista sem contagem completa confirmada.')
            proxima.click(timeout=10000)
            anterior = assinatura
        raise ConciliacaoNCError('Consulta automática: limite de páginas atingido.')
    except Exception as erro:
        return {'Numeros': sorted(candidatos), 'Completa': False, 'TotalLido': len(vistos), 'Motivo': str(erro),
                'Diagnostico': registrar_diagnostico_consulta_nc(consulta)}
    finally:
        if consulta is not None:
            try:
                consulta.close()
            except Exception:
                pass


def vincular_conciliacao_automatica_nc(page, pasta, estado, pedido, numero, origem):
    conferir_vinculo_conciliacao_existente_nc(page, estado, numero, pedido)
    registro = {'Em': datetime.now().isoformat(timespec='seconds'), 'Acao': 'VINCULO_AUTOMATICO_CONFERIDO',
                'OrigemConferencia': origem, 'RequisicaoConciliacao': numero}
    salvar_controle_nc(pasta, estado, Fase='CRIADA_ANEXO_PENDENTE', Pedido=pedido,
                      RequisicaoConciliacao=numero, ServicoConciliacao=SERVICO_CONCILIACAO,
                      HistoricoRetomadasConciliacao=list(estado.get('HistoricoRetomadasConciliacao') or []) + [registro],
                      MensagemConciliacao=f'Conciliação {numero} vinculada automaticamente; conferindo a planilha.')
    print(f'[NC CONCILIAÇÃO] Origem {estado["Requisicao"]} -> PORTAL {numero}; vínculo automático confirmado.')


def registrar_diagnostico_consulta_nc(consulta):
    """Registra rótulos da interface que faltou mapear; sem valores preenchidos/URLs."""
    telas = []
    if consulta is None:
        return telas
    try:
        for frame in consulta.frames:
            tela = frame.locator('body').evaluate(r'''body => {
                const vis = e => !!e.getClientRects().length && getComputedStyle(e).visibility !== 'hidden';
                const limpa = s => String(s || '').replace(/\s+/g,' ').trim().slice(0,140);
                return {
                    links: [...body.querySelectorAll('a')].filter(vis).map(e=>limpa(e.innerText || e.title))
                        .filter(s=>/requisi|solicita|chamado|consulta|pesquis|acompanha|minhas|meus/i.test(s)).slice(0,30),
                    botoes: [...body.querySelectorAll('button,input[type=button],input[type=submit]')]
                        .filter(vis).map(e=>limpa(e.innerText || e.value || e.title)).slice(0,20),
                    campos: [...body.querySelectorAll('select,input:not([type=hidden])')].filter(vis)
                        .map(e=>({tipo:e.type,rotulo:limpa(e.getAttribute('aria-label') || e.placeholder ||
                            [...(e.labels || [])].map(l=>l.innerText).join(' ') || e.name || e.id)})).slice(0,30),
                    cabecalhos: [...body.querySelectorAll('th')].filter(vis).map(e=>limpa(e.innerText)).slice(0,20)
                };
            }''', timeout=3000)
            if any(tela.values()):
                telas.append(tela)
            if len(telas) >= 5:
                break
        # O diagnóstico completo também fica dentro do controle da requisição.
        print('[NC CONSULTA DIAGNÓSTICO] ' + json.dumps(telas, ensure_ascii=False)[:4500])
    except Exception:
        pass  # Nunca mascara a causa da falha de consulta.
    return telas


def aguardar_confirmacao_envio_nc(page, pasta, estado, pedido, inicio):
    """Dá prazo ao envio atual e relê seu recibo dentro desta mesma execução."""
    try:
        idade = (datetime.now() - datetime.fromisoformat(inicio)).total_seconds()
    except (TypeError, ValueError):
        raise ConciliacaoNCError('Envio anterior sem data válida: não é possível confirmar se ainda está em processamento.')
    if idade < -5:
        raise ConciliacaoNCError('Data do envio está no futuro; conferir o relógio antes de permitir nova criação.')
    espera = max(0, 120 - max(0, idade))
    if espera <= 0:
        return False
    limite = time.monotonic() + espera
    proximo_log = 0
    req = estado['Requisicao']
    while time.monotonic() < limite:
        portal = globals().get('PORTAIS_ENVIO_NC', {}).get(req)
        numero = ''
        if portal is not None:
            try:
                numero = numero_conciliacao_criada_nc(ler_controles_portal_nc(portal), req)
            except Exception:
                pass
        if numero:
            vincular_conciliacao_automatica_nc(page, pasta, estado, pedido, numero, 'RECIBO_NA_MESMA_EXECUCAO')
            try:
                portal.close()
            finally:
                globals().get('PORTAIS_ENVIO_NC', {}).pop(req, None)
            return True
        agora = time.monotonic()
        if agora >= proximo_log:
            print(f'[NC CONCILIAÇÃO] Origem {req}: acompanhando a confirmação nesta execução '
                  f'({max(1, int(limite - agora))}s restantes antes de nova conferência).')
            proximo_log = agora + 15
        page.wait_for_timeout(max(1, min(500, int((limite - agora) * 1000))))
    return False


def concluir_conciliacao_na_mesma_execucao_nc(page, pasta, estado):
    """Conclui o pós-encerramento; uma recuperação e até dois envios por chamada.

    Retomadas exigem ausência confirmada ou a exceção legada única, registrada
    pela diretriz de criação automática. Não repete SAP nem encerra a fila.
    """
    envios = 0
    recuperacoes = 0
    while True:
        fase = estado['Fase']
        if fase == 'CONCLUIDA':
            return
        if fase == 'CRIADA_ANEXO_PENDENTE':
            confirmar_conciliacao_gravada_nc(page, pasta, estado)
            return
        if fase == 'ENVIO_INICIADO':
            if recuperacoes >= 1:
                raise ConciliacaoNCError('A nova tentativa também ficou sem confirmação. '
                                        'Limite de recuperação atingido; resultado preservado para conferência. '
                                        + str(estado.get('UltimoRetornoEnvio') or ''))
            recuperacoes += 1
            resolver_envio_iniciado_nc(page, pasta, estado)
            if estado['Fase'] not in {'ORIGINAL_ENCERRADA', 'CRIADA_ANEXO_PENDENTE', 'CONCLUIDA'}:
                raise ConciliacaoNCError('Recuperação não confirmou a próxima etapa da conciliação.')
            continue
        if fase != 'ORIGINAL_ENCERRADA':
            raise ConciliacaoNCError(f'Etapa {fase} não permite abrir conciliação.')
        if envios >= 2:
            raise ConciliacaoNCError('Limite de tentativas de criação nesta execução atingido.')
        validar_documentos_conciliacao_nc(estado)
        envios += 1
        print(f'[NC CONCILIAÇÃO] PORTAL {estado["Requisicao"]} encerrado e PDF conferido. '
              f'Criando Conciliação de Crédito agora (tentativa {envios}/2).')
        try:
            enviar_conciliacao_portal_nc(page, pasta, estado)
        except Exception as erro:
            salvar_controle_nc(pasta, estado, UltimoRetornoEnvio=str(erro))
            if estado['Fase'] == 'ENVIO_INICIADO':
                # O registro foi gravado antes do clique. Nenhum novo formulário
                # será enviado sem a resolução do envio anterior.
                continue
            if (estado['Fase'] == 'ORIGINAL_ENCERRADA' and envios < 2
                    and isinstance(erro, PlaywrightError)):
                # Falha de navegação anterior ao primeiro clique: sem envio iniciado.
                portal = globals().get('PORTAIS_ENVIO_NC', {}).pop(estado['Requisicao'], None)
                if portal is not None:
                    try:
                        portal.close()
                    except Exception:
                        pass
                print('[NC CONCILIAÇÃO] Falha de navegação antes do envio; reabrindo o formulário nesta execução.')
                continue
            raise
        if estado['Fase'] not in {'CRIADA_ANEXO_PENDENTE', 'CONCLUIDA'}:
            raise ConciliacaoNCError('Rotina de envio retornou sem confirmar o número da conciliação.')


def preparar_recuperacao_legada_sem_numero_nc(page, pasta, estado, pedido, consulta):
    """Uma tentativa adicional para o erro legado observado em REQ_DEMO.

    Diretriz de validação: criar Glosa automaticamente após a NC
    validada e encerrada, sem confirmação manual. Esta exceção NÃO comprova que
    o envio antigo falhou: o risco de duplicidade fica registrado no controle.
    Não vale para envios rastreados por esta versão, número já conhecido,
    candidatos encontrados, uma aba de envio ainda presente ou outro erro.
    """
    if (estado.get('Fase') != 'ENVIO_INICIADO'
            or estado.get('OrigemOVs') != 'PORTAL_NC_DUAS_PLANILHAS'
            or estado.get('RequisicaoConciliacao')
            or 'TentativaEnvioConciliacao' in estado
            or 'RecuperacaoLegadaSemNumero' in estado
            or consulta.get('Completa') or consulta.get('Numeros')
            or globals().get('PORTAIS_ENVIO_NC', {}).get(estado['Requisicao']) is not None):
        return False
    retorno = normalizar_texto_busca(estado.get('UltimoRetornoEnvio', ''))
    motivo = normalizar_texto_busca(consulta.get('Motivo', ''))
    if ('numero de conciliacao invalido ou igual ao portal original' not in retorno
            or 'menu de requisicoes do solicitante nao reconhecido' not in motivo):
        return False
    rodada = globals().get('ID_RODADA', '')
    if not rodada or estado.get('RodadaUltimoReenvio') == rodada:
        return False
    # Reconfere a aprovação, a associação OV/documento, os hashes e o PDF
    # realmente anexado à origem encerrada. Não usa só o status do relatório.
    validar_documentos_conciliacao_nc(estado)
    confirmar_original_encerrada_nc(page, pasta, estado)
    registro = {
        'Em': datetime.now().isoformat(timespec='seconds'),
        'Acao': 'TENTATIVA_LEGADA_UNICA_PELA_DIRETRIZ_AUTOMATICA',
        'Situacao': 'PREPARADA', 'RequisicaoOrigem': estado['Requisicao'],
        'OV': estado['Validacao']['Ordem'],
        'DocumentoFaturamento': estado['ConsultaSAP']['NumeroNF'],
        'HashPlanilha': estado['HashPlanilha'], 'HashPDF': estado['HashPDF'],
        'HashPedido': hashlib.sha256(json.dumps(pedido, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest(),
        'Rodada': rodada, 'AusenciaAnteriorConfirmada': False,
        'MotivoConsulta': consulta.get('Motivo', ''),
        'RetornoLegado': estado.get('UltimoRetornoEnvio', ''),
        'EnvioLegadoEm': estado.get('PrimeiroEnvioIncertoEm', ''),
        'Risco': 'O envio antigo pode ter criado uma conciliação cujo número não foi registrado.'}
    salvar_controle_nc(
        pasta, estado, Fase='ORIGINAL_ENCERRADA', Pedido=pedido,
        RodadaUltimoReenvio=rodada, RecuperacaoLegadaSemNumero=registro,
        HistoricoRetomadasConciliacao=list(estado.get('HistoricoRetomadasConciliacao') or []) + [dict(registro)],
        MensagemConciliacao='Retomada legada: uma tentativa adicional automática. '
                            'A ausência da conciliação anterior não foi comprovada.')
    print(f'[NC CONCILIAÇÃO] Origem {estado["Requisicao"]}: envio legado sem número; '
          'uma tentativa adicional automática registrada, sem repetir VF01 ou encerramento. '
          'A existência de conciliação anterior permanece desconhecida.')
    return True


def resolver_envio_iniciado_nc(page, pasta, estado):
    """Retoma sem input: aproveita recibo/listagem ou avalia a política de retomada.

    Não altera o SAP, não aceita número igual à origem e não interpreta falha de
    pesquisa como ausência. O erro legado conhecido admite uma tentativa única
    pela diretriz automática; os demais dependem de ausência confirmada.
    """
    if estado['Fase'] != 'ENVIO_INICIADO':
        raise ConciliacaoNCError('A recuperação automática exige um envio anterior sem confirmação.')
    pedido = montar_solicitacao_conciliacao_nc(estado)
    req = estado['Requisicao']
    inicio = estado.get('EnvioIniciadoEm') or estado.get('PrimeiroEnvioIncertoEm') or estado.get('AtualizadoEm')
    retorno = str(estado.get('UltimoRetornoEnvio') or '')
    anterior = str(estado.get('MensagemConciliacao') or '')
    if not retorno and anterior and not anterior.startswith(('Envio anterior sem número confirmado', 'Consulta automática')):
        retorno = anterior
    salvar_controle_nc(pasta, estado, PrimeiroEnvioIncertoEm=inicio, UltimoRetornoEnvio=retorno)
    confirmar_original_encerrada_nc(page, pasta, estado)
    print(f'[NC CONCILIAÇÃO] Origem {req}: conferindo automaticamente o envio anterior.')
    if estado.get('RequisicaoConciliacao'):
        vincular_conciliacao_automatica_nc(page, pasta, estado, pedido,
                                         str(estado['RequisicaoConciliacao']), 'CONTROLE')
        return
    # Primeiro relê o recibo da própria aba do envio, se ainda existir na sessão.
    portal = globals().get('PORTAIS_ENVIO_NC', {}).get(req)
    if portal is not None:
        try:
            tela = ler_controles_portal_nc(portal)
            numero = numero_conciliacao_criada_nc(tela, req)
        except Exception:
            tela, numero = {}, ''
        if numero:
            vincular_conciliacao_automatica_nc(page, pasta, estado, pedido, numero, 'RECIBO_PORTAL')
            portal.close()
            globals()['PORTAIS_ENVIO_NC'].pop(req, None)
            return
    if aguardar_confirmacao_envio_nc(page, pasta, estado, pedido, inicio):
        return
    consulta = consultar_conciliacoes_existentes_nc(page, estado, pedido)
    salvar_controle_nc(pasta, estado, ConsultaAutomaticaConciliacao=consulta)
    numeros = consulta['Numeros']
    if len(numeros) == 1:
        vincular_conciliacao_automatica_nc(page, pasta, estado, pedido, numeros[0], 'LISTAGEM_PORTAL')
        return
    if len(numeros) > 1:
        raise ConciliacaoNCError('Consulta automática: mais de um chamado menciona a origem; '
                                'novo envio bloqueado. Candidatos: ' + ', '.join(numeros))
    if not consulta['Completa']:
        if preparar_recuperacao_legada_sem_numero_nc(page, pasta, estado, pedido, consulta):
            return
        if estado.get('RecuperacaoLegadaSemNumero'):
            raise ConciliacaoNCError('A tentativa adicional do envio legado já foi registrada; '
                                    'não será repetida sem localizar o chamado ou confirmar a ausência. '
                                    + consulta['Motivo'])
        raise ConciliacaoNCError('Consulta automática inconclusiva; novo envio bloqueado nesta rodada. '
                                + consulta['Motivo'] + (' Último retorno do envio: ' + retorno if retorno else ''))
    rodada = globals().get('ID_RODADA', '')
    if not rodada or estado.get('RodadaUltimoReenvio') == rodada:
        raise ConciliacaoNCError('Consulta automática: limite de uma nova tentativa por rodada; novo envio bloqueado.')
    if portal is not None:
        tela = ler_controles_portal_nc(portal)
        if tela.get('pendente') or separar_avisos_upload_portal_nc(tela)[0]:
            raise ConciliacaoNCError('Consulta automática: portal ainda processando o envio; novo envio bloqueado.')
    segunda = consultar_conciliacoes_existentes_nc(page, estado, pedido)
    salvar_controle_nc(pasta, estado, SegundaConsultaAutomaticaConciliacao=segunda)
    if len(segunda['Numeros']) == 1:
        vincular_conciliacao_automatica_nc(page, pasta, estado, pedido, segunda['Numeros'][0], 'SEGUNDA_LISTAGEM_PORTAL')
        return
    if not segunda['Completa'] or segunda['Numeros']:
        raise ConciliacaoNCError('Consulta automática: ausência não confirmada na segunda leitura; novo envio bloqueado. '
                                + segunda['Motivo'])
    validar_documentos_conciliacao_nc(estado)
    confirmar_original_encerrada_nc(page, pasta, estado)
    registro = {'Em': datetime.now().isoformat(timespec='seconds'), 'Acao': 'AUSENCIA_CONFERIDA_AUTOMATICAMENTE',
                'RequisicaoOrigem': req, 'OV': estado['Validacao']['Ordem'],
                'DocumentoFaturamento': estado['ConsultaSAP']['NumeroNF'],
                'TotalLido': segunda['TotalLido'], 'Rodada': rodada}
    salvar_controle_nc(pasta, estado, Fase='ORIGINAL_ENCERRADA', RodadaUltimoReenvio=rodada,
                      HistoricoRetomadasConciliacao=list(estado.get('HistoricoRetomadasConciliacao') or []) + [registro],
                      MensagemConciliacao='Ausência conferida em duas consultas completas; criação será retomada automaticamente.')
    if portal is not None:
        portal.close()
        globals()['PORTAIS_ENVIO_NC'].pop(req, None)
    print(f'[NC CONCILIAÇÃO] Origem {req}: ausência conferida; retomando somente a criação da conciliação.')


def continuar_conciliacao_nc(page, pasta, estado, retomada=False):
    """Não chama VF01: retoma exclusivamente PORTAL/conciliação a partir do controle."""
    try:
        if estado['Fase'] == 'CONCLUIDA':
            return resultado_controle_nc(estado, retomada)
        if estado['Fase'] in {'SAP_EM_EXECUCAO', 'SAP_PENDENTE'}:
            raise ConciliacaoNCError('Etapa SAP pendente ou interrompida. Conferir o documento no SAP antes de qualquer nova execução; VF01 não repetida.')
        validar_documentos_conciliacao_nc(estado)
        if estado['Fase'] in {'SAP_CONFIRMADO', 'PORTAL_FINALIZACAO'}:
            frame = abrir_requisicao_portal(page, estado['Requisicao'])
            status = confirmar_status_cabecalho_portal(page, frame, estado['Requisicao'])
            if normalizar_texto_busca(status).strip() != 'encerrada':
                if status_requisicao_portal_nao_acionavel(status):
                    raise ConciliacaoNCError(f'PORTAL original em estado {status}; não abrir conciliação.')
                salvar_controle_nc(pasta, estado, Fase='PORTAL_FINALIZACAO')
                resultado = anexar_pdfs_e_encerrar_portal(page, estado['Requisicao'], pasta, [estado['ConsultaSAP']])
                confirmado_pela_rotina = dict(resultado)
                if resultado.get('StatusPORTAL') == 'CONCLUIDO':
                    confirmado_pela_rotina['StatusPORTAL'] = 'ENCERRAMENTO_A_CONFIRMAR'
                    confirmado_pela_rotina['MensagemPORTAL'] += ' | Aguardando releitura do status e conferência do PDF.'
                salvar_controle_nc(pasta, estado, ResultadoPORTAL=confirmado_pela_rotina)
                if resultado.get('StatusPORTAL') != 'CONCLUIDO':
                    raise ConciliacaoNCError('Conciliação não iniciada: ' + resultado.get('MensagemPORTAL', 'PORTAL original não concluído.'))
            confirmar_original_encerrada_nc(page, pasta, estado)
            salvar_controle_nc(pasta, estado, Fase='ORIGINAL_ENCERRADA')
        elif estado['Fase'] == 'ORIGINAL_ENCERRADA':
            confirmar_original_encerrada_nc(page, pasta, estado)
        concluir_conciliacao_na_mesma_execucao_nc(page, pasta, estado)
    except Exception as erro:
        mensagem = str(erro)
        alteracoes = {'MensagemConciliacao': mensagem}
        if (estado['Fase'] == 'ENVIO_INICIADO' and not estado.get('UltimoRetornoEnvio')
                and not mensagem.startswith('Envio anterior sem número confirmado')):
            alteracoes['UltimoRetornoEnvio'] = mensagem
        salvar_controle_nc(pasta, estado, **alteracoes)
        print(f'[NC CONCILIAÇÃO] PORTAL {estado["Requisicao"]}: {mensagem}')
    return resultado_controle_nc(estado, retomada)


def executar_nc_duas_planilhas(page, page_fiori, req, pasta, atividade, descricao, dados, consultas):
    """Ramo novo. Os fluxos de uma planilha e Locação não passam por aqui."""
    with TravaNC(pasta):
        anterior = carregar_controle_nc(pasta, req)
        if anterior and anterior['Fase'] != 'VALIDACAO_REPROVADA':
            raise ConciliacaoNCError('Há controle anterior desta NC; retomar pela etapa registrada.')
        if classificar_atividade_suporte_venda(atividade) != ATIVIDADE_NOTA_CREDITO or not dados:
            raise ConciliacaoNCError('Conciliação permitida apenas para NC validada com dois Excels.')
        estado = {'VersaoControle': 1, 'Requisicao': str(req), 'Atividade': atividade,
                  'OrigemOVs': 'PORTAL_NC_DUAS_PLANILHAS', 'Descricao': descricao,
                  'Validacao': dados, 'HashPlanilha': dados['HashPlanilhaValidada']}
        if hash_arquivo_nc(dados['CaminhoPlanilha']) != estado['HashPlanilha']:
            raise ConciliacaoNCError('Planilha alterada após a leitura; SAP não iniciado.')
        preparar_fiori_nova_requisicao(page_fiori=page_fiori, requisicao=req)
        salvar_controle_nc(pasta, estado, Fase='SAP_EM_EXECUCAO')
        try:
            consulta = processar_ov_sap(page_fiori=page_fiori, requisicao=req, ordem=dados['Ordem'],
                                      atividade=atividade, validacao_credito=dados)
        except Exception as erro:
            raise ConciliacaoNCError(f'Execução SAP interrompida: {erro}. Controle preservado; conferir antes de repetir.') from erro
        consultas[:] = [consulta]
        validar_associacao_resultado_sap(req, dados['Ordem'], consulta)
        if consulta.get('StatusSAP') == 'NC_VALIDACAO_REPROVADA':
            salvar_controle_nc(pasta, estado, Fase='VALIDACAO_REPROVADA', ConsultaSAP=consulta,
                              ResultadoPORTAL={'StatusPORTAL': 'PENDENTE_VALIDACAO_NC', 'QtdPDFsAnexados': 0,
                                            'MensagemPORTAL': consulta.get('Mensagem', '')},
                              MensagemConciliacao='Não iniciada: validação prévia reprovada.')
            return resultado_controle_nc(estado)
        salvar_controle_nc(pasta, estado, Fase='SAP_PENDENTE', ConsultaSAP=consulta)
        if consulta.get('StatusSAP') == 'OK_VA03':
            salvar_controle_nc(pasta, estado, HashPDF=hash_arquivo_nc(consulta['ArquivoPDF']))
            validar_documentos_conciliacao_nc(estado)
            salvar_controle_nc(pasta, estado, Fase='SAP_CONFIRMADO')
        return continuar_conciliacao_nc(page, pasta, estado)


def retomar_nc_duas_planilhas(page, req, pasta, consultas):
    with TravaNC(pasta):
        estado = carregar_controle_nc(pasta, req)
        if estado is None or estado['Fase'] == 'VALIDACAO_REPROVADA':
            raise ConciliacaoNCError('Controle para retomada não encontrado.')
        consulta = estado.get('ConsultaSAP') or {
            'Ordem': estado['Validacao']['Ordem'], 'StatusSAP': 'PENDENTE_CONFERENCIA_SAP',
            'Mensagem': 'Execução SAP anterior interrompida; faturamento não repetido.'}
        consultas[:] = [consulta]
        return continuar_conciliacao_nc(page, pasta, estado, retomada=True)


def frame_principal_portal(page):
    """Retorna o frame principal do PORTAL usado em todo o robô."""
    return (
        page.frame_locator(FRAME_ABERTURA)
        .frame_locator(FRAME_PRINCIPAL)
    )


def frame_portalmain(frame):
    """Retorna o iframe PortalPainelDemo."""
    return frame.frame_locator(FRAME_PORTAL_MAIN)


def restaurar_browser_portal(page):
    """
    Traz a aba do PORTAL para frente antes de anexar e concluir.

    Na versão Fiori o navegador não é mais minimizado para alternar entre SAP
    e PORTAL. Mantemos apenas o controle nativo do próprio Playwright/CDP.
    """
    try:
        page.bring_to_front()
        client = page.context.new_cdp_session(page)
        info = client.send("Browser.getWindowForTarget")
        window_id = info.get("windowId")
        if window_id is not None:
            try:
                client.send("Browser.setWindowBounds", {
                    "windowId": window_id,
                    "bounds": {"windowState": "normal"}
                })
                time.sleep(0.8)
            except Exception as e_normal:
                print(
                    f"[PORTAL] Aviso ao restaurar {NOME_NAVEGADOR} "
                    f"para normal via CDP: {e_normal}"
                )

            client.send("Browser.setWindowBounds", {
                "windowId": window_id,
                "bounds": {"windowState": "maximized"}
            })
            time.sleep(1)
            try:
                page.bring_to_front()
            except Exception:
                pass
            print(
                f"[PORTAL] {NOME_NAVEGADOR}/PORTAL restaurado para "
                "anexo/encerramento via CDP."
            )
            return True
    except Exception as e:
        print(f"[PORTAL] Não consegui restaurar {NOME_NAVEGADOR}/PORTAL via CDP: {e}")
        try:
            page.bring_to_front()
            return True
        except Exception:
            return False


PADRAO_SUCESSO_GRAVAR_PORTAL = re.compile(
    r"Opera(?:c|ç)[aã]o\s+efetuada\s+com\s+sucesso\.?",
    re.I,
)
PADRAO_BOTAO_OK_PORTAL = re.compile(r"^\s*Ok\s*$", re.I)


def _primeiro_locator_visivel(locator, limite=10):
    """Retorna o primeiro elemento visível de um locator, sem assumir unicidade."""
    try:
        quantidade = min(locator.count(), limite)
    except Exception:
        return None

    for indice in range(quantidade):
        candidato = locator.nth(indice)
        try:
            if candidato.is_visible():
                return candidato
        except Exception:
            continue

    return None


def _acionar_ok_visivel_no_contexto(contexto):
    """Localiza e aciona um botão/texto Ok visível no contexto informado."""
    candidatos = [
        contexto.get_by_role("button", name=PADRAO_BOTAO_OK_PORTAL),
        contexto.locator("button").filter(has_text=PADRAO_BOTAO_OK_PORTAL),
        contexto.locator(
            "input[type='button'][value='Ok' i], "
            "input[type='submit'][value='Ok' i]"
        ),
        contexto.get_by_text(PADRAO_BOTAO_OK_PORTAL),
    ]

    erros = []
    for locator in candidatos:
        alvo = _primeiro_locator_visivel(locator)
        if alvo is None:
            continue

        try:
            alvo.click(timeout=4000)
            return True, "clique normal"
        except Exception as e_normal:
            erros.append(str(e_normal))

        try:
            alvo.click(timeout=3000, force=True)
            return True, "clique forçado"
        except Exception as e_forcado:
            erros.append(str(e_forcado))

        try:
            alvo.evaluate("elemento => elemento.click()")
            return True, "clique pelo DOM"
        except Exception as e_dom:
            erros.append(str(e_dom))

    detalhe = erros[-1] if erros else "nenhum botão Ok visível no contexto"
    return False, detalhe


def fechar_confirmacao_gravar_portal(
    page,
    frame_principal=None,
    timeout=15000,
    aguardar_aparecer=True,
):
    """
    Fecha a confirmação HTML do PORTAL:
    Gravar -> Operação efetuada com sucesso. -> Ok.

    AV8:
    essa confirmação pode ser renderizada em um frame diferente daquele que
    contém o botão Salvar. Por isso a busca percorre todos os frames reais da
    página, inclusive frames internos, antes de seguir para o próximo anexo.
    """
    limite = time.monotonic() + max(timeout, 0) / 1000
    ultimo_detalhe = ""

    while True:
        contextos = []

        # Frames internos primeiro: o modal de gravação costuma estar em um
        # deles, embora visualmente cubra toda a tela do PORTAL.
        try:
            for indice, fr in enumerate(reversed(page.frames)):
                nome = fr.name or ("<principal>" if fr == page.main_frame else "<sem nome>")
                contextos.append((f"frame '{nome}'", fr))
        except Exception as e_frames:
            ultimo_detalhe = str(e_frames)

        # Mantém também o FrameLocator já validado do PORTAL como fallback.
        if frame_principal is not None:
            contextos.append(("frame principal do PORTAL", frame_principal))

        for rotulo, contexto in contextos:
            try:
                mensagens = contexto.get_by_text(PADRAO_SUCESSO_GRAVAR_PORTAL)
                mensagem = _primeiro_locator_visivel(mensagens)
                if mensagem is None:
                    continue

                print(f"[PORTAL] Confirmação 'Operação efetuada com sucesso' localizada no {rotulo}.")
                ok, modo = _acionar_ok_visivel_no_contexto(contexto)
                if not ok:
                    ultimo_detalhe = f"{rotulo}: {modo}"
                    continue

                # Não basta disparar o clique: só liberamos o próximo PDF
                # depois que a confirmação realmente deixar de bloquear a tela.
                try:
                    mensagem.wait_for(state="hidden", timeout=5000)
                except Exception:
                    try:
                        if mensagem.is_visible():
                            ultimo_detalhe = (
                                f"{rotulo}: o Ok recebeu {modo}, mas a janela permaneceu visível"
                            )
                            continue
                    except Exception:
                        # Frame/modal removido do DOM após o clique: fechamento confirmado.
                        pass

                print(f"[PORTAL] Janela 'Gravar' fechada pelo botão Ok ({modo}).")
                return True

            except Exception as e_contexto:
                ultimo_detalhe = f"{rotulo}: {e_contexto}"
                continue

        if not aguardar_aparecer or time.monotonic() >= limite:
            if ultimo_detalhe and aguardar_aparecer:
                print(f"[PORTAL] Detalhe da busca pela confirmação de gravação: {ultimo_detalhe}")
            return False

        try:
            page.wait_for_timeout(250)
        except Exception:
            time.sleep(0.25)


def fechar_modal_residual_portal(page, motivo="", timeout=8000):
    """
    TESTE 35 - limpa overlays/modais residuais do portal de atendimento.

    Diagnóstico da rodada 34:
    requisições que usam planilha podem deixar o modal de seleção/documentos
    aberto depois do download. Visualmente ele pode ficar branco, com apenas
    o X no canto. Enquanto permanece aberto, o Bootstrap intercepta cliques em:
      - Assumir Responsabilidade
      - Documentos
      - demais abas da requisição
    e o problema pode vazar para a próxima requisição.

    Regras de segurança:
    1. fecha primeiro uma confirmação legítima "Operação efetuada com sucesso";
    2. tenta especificamente #modal-selecao2, já observado no Codegen;
    3. procura somente modais VISÍVEIS no contêiner #auxiliares;
    4. aciona apenas controles neutros de fechamento (X / Fechar / Close /
       Cancelar); nunca executa ação de negócio;
    5. confirma que o overlay deixou de bloquear antes de seguir.
    """
    inicio = time.monotonic()
    limite = inicio + max(1.0, timeout / 1000.0)
    encontrou = False
    ultimo_log = -999

    # Uma confirmação de gravação pendente deve ser tratada pelo fluxo próprio.
    try:
        if fechar_confirmacao_gravar_portal(
            page,
            frame_principal=frame_principal_portal(page),
            timeout=0,
            aguardar_aparecer=False,
        ):
            print(
                "[PORTAL] Confirmação de gravação pendente foi fechada antes de "
                "limpar modais residuais."
            )
            encontrou = True
    except Exception:
        pass

    while time.monotonic() < limite:
        try:
            principal = frame_principal_portal(page)
        except Exception:
            return encontrou

        candidatos = []

        # Modal conhecido do fluxo de Documentos/planilha.
        try:
            candidatos.append(("modal-selecao2", principal.locator("#modal-selecao2")))
        except Exception:
            pass

        # Modais Bootstrap/legados renderizados no contêiner auxiliar do PORTAL.
        for seletor in (
            "#auxiliares .modal",
            "#auxiliares [role='dialog']",
            ".modal.show",
            ".modal.in",
        ):
            try:
                candidatos.append((seletor, principal.locator(seletor)))
            except Exception:
                continue

        visiveis = []
        vistos = set()

        for rotulo, locator in candidatos:
            try:
                qtd = min(locator.count(), 20)
            except Exception:
                qtd = 0

            for i in range(qtd):
                modal = locator.nth(i)
                try:
                    if not modal.is_visible(timeout=200):
                        continue
                except Exception:
                    continue

                try:
                    chave = modal.get_attribute("id") or f"{rotulo}:{i}"
                except Exception:
                    chave = f"{rotulo}:{i}"

                if chave in vistos:
                    continue
                vistos.add(chave)
                visiveis.append((rotulo, modal, chave))

        if not visiveis:
            if encontrou:
                print(
                    "[PORTAL] Modal/overlay residual removido"
                    + (f" ({motivo})" if motivo else "")
                    + "."
                )
            return encontrou

        encontrou = True
        decorrido = time.monotonic() - inicio
        if int(decorrido) - ultimo_log >= 2:
            ultimo_log = int(decorrido)
            print(
                f"[PORTAL] {len(visiveis)} modal(is) residual(is) visível(is)"
                + (f" durante {motivo}" if motivo else "")
                + "; tentando liberar a interface."
            )

        fechou_algum = False

        for rotulo, modal, chave in visiveis:
            try:
                texto_modal = re.sub(
                    r"\s+",
                    " ",
                    str(modal.inner_text(timeout=700) or ""),
                ).strip()
            except Exception:
                texto_modal = ""

            if texto_modal:
                print(
                    f"[PORTAL] Modal residual {chave}: "
                    f"{texto_modal[:220]}"
                )
            else:
                print(
                    f"[PORTAL] Modal residual {chave}: sem texto útil "
                    "(possível janela branca de Documentos)."
                )

            # Se for uma confirmação de sucesso que surgiu agora, usa o helper
            # específico em vez de simplesmente fechar o X.
            if PADRAO_SUCESSO_GRAVAR_PORTAL.search(texto_modal or ""):
                try:
                    if fechar_confirmacao_gravar_portal(
                        page,
                        frame_principal=principal,
                        timeout=3000,
                        aguardar_aparecer=True,
                    ):
                        fechou_algum = True
                        continue
                except Exception:
                    pass

            controles = [
                modal.locator(
                    "button.close, .modal-header .close, "
                    "[data-dismiss='modal'], [data-bs-dismiss='modal']"
                ),
                modal.get_by_text("×", exact=True),
                modal.get_by_role(
                    "button",
                    name=re.compile(r"^\s*(?:Fechar|Close|Cancelar|Cancel|×)\s*$", re.I),
                ),
            ]

            for controle in controles:
                try:
                    alvo = _primeiro_locator_visivel(controle, limite=10)
                    if alvo is None:
                        continue

                    try:
                        alvo.click(timeout=3500)
                        modo = "clique normal"
                    except Exception:
                        try:
                            alvo.click(timeout=2500, force=True)
                            modo = "clique forçado"
                        except Exception:
                            alvo.evaluate("el => el.click()")
                            modo = "clique DOM"

                    print(
                        f"[PORTAL] Fechamento neutro acionado no modal {chave} "
                        f"({modo})."
                    )
                    fechou_algum = True
                    page.wait_for_timeout(500)
                    break
                except Exception:
                    continue

        if not fechou_algum:
            # Último recurso seguro para um modal residual: Escape.
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(500)
                print("[PORTAL] Escape enviado para liberar modal residual.")
            except Exception:
                pass

        page.wait_for_timeout(300)

    # Não lança exceção aqui: o chamador pode tentar a ação e produzir um erro
    # mais específico caso o modal realmente continue bloqueando.
    print(
        "[PORTAL] Aviso: ainda pode existir modal/overlay residual após "
        f"{int(timeout/1000)} s"
        + (f" ({motivo})." if motivo else ".")
    )
    return encontrou


def clicar_ok_portal(frame, timeout=10000, page=None):
    """Clica em Ok em mensagens do PORTAL, incluindo modais existentes em frames internos."""
    if page is not None:
        if fechar_confirmacao_gravar_portal(
            page,
            frame_principal=frame,
            timeout=timeout,
            aguardar_aparecer=True,
        ):
            return True

    tentativas = [
        lambda: frame.get_by_text("Ok", exact=True).click(timeout=timeout),
        lambda: frame.get_by_text("Ok").click(timeout=timeout),
        lambda: frame.get_by_role("button", name=re.compile(r"^\s*Ok\s*$", re.I)).click(timeout=timeout),
    ]

    for tentativa in tentativas:
        try:
            tentativa()
            time.sleep(1)
            return True
        except Exception:
            pass

    return False


def configurar_dialogs_seguros(page):
    """Confirma somente ações reconhecidas; rejeita nova versão e confirm desconhecido."""
    def tratar_dialog(dialog):
        msg = dialog.message or ""
        normal = normalizar_texto_busca(msg)
        duplicado = "arquivo ja foi anexado" in normal
        reconhecido = bool(re.fullmatch(
            r"\s*(?:(?:voce deseja|deseja|confirma(?: que deseja)?|confirmar|tem certeza(?: de)?(?: que deseja)?)\s+)?"
            r"(?:realmente\s+)?(?:assumir(?: a)? responsabilidade(?: (?:desta|pela|por esta|da) requisicao)?|"
            r"(?:concluir|encerrar)(?: a)? requisicao|(?:a )?conclusao da requisicao|"
            r"(?:salvar|gravar)(?: (?:os |as |a )?(?:dados|alteracoes|conclusao))?)"
            r"\s*[?.!]?\s*", normal))
        aceitar = not duplicado and (dialog.type == "alert" or
                                     dialog.type == "confirm" and reconhecido)
        PORTAL_DIALOG_STATE["ultimo_dialog"] = msg
        PORTAL_DIALOG_STATE["historico"].append(msg)
        PORTAL_DIALOG_STATE.setdefault("eventos", []).append({
            "mensagem": msg, "tipo": dialog.type, "aceito": aceitar,
            "bloqueado": not aceitar and not duplicado,
        })
        if duplicado:
            PORTAL_DIALOG_STATE["arquivo_ja_anexado"] = True
        print(f"[PORTAL DIALOG] {dialog.type}: {msg} | {'aceito' if aceitar else 'recusado'}")
        try:
            if aceitar:
                dialog.accept()
            else:
                dialog.dismiss()
        except PlaywrightError as erro:
            print(f"[PORTAL DIALOG] Não foi possível tratar a confirmação: {erro}")
    page.on("dialog", tratar_dialog)
    return True


def verificar_dialogs_portal_desde(indice):
    """Um confirm recusado ou alerta de falha não comprova sucesso da ação."""
    for evento in PORTAL_DIALOG_STATE.get("eventos", [])[indice:]:
        normal = normalizar_texto_busca(evento.get("mensagem", ""))
        falha = re.search(r"\b(?:erro|falha|negado|negada|nao permitido|nao autorizad[oa]|"
                          r"sem permissao|outro responsavel|outra pessoa|nao foi possivel)\b", normal)
        if evento.get("bloqueado") or falha:
            raise RuntimeError("Ação do PORTAL não confirmada: " + evento.get("mensagem", ""))

def _localizar_fila_suporte_venda_visivel(arvore, timeout_ms=8000):
    """
    Retorna somente o label VISÍVEL da fila Equipe de Operações - Faturamento.

    O portal de atendimento mantém cópias ocultas do mesmo label na árvore. Usar
    simplesmente `.first` pode escolher uma cópia invisível e gerar timeout,
    mesmo com a fila correta aparecendo na tela.
    """
    padrao = re.compile(
        re.escape(FILA_PORTAL_SUPORTE_VENDA).replace(r"\ ", r"\s*") + r"(?:\s*\(\d+\))?",
        re.I,
    )
    limite = time.monotonic() + max(1.0, timeout_ms / 1000.0)
    ultimo_total = 0

    while time.monotonic() < limite:
        candidatos = arvore.get_by_text(padrao)
        try:
            ultimo_total = candidatos.count()
        except Exception:
            ultimo_total = 0

        for indice in range(ultimo_total):
            candidato = candidatos.nth(indice)
            try:
                if candidato.is_visible(timeout=500):
                    return candidato
            except Exception:
                continue

        time.sleep(0.20)

    raise RuntimeError(
        "Fila Equipe de Operações - Faturamento não possui label visível no "
        f"frameArvoreDemo. Candidatos encontrados no DOM: {ultimo_total}."
    )


def _acessar_fila_suporte_venda_portal_uma_vez(page):
    """Uma tentativa de abrir Lista de Chamados -> ATENDIMENTO -> Equipe de Operações - Faturamento."""
    frame = frame_principal_portal(page)
    print("[PORTAL] Acessando Lista de Chamados...")

    try:
        link_lista = frame.get_by_role("link", name="Lista de Chamados")
        link_lista.wait_for(state="visible", timeout=8000)
        link_lista.click(timeout=15000)
        page.wait_for_timeout(2000)
    except Exception as e:
        print(
            "[PORTAL] Lista de Chamados não está disponível na tela atual; "
            f"retornando ao atendimento. Detalhe: {e}"
        )
        page.goto(
            PORTAL_URL_ATENDIMENTO,
            wait_until="domcontentloaded",
            timeout=60000,
        )
        page.wait_for_timeout(3000)
        frame = frame_principal_portal(page)
        link_lista = frame.get_by_role("link", name="Lista de Chamados")
        link_lista.wait_for(state="visible", timeout=30000)
        link_lista.click(timeout=20000)
        page.wait_for_timeout(2500)

    portalmain = frame_portalmain(frame)
    arvore = portalmain.frame_locator(FRAME_ARVORE)
    resultado = portalmain.frame_locator(FRAME_RESULTADOS)

    # TESTE 12: o PORTAL pode manter dezenas de labels ocultos com o mesmo texto.
    # Primeiro procuramos explicitamente a instância visível.
    try:
        categoria = _localizar_fila_suporte_venda_visivel(arvore, timeout_ms=2500)
    except Exception:
        print("[PORTAL] Expandindo a árvore ATENDIMENTO...")

        expandiu = False
        for seletor in (PORTAL_SELETOR_EXPANDIR_ARVORE, PORTAL_SELETOR_ARVORE_ALTERNATIVO):
            try:
                controle = arvore.locator(seletor).first
                if controle.is_visible(timeout=1500):
                    controle.click(timeout=15000)
                    expandiu = True
                    break
            except Exception:
                continue

        if not expandiu:
            print(
                "[PORTAL] Controle padrão de ATENDIMENTO não estava visível; "
                "vou aguardar a árvore materializar a fila."
            )

        page.wait_for_timeout(1200)
        categoria = _localizar_fila_suporte_venda_visivel(
            arvore, timeout_ms=20000
        )

    try:
        texto_categoria = str(categoria.inner_text(timeout=3000) or "").strip()
    except Exception:
        texto_categoria = FILA_PORTAL_SUPORTE_VENDA

    print(f"[PORTAL] Fila visível selecionada na árvore: {texto_categoria}")

    # TESTE 14:
    # O body do frameResultadosDemo pode ficar visível antes de o PORTAL terminar de
    # materializar as linhas da fila. No TESTE 13 isso gerou uma falsa fila
    # vazia: o label mostrava dezenas de chamados, mas o mapeamento começou
    # enquanto o grid ainda tinha 0 links de requisição.
    #
    # Agora só consideramos a fila carregada depois de aparecer pelo menos um
    # número de requisição no frameResultadosDemo. Se o PORTAL não materializar o grid,
    # clicamos novamente na fila visível e reacquirimos os frames.
    grid_carregado = False
    ultimo_total_links = 0

    for tentativa_grid in range(1, 4):
        try:
            categoria.scroll_into_view_if_needed(timeout=10000)
            categoria.click(timeout=20000)
        except Exception as e:
            print(
                f"[PORTAL] Tentativa {tentativa_grid}/3: clique na fila não foi "
                f"confirmado. Detalhe: {e}"
            )

        limite_grid = time.monotonic() + 15.0
        while time.monotonic() < limite_grid:
            page.wait_for_timeout(350)
            try:
                links_req = resultado.get_by_role(
                    "link", name=re.compile(r"^\s*\d{7}\s*$")
                )
                ultimo_total_links = links_req.count()
                if ultimo_total_links > 0:
                    grid_carregado = True
                    print(
                        f"[PORTAL] Grid da fila materializado com "
                        f"{ultimo_total_links} requisição(ões) na página atual."
                    )
                    break
            except Exception:
                ultimo_total_links = 0

        if grid_carregado:
            break

        if tentativa_grid < 3:
            print(
                f"[PORTAL] Grid ainda sem requisições após a tentativa "
                f"{tentativa_grid}/3; reacessando a fila antes de mapear."
            )
            # Reacquire porque o portal de atendimento pode substituir internamente os
            # iframes ao trocar o contexto da árvore.
            frame = frame_principal_portal(page)
            portalmain = frame_portalmain(frame)
            arvore = portalmain.frame_locator(FRAME_ARVORE)
            resultado = portalmain.frame_locator(FRAME_RESULTADOS)

            try:
                categoria = _localizar_fila_suporte_venda_visivel(
                    arvore, timeout_ms=8000
                )
            except Exception:
                try:
                    controle = arvore.locator(f"{PORTAL_SELETOR_EXPANDIR_ARVORE}, {PORTAL_SELETOR_ARVORE_ALTERNATIVO}").first
                    if controle.is_visible(timeout=1500):
                        controle.click(timeout=10000)
                        page.wait_for_timeout(800)
                except Exception:
                    pass
                categoria = _localizar_fila_suporte_venda_visivel(
                    arvore, timeout_ms=12000
                )

    if not grid_carregado:
        raise RuntimeError(
            "A fila Equipe de Operações - Faturamento foi selecionada, mas o grid "
            "frameResultadosDemo não materializou nenhuma requisição após 3 tentativas. "
            f"Última contagem de links: {ultimo_total_links}."
        )

    print(f"[PORTAL] Fila {FILA_PORTAL_SUPORTE_VENDA} carregada.")
    return frame, resultado


def _recarregar_portal_quando_incompleto(page, tentativa, motivo):
    """
    Recupera o portal de atendimento quando o frameset carrega parcialmente.

    Em alguns acessos o campo "Consultar requisição..." já existe, porém a
    árvore/frameArvoreDemo ainda não foi materializada. Um F5 manual recompõe o
    frameset; aqui reproduzimos esse comportamento de forma controlada.
    """
    print(
        f"[PORTAL] Estrutura do PORTAL incompleta na tentativa {tentativa}. "
        f"Detalhe: {motivo}"
    )
    print(
        "[PORTAL] Executando atualização controlada da página (equivalente a F5) "
        "para reconstruir o frameset do portal de atendimento..."
    )

    try:
        page.bring_to_front()
    except Exception:
        pass

    atualizou = False
    try:
        page.reload(wait_until="domcontentloaded", timeout=60000)
        atualizou = True
        print("[PORTAL] F5 controlado concluído; aguardando a tela principal do PORTAL...")
    except Exception as e_reload:
        print(f"[PORTAL] Reload não concluiu normalmente: {e_reload}")

    # Se o reload não for suficiente, a URL /atendimento recompõe o mesmo
    # frameset sem refazer as credenciais da sessão já autenticada.
    if not atualizou:
        try:
            page.goto(
                PORTAL_URL_ATENDIMENTO,
                wait_until="domcontentloaded",
                timeout=60000,
            )
            print("[PORTAL] Tela /atendimento reaberta após falha do reload.")
        except Exception as e_goto:
            print(f"[PORTAL] Não consegui reabrir /atendimento: {e_goto}")

    limite = time.monotonic() + 30.0
    ultimo_erro = ""
    while time.monotonic() < limite:
        try:
            frame = frame_principal_portal(page)
            campo = frame.get_by_role("textbox", name="Consultar requisição...")
            if campo.is_visible(timeout=1200):
                print(
                    "[PORTAL] Campo 'Consultar requisição...' confirmado após a "
                    "atualização. O robô tentará abrir a fila novamente."
                )
                page.wait_for_timeout(1200)
                return True
        except Exception as e:
            ultimo_erro = str(e)
        page.wait_for_timeout(500)

    print(
        "[PORTAL] A tela principal não ficou estável após a atualização. "
        f"Último detalhe: {ultimo_erro or 'sem detalhe adicional'}"
    )
    return False


def acessar_fila_suporte_venda_portal(page):
    """
    Abre a fila com recuperação automática do carregamento parcial do PORTAL.

    TESTE 22: se o portal de atendimento mostrar apenas parte do frameset (por exemplo,
    o campo de consulta existe, mas ATENDIMENTO/frameArvoreDemo ainda não carregou),
    executa um F5 controlado e repete a navegação. Isso evita encerrar a rodada
    só porque a primeira renderização do PORTAL ficou incompleta.
    """
    ultimo_erro = None

    for tentativa in range(1, 4):
        try:
            return _acessar_fila_suporte_venda_portal_uma_vez(page)
        except Exception as e:
            ultimo_erro = e

            if tentativa >= 3:
                break

            print(
                f"[PORTAL] Não foi possível materializar a fila na tentativa "
                f"{tentativa}/3. Vou recuperar a tela antes de tentar novamente."
            )
            recuperou = _recarregar_portal_quando_incompleto(
                page,
                tentativa=tentativa,
                motivo=str(e),
            )

            if not recuperou:
                print(
                    "[PORTAL] A recuperação não foi confirmada, mas ainda farei "
                    "uma nova tentativa de navegação pela sessão atual."
                )

    raise RuntimeError(
        "Não foi possível acessar a fila Equipe de Operações - Faturamento mesmo "
        f"após recuperação automática do carregamento inicial. Último erro: {ultimo_erro}"
    )


def _numeros_requisicoes_pagina_portal(resultado):
    """Retorna os números de requisição materializados na página atual do grid."""
    numeros = []
    links = resultado.get_by_role("link", name=re.compile(r"^\s*\d{7}\s*$"))
    try:
        quantidade = links.count()
    except Exception:
        quantidade = 0

    for indice in range(quantidade):
        try:
            numero = str(links.nth(indice).inner_text(timeout=2500) or "").strip()
            if re.fullmatch(r"\d{7}", numero) and numero not in numeros:
                numeros.append(numero)
        except Exception:
            continue
    return numeros


def _assinatura_pagina_portal(resultado):
    """Cria uma assinatura simples para detectar quando a página do grid mudou."""
    numeros = _numeros_requisicoes_pagina_portal(resultado)
    if numeros:
        return "|".join(numeros)

    try:
        texto = str(resultado.locator("body").inner_text(timeout=3000) or "")
    except Exception:
        texto = ""
    return normalizar_texto_busca(texto)[:600]


def _esperar_troca_pagina_portal(page, resultado, assinatura_anterior, timeout_ms=10000):
    limite = time.monotonic() + max(1.0, timeout_ms / 1000.0)
    while time.monotonic() < limite:
        page.wait_for_timeout(300)
        try:
            atual = _assinatura_pagina_portal(resultado)
        except Exception:
            atual = ""
        if atual and atual != assinatura_anterior:
            return True
    return False


def _clicar_candidato_proxima_pagina_portal(page, resultado, candidato, assinatura_anterior, descricao):
    try:
        if candidato.count() <= 0:
            return False
        total = min(candidato.count(), 8)
        for i in range(total):
            item = candidato.nth(i)
            try:
                if not item.is_visible(timeout=700):
                    continue
            except Exception:
                continue

            try:
                disabled = item.get_attribute("disabled")
                aria_disabled = str(item.get_attribute("aria-disabled") or "").lower()
                classe = str(item.get_attribute("class") or "").lower()
                if disabled is not None or aria_disabled == "true" or "disabled" in classe:
                    continue
            except Exception:
                pass

            try:
                item.scroll_into_view_if_needed(timeout=2500)
            except Exception:
                pass

            try:
                item.click(timeout=5000)
            except Exception:
                continue

            if _esperar_troca_pagina_portal(page, resultado, assinatura_anterior):
                print(f"[PORTAL] Paginação: avançou usando {descricao}.")
                return True
        return False
    except Exception:
        return False


def _ler_indicador_paginacao_portal(page):
    """Tenta ler o indicador visual do portal de atendimento, ex.: 'Pág. 2 de 5'."""
    try:
        portalmain = frame_portalmain(frame_principal_portal(page))
        texto = str(portalmain.locator("body").inner_text(timeout=3000) or "")
        achado = re.search(r"P[aá]g\.?\s*(\d+)\s*de\s*(\d+)", texto, re.I)
        if achado:
            return int(achado.group(1)), int(achado.group(2))
    except Exception:
        pass
    return None, None


def avancar_pagina_resultado_portal(page, resultado, pagina_atual):
    """
    Avança uma página no grid legado do PORTAL.

    TESTE 04:
    O Codegen real do portal de atendimento mostrou que o paginador NÃO fica dentro de
    frameResultadosDemo. Ele é exposto diretamente no PortalPainelDemo e, ao avançar, o botão
    possui nome acessível correspondente à página de destino: na página 1,
    get_by_role('img', name='2'); na página 2, name='3'; etc.

    A troca só é aceita se a assinatura das requisições materializadas no
    frameResultadosDemo realmente mudar.
    """
    assinatura_anterior = _assinatura_pagina_portal(resultado)
    proxima = str(int(pagina_atual) + 1)

    atual_ui, total_ui = _ler_indicador_paginacao_portal(page)
    if atual_ui is not None and total_ui is not None:
        print(f"[PORTAL] Paginação visual: Pág. {atual_ui} de {total_ui}.")
        if atual_ui >= total_ui:
            return False
        # Se o indicador divergir da nossa contagem, o destino correto é o
        # próximo número exibido pela própria tela.
        proxima = str(atual_ui + 1)

    # ESTRATÉGIA PRINCIPAL: exatamente o frame/controle gravado pelo Playwright
    # no mapeamento manual enviado pelo usuário.
    try:
        portalmain = frame_portalmain(frame_principal_portal(page))
        candidato_codegen = portalmain.get_by_role("img", name=proxima, exact=True)
        if _clicar_candidato_proxima_pagina_portal(
            page,
            resultado,
            candidato_codegen,
            assinatura_anterior,
            f"controle PortalPainelDemo img name={proxima}",
        ):
            return True
    except Exception as e:
        print(f"[PORTAL] Controle PortalPainelDemo para página {proxima} não acionado: {e}")

    # Fallback: setas/controles técnicos também no PortalPainelDemo.
    try:
        portalmain = frame_portalmain(frame_principal_portal(page))
        candidatos_portalmain = [
            (
                portalmain.locator(
                    "img[title*='Próx'],img[title*='Prox'],"
                    "input[title*='Próx'],input[title*='Prox'],"
                    "img[alt*='Próx'],img[alt*='Prox'],"
                    "input[alt*='Próx'],input[alt*='Prox']"
                ),
                "controle PortalPainelDemo title/alt Próxima",
            ),
            (
                portalmain.locator(
                    "[id*='prox' i],[name*='prox' i],[onclick*='prox' i],"
                    "[id*='next' i],[name*='next' i],[onclick*='next' i],"
                    "img[src*='prox' i],img[src*='next' i],"
                    "img[src*='direita' i],img[src*='right' i]"
                ),
                "controle técnico PortalPainelDemo prox/next",
            ),
        ]
        for candidato, descricao in candidatos_portalmain:
            if _clicar_candidato_proxima_pagina_portal(
                page, resultado, candidato, assinatura_anterior, descricao
            ):
                return True
    except Exception:
        pass

    # Último fallback: mantém as estratégias antigas dentro do frameResultadosDemo,
    # caso outra versão do PORTAL renderize o paginador ali.
    padrao_texto = re.compile(
        r"^\s*(?:Pr[oó]xim[oa](?:\s+p[aá]gina)?|Next(?:\s+page)?)\s*$",
        re.I,
    )
    candidatos_resultado = [
        (resultado.get_by_role("link", name=padrao_texto), "link Próxima página em frameResultadosDemo"),
        (resultado.get_by_role("button", name=padrao_texto), "botão Próxima página em frameResultadosDemo"),
    ]
    for candidato, descricao in candidatos_resultado:
        if _clicar_candidato_proxima_pagina_portal(
            page, resultado, candidato, assinatura_anterior, descricao
        ):
            return True

    return False


def _coletar_linhas_atividade_pagina_portal(resultado):
    """Coleta número, atividade e diagnóstico das linhas da página atual."""
    registros = []
    ignoradas_produto = []
    outras_fora_escopo = []
    linhas_nao_lidas = []

    numeros = _numeros_requisicoes_pagina_portal(resultado)
    for numero in numeros:
        texto_linha = ""
        try:
            linha_id = resultado.locator(f"#tr{numero}")
            if linha_id.count() > 0:
                texto_linha = linha_id.first.inner_text(timeout=4000)
        except Exception:
            pass

        if not texto_linha:
            try:
                link = resultado.get_by_role("link", name=numero, exact=True).first
                linha = link.locator("xpath=ancestor::tr[1]")
                texto_linha = linha.inner_text(timeout=4000)
            except Exception:
                texto_linha = ""

        atividade = classificar_atividade_suporte_venda(texto_linha)
        if atividade:
            registros.append((numero, atividade))
            continue

        txt_norm = normalizar_texto_busca(texto_linha)
        if re.search(
            r"devolucao\s+e\s+refaturamento\s*\(\s*produto\s*\)",
            txt_norm,
            re.I,
        ):
            ignoradas_produto.append(numero)
        elif texto_linha.strip():
            outras_fora_escopo.append(numero)
        else:
            linhas_nao_lidas.append(numero)

    return registros, ignoradas_produto, outras_fora_escopo, linhas_nao_lidas

def capturar_requisicoes_suporte_venda(page):
    """
    Mapeia TODAS as páginas abertas da fila antes de apresentar a seleção.

    Atividades permitidas:
    - Nota de Credito
    - Devolução e Refaturamento (Locação)

    "Devolução e Refaturamento (Produto)" é explicitamente ignorada.
    """
    global ATIVIDADE_POR_REQUISICAO, PAGINA_POR_REQUISICAO
    ATIVIDADE_POR_REQUISICAO = {}
    PAGINA_POR_REQUISICAO = {}

    _, resultado = acessar_fila_suporte_venda_portal(page)

    # Ordena pelo cabeçalho Atividade como no mapeamento manual. A partir daqui
    # percorremos todas as páginas, em vez de considerar apenas a página visível.
    try:
        cabecalho = resultado.get_by_text("Atividade", exact=True).first
        if cabecalho.is_visible(timeout=2500):
            cabecalho.click(timeout=5000)
            page.wait_for_timeout(1000)
    except Exception:
        pass

    requisicoes = []
    ignoradas_produto = []
    outras_fora_escopo = set()
    linhas_nao_lidas = set()
    assinaturas_vistas = set()
    pagina_atual = 1

    print("[PORTAL] Iniciando mapeamento completo das páginas da fila...")

    while pagina_atual <= MAX_PAGINAS_PORTAL:
        assinatura = _assinatura_pagina_portal(resultado)
        if assinatura in assinaturas_vistas:
            print(
                f"[PORTAL] Paginação interrompida: a página {pagina_atual} repetiu "
                "uma página já visitada."
            )
            break
        assinaturas_vistas.add(assinatura)

        registros, produtos, outras, nao_lidas = _coletar_linhas_atividade_pagina_portal(resultado)
        ignoradas_produto.extend(produtos)
        outras_fora_escopo.update(outras)
        linhas_nao_lidas.update(nao_lidas)

        encontrados_pagina = 0
        for numero, atividade in registros:
            if numero not in ATIVIDADE_POR_REQUISICAO:
                requisicoes.append(numero)
                encontrados_pagina += 1
            ATIVIDADE_POR_REQUISICAO[numero] = atividade
            PAGINA_POR_REQUISICAO[numero] = pagina_atual

        total_linhas_pagina = len(_numeros_requisicoes_pagina_portal(resultado))
        print(
            f"[PORTAL] Página {pagina_atual}: {total_linhas_pagina} requisição(ões) "
            f"visível(is); {encontrados_pagina} nova(s) dentro do escopo."
        )

        if not avancar_pagina_resultado_portal(page, resultado, pagina_atual):
            print(f"[PORTAL] Página {pagina_atual}: não há próxima página detectável.")
            break

        pagina_atual += 1
        page.wait_for_timeout(800)

    requisicoes = list(dict.fromkeys(requisicoes))
    ignoradas_produto = list(dict.fromkeys(ignoradas_produto))

    qtd_nc = sum(
        1 for n in requisicoes
        if ATIVIDADE_POR_REQUISICAO.get(n) == ATIVIDADE_NOTA_CREDITO
    )
    qtd_locacao = sum(
        1 for n in requisicoes
        if ATIVIDADE_POR_REQUISICAO.get(n) == ATIVIDADE_DEVOLUCAO_REFATURAMENTO
    )

    print(
        f"[PORTAL] Mapeamento concluído em {pagina_atual} página(s): "
        f"{len(requisicoes)} requisição(ões) no escopo."
    )
    print(f"[PORTAL] Nota de Credito: {qtd_nc}")
    print(f"[PORTAL] Devolução e Refaturamento (Locação): {qtd_locacao}")

    for numero in requisicoes:
        atividade = ATIVIDADE_POR_REQUISICAO.get(numero, "")
        pagina_num = PAGINA_POR_REQUISICAO.get(numero, "?")
        print(f"[PORTAL] {numero} -> {atividade} [página {pagina_num}]")

    if ignoradas_produto:
        print(
            f"[PORTAL] {len(ignoradas_produto)} requisição(ões) de "
            "'Devolução e Refaturamento (Produto)' ignorada(s) por estar(em) "
            "fora do escopo deste robô."
        )

    classificadas = set(requisicoes) | set(ignoradas_produto)
    outras_fora_escopo.difference_update(classificadas)
    linhas_nao_lidas.difference_update(classificadas | outras_fora_escopo)
    print(f"[PORTAL] Outras atividades fora do escopo: {len(outras_fora_escopo)}")
    if outras_fora_escopo:
        print("[PORTAL] Requisições de outras atividades: " + ", ".join(sorted(outras_fora_escopo)))
    if linhas_nao_lidas:
        print(f"[PORTAL] Atividade não pôde ser lida em {len(linhas_nao_lidas)} linha(s): " + ", ".join(sorted(linhas_nao_lidas)))
    total_mapeado = len(classificadas | outras_fora_escopo | linhas_nao_lidas)
    print(f"[PORTAL] Resumo da triagem: {total_mapeado} registros únicos; {len(requisicoes)} no escopo; "
          f"{len(ignoradas_produto)} Produto; {len(outras_fora_escopo)} outras atividades; "
          f"{len(linhas_nao_lidas)} sem leitura da atividade.")

    return requisicoes

def interpretar_selecao_requisicoes(entrada, total_requisicoes):
    """
    Interpreta uma seleção informada pelo usuário usando posições da lista.

    Formatos aceitos:
    - 1,3,5
    - 1-4
    - 1-3,7,10

    Retorna uma lista de índices base zero, sem duplicidades e na ordem
    original em que as requisições aparecem na fila.
    """
    entrada = str(entrada or "").strip().replace(" ", "")

    if not entrada:
        raise ValueError("A seleção não pode ficar vazia.")

    partes = entrada.split(",")
    posicoes = set()

    for parte in partes:
        if not parte:
            raise ValueError("Há uma posição vazia na seleção.")

        if "-" in parte:
            if parte.count("-") != 1:
                raise ValueError(f"Intervalo inválido: {parte}")

            inicio_txt, fim_txt = parte.split("-", 1)
            if not inicio_txt.isdigit() or not fim_txt.isdigit():
                raise ValueError(f"Intervalo inválido: {parte}")

            inicio = int(inicio_txt)
            fim = int(fim_txt)

            if inicio > fim:
                raise ValueError(f"O intervalo {parte} está invertido.")

            if inicio < 1 or fim > total_requisicoes:
                raise ValueError(
                    f"O intervalo {parte} está fora da lista, que vai de 1 a {total_requisicoes}."
                )

            posicoes.update(range(inicio, fim + 1))

        else:
            if not parte.isdigit():
                raise ValueError(f"Posição inválida: {parte}")

            posicao = int(parte)
            if posicao < 1 or posicao > total_requisicoes:
                raise ValueError(
                    f"A posição {posicao} está fora da lista, que vai de 1 a {total_requisicoes}."
                )

            posicoes.add(posicao)

    if not posicoes:
        raise ValueError("Nenhuma requisição válida foi selecionada.")

    return [posicao - 1 for posicao in sorted(posicoes)]

def confirmar_requisicoes_selecionadas(requisicoes_selecionadas):
    """Exibe as requisições escolhidas e solicita confirmação explícita."""
    print("\n========================================")
    print("CONFIRMAÇÃO DA EXECUÇÃO")
    print("========================================")
    print(f"Serão processadas {len(requisicoes_selecionadas)} requisição(ões):")

    for requisicao in requisicoes_selecionadas:
        atividade = ATIVIDADE_POR_REQUISICAO.get(str(requisicao).strip(), "")
        sufixo = f" - {atividade}" if atividade else ""
        print(f"- {requisicao}{sufixo}")

    while True:
        confirmacao = input("\nConfirma o início? [S/N]: ").strip().upper()

        if confirmacao in {"S", "SIM"}:
            return True

        if confirmacao in {"N", "NAO", "NÃO"}:
            return False

        print("[SELEÇÃO] Resposta inválida. Digite S para confirmar ou N para voltar.")

def selecionar_requisicoes_para_processar(requisicoes_encontradas):
    """
    Permite escolher quais requisições da fila serão processadas.

    Retorno:
    - (lista_selecionada, False): execução confirmada;
    - ([], True): execução cancelada pelo usuário.

    A lista original não é alterada. A ordem da fila é sempre preservada.
    """
    requisicoes_encontradas = list(requisicoes_encontradas or [])

    if not requisicoes_encontradas:
        return [], False

    while True:
        print("\n========================================")
        print("REQUISIÇÕES LOCALIZADAS NA FILA")
        print("========================================")

        for indice, requisicao in enumerate(requisicoes_encontradas, start=1):
            atividade = ATIVIDADE_POR_REQUISICAO.get(str(requisicao).strip(), "")
            sufixo = f" - {atividade}" if atividade else ""
            print(f"[{indice}] {requisicao}{sufixo}")

        print(f"\nTotal localizado: {len(requisicoes_encontradas)} requisição(ões)")
        print("\nO que deseja fazer?")
        print("[1] Processar todas")
        print("[2] Selecionar requisições")
        print("[3] Cancelar execução")

        opcao = input("\nEscolha uma opção: ").strip()

        if opcao == "1":
            selecionadas = requisicoes_encontradas.copy()

            if confirmar_requisicoes_selecionadas(selecionadas):
                return selecionadas, False

            print("[SELEÇÃO] Confirmação recusada. Voltando ao menu de seleção.")
            continue

        if opcao == "2":
            print("\nDigite as posições desejadas.")
            print("Exemplos: 1,3,5 | 1-4 | 1-3,7,10")

            entrada = input("Seleção: ").strip()

            try:
                indices = interpretar_selecao_requisicoes(
                    entrada=entrada,
                    total_requisicoes=len(requisicoes_encontradas),
                )
                selecionadas = [requisicoes_encontradas[i] for i in indices]
            except ValueError as e:
                print(f"[SELEÇÃO] {e}")
                continue

            if confirmar_requisicoes_selecionadas(selecionadas):
                return selecionadas, False

            print("[SELEÇÃO] Confirmação recusada. Voltando ao menu de seleção.")
            continue

        if opcao == "3":
            print("\n[CANCELADO] Execução cancelada pelo usuário antes de iniciar o processamento.")
            return [], True

        print("[SELEÇÃO] Opção inválida. Digite 1, 2 ou 3.")

def validar_requisicao_portal_aberta(frame, requisicao, timeout=30000):
    """
    Confirma o número da requisição antes de ler, anexar ou concluir o PORTAL.

    Correção AV13:
    - o título da requisição fica dentro do iframe PortalPainelDemo;
    - a AV12 procurava somente no framePrincipalDemo e bloqueava requisições válidas;
    - lemos o texto do PortalPainelDemo e mantemos o framePrincipalDemo apenas como fallback;
    - a proteção contra cruzamento continua ativa: só aceita o número atual.
    """
    requisicao = str(requisicao).strip()
    padrao_cabecalho = re.compile(
        rf"(?<!\d){re.escape(requisicao)}\s*-",
        re.I,
    )

    fontes_texto = [
        ("PortalPainelDemo", frame_portalmain(frame).locator("body")),
        ("framePrincipalDemo", frame.locator("body")),
    ]
    limite = time.monotonic() + max(1.0, float(timeout) / 1000.0)
    cabecalhos_observados = []
    ultimo_erro = ""

    while time.monotonic() < limite:
        for nome_fonte, corpo in fontes_texto:
            restante_ms = max(250, int((limite - time.monotonic()) * 1000))

            try:
                texto_tela = str(
                    corpo.inner_text(timeout=min(2000, restante_ms)) or ""
                )
            except Exception as e:
                ultimo_erro = f"{nome_fonte}: {e}"
                continue

            if padrao_cabecalho.search(texto_tela):
                print(
                    f"[PORTAL] Cabeçalho confirmado para a requisição "
                    f"{requisicao} no {nome_fonte}."
                )
                return True

            for numero, titulo in re.findall(
                r"(?<!\d)(\d{7})\s*-\s*([^\r\n]*)",
                texto_tela,
            ):
                observado = f"{numero} - {titulo.strip()}".rstrip(" -")
                if observado and observado not in cabecalhos_observados:
                    cabecalhos_observados.append(observado)

        time.sleep(0.35)

    vistos = "; ".join(cabecalhos_observados[-5:]) or "nenhum cabeçalho de 7 dígitos"
    detalhe = f" Último detalhe técnico: {ultimo_erro}" if ultimo_erro else ""
    raise RuntimeError(
        f"O PORTAL não confirmou a requisição {requisicao} dentro do PortalPainelDemo. "
        f"Cabeçalho(s) observado(s): {vistos}. A leitura/anexação foi "
        f"bloqueada para evitar cruzamento entre requisições.{detalhe}"
    )


def _localizar_link_requisicao_na_pagina_portal(resultado, requisicao):
    requisicao = str(requisicao).strip()

    try:
        linha = resultado.locator(f"#tr{requisicao}").first
        if linha.count() > 0:
            candidato = linha.get_by_role("link", name=requisicao, exact=True).first
            if candidato.count() > 0 and candidato.is_visible(timeout=1200):
                return candidato
    except Exception:
        pass

    try:
        candidato = resultado.get_by_role("link", name=requisicao, exact=True).first
        if candidato.count() > 0 and candidato.is_visible(timeout=1200):
            return candidato
    except Exception:
        pass

    return None


def abrir_requisicao_portal(page, requisicao):
    """
    Abre uma requisição PORTAL de forma robusta.

    TESTE 13:
    - o inventário inicial continua sendo feito pela fila/paginação;
    - depois que o usuário seleciona as requisições, a abertura passa a usar
      prioritariamente o campo global "Consultar requisição...";
    - isso evita depender da página/ordenação atual do grid, que pode mudar
      entre o inventário e o processamento;
    - somente se a consulta direta falhar usamos a fila como fallback.
    """
    requisicao = str(requisicao).strip()

    # TESTE 35: um download de planilha da requisição anterior pode deixar
    # modal/overlay do PORTAL aberto e bloquear a consulta da próxima requisição.
    fechar_modal_residual_portal(
        page,
        motivo=f"antes de abrir a requisição {requisicao}",
        timeout=5000,
    )

    # -------------------------------------------------------------
    # CAMINHO PRINCIPAL: CONSULTA DIRETA PELO NÚMERO DA REQUISIÇÃO
    # -------------------------------------------------------------
    ultimo_erro_consulta = None
    for tentativa in range(1, 3):
        try:
            frame = frame_principal_portal(page)
            busca = frame.get_by_role("textbox", name="Consultar requisição...")

            if not busca.is_visible(timeout=5000):
                raise RuntimeError("Campo 'Consultar requisição...' não está visível.")

            print(
                f"[PORTAL] Abrindo requisição {requisicao} pela consulta direta "
                f"(tentativa {tentativa}/2)..."
            )

            busca.click(timeout=8000)
            busca.fill("")
            busca.fill(requisicao)
            busca.press("Enter")
            page.wait_for_timeout(TEMPO_PORTAL_REQUISICAO)

            validar_requisicao_portal_aberta(
                frame,
                requisicao,
                timeout=30000,
            )

            print(
                f"[PORTAL] Requisição {requisicao} aberta com sucesso pelo campo "
                "'Consultar requisição...'."
            )
            return frame

        except Exception as e:
            ultimo_erro_consulta = e
            print(
                f"[PORTAL] Consulta direta da requisição {requisicao} não foi "
                f"confirmada na tentativa {tentativa}/2. Detalhe: {e}"
            )

            if tentativa == 1:
                try:
                    page.goto(
                        PORTAL_URL_ATENDIMENTO,
                        wait_until="domcontentloaded",
                        timeout=60000,
                    )
                    page.wait_for_timeout(2500)
                except Exception as nav_e:
                    print(
                        "[PORTAL] Não foi possível recarregar a tela principal "
                        f"antes da segunda tentativa. Detalhe: {nav_e}"
                    )

    # -------------------------------------------------------------
    # FALLBACK: REENCONTRAR NA FILA/PAGINAÇÃO
    # -------------------------------------------------------------
    print(
        f"[PORTAL] Consulta direta não abriu {requisicao}; usando a fila como "
        "fallback."
    )
    frame, resultado = acessar_fila_suporte_venda_portal(page)

    # Reproduz o mesmo estado de ordenação usado no inventário inicial.
    try:
        cabecalho = resultado.get_by_text("Atividade", exact=True).first
        if cabecalho.is_visible(timeout=2500):
            cabecalho.click(timeout=5000)
            page.wait_for_timeout(1000)
    except Exception:
        pass

    pagina_prevista = PAGINA_POR_REQUISICAO.get(requisicao)
    if pagina_prevista:
        print(
            f"[PORTAL] Requisição {requisicao} foi mapeada inicialmente na "
            f"página {pagina_prevista}. Procurando no grid de fallback..."
        )

    assinaturas_vistas = set()
    pagina_atual = 1
    link_requisicao = None

    while pagina_atual <= MAX_PAGINAS_PORTAL:
        assinatura = _assinatura_pagina_portal(resultado)
        if assinatura in assinaturas_vistas:
            break
        assinaturas_vistas.add(assinatura)

        link_requisicao = _localizar_link_requisicao_na_pagina_portal(
            resultado, requisicao
        )
        if link_requisicao is not None:
            print(
                f"[PORTAL] Requisição {requisicao} localizada na página "
                f"{pagina_atual} do grid de fallback."
            )
            break

        if not avancar_pagina_resultado_portal(page, resultado, pagina_atual):
            break

        pagina_atual += 1
        page.wait_for_timeout(650)

    if link_requisicao is None:
        raise RuntimeError(
            f"Requisição {requisicao} não pôde ser aberta pela consulta direta "
            "nem reencontrada na paginação da fila. "
            f"Último erro da consulta direta: {ultimo_erro_consulta}"
        )

    link_requisicao.scroll_into_view_if_needed(timeout=10000)
    link_requisicao.click(timeout=20000)
    page.wait_for_timeout(TEMPO_PORTAL_REQUISICAO)

    validar_requisicao_portal_aberta(
        frame,
        requisicao,
        timeout=30000,
    )

    print(
        f"[PORTAL] Requisição {requisicao} aberta pelo fallback da fila "
        "Operações de Faturamento."
    )
    return frame

def _normalizar_valor_status_portal(valor):
    """Converte somente valores conhecidos do campo visual 'Status:' do PORTAL."""
    valor = re.sub(r"\s+", " ", str(valor or "")).strip()
    if not valor:
        return ""

    # Se algum container trouxer outro rótulo depois do status, corta ali.
    valor = re.split(
        r"\s+(?:Prazo|Atividade|Tempo\s+(?:Restante|ativo))\s*:",
        valor,
        maxsplit=1,
        flags=re.I,
    )[0].strip()

    estados = {
        "em aberto": "Em Aberto",
        "em andamento": "Em andamento",
        "em resolucao": "Em resolução",
        "cancelada": "Cancelada",
        "cancelado": "Cancelada",
        "encerrada": "Encerrada",
        "encerrado": "Encerrada",
        "concluida": "Concluída",
        "concluido": "Concluída",
        "finalizada": "Finalizada",
        "finalizado": "Finalizada",
    }
    return estados.get(normalizar_texto_busca(valor), "")


def _ler_candidatos_status_visiveis_portal(frame):
    """Lê apenas elementos VISÍVEIS cujo texto começa por 'Status:' no PortalPainelDemo."""
    portalmain = frame_portalmain(frame)
    return portalmain.locator("body").evaluate(r"""body => {
        const visivel = el => {
            const r = el.getBoundingClientRect();
            const css = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && css.display !== 'none' &&
                   css.visibility !== 'hidden' && !el.closest('[aria-hidden="true"]');
        };
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const todos = [...body.querySelectorAll('td,th,div,span,p,b,strong,label')];
        const candidatos = todos.filter(el => {
            if (!visivel(el) || el.closest('textarea,[contenteditable="true"]')) return false;
            for (let p = el; p && p !== body; p = p.parentElement) {
                const marca = `${p.id || ''} ${typeof p.className === 'string' ? p.className : ''}`;
                if (/histor|descri|coment|mensagem/i.test(marca)) return false;
            }
            const texto = norm(el.innerText || el.textContent || '');
            // Exige valor depois dos dois-pontos. Um <b>Status:</b> isolado não basta;
            // seu pai/célula, que contém também o valor, será considerado.
            return texto.length <= 180 && /^Status\s*:\s*\S/i.test(texto);
        });
        // Mantém os candidatos mais internos para não contar o mesmo status em
        // vários containers ancestrais.
        return candidatos
            .filter(el => !candidatos.some(outro => outro !== el && el.contains(outro)))
            .map(el => ({
                tag: el.tagName,
                id: el.id || '',
                className: typeof el.className === 'string' ? el.className : '',
                texto: norm(el.innerText || el.textContent || '')
            }));
    }""")


def ler_status_requisicao_portal(frame, requisicao, timeout=8000):
    """Confirma o Status visual da requisição em duas leituras consecutivas.

    TESTE 46:
    - a requisição correta é validada separadamente pelo cabeçalho;
    - o campo Status é localizado de forma independente no PortalPainelDemo;
    - nunca procura 'encerrada/cancelada' como palavra solta no body;
    - zero ou múltiplos candidatos visíveis ficam inconclusivos por segurança.
    """
    requisicao = str(requisicao).strip()
    limite = time.monotonic() + max(0, timeout) / 1000.0
    anterior = None
    resultado = {
        "status": "",
        "evidencia": [],
        "motivo": "campo visual 'Status:' não localizado",
    }

    while True:
        try:
            candidatos = _ler_candidatos_status_visiveis_portal(frame) or []
            # Deduplica evidências DOM equivalentes sem transformar dois campos
            # de status diferentes em um só.
            unicos = []
            chaves = set()
            for c in candidatos:
                chave = (
                    str(c.get("texto", "")).strip(),
                    str(c.get("id", "")).strip(),
                    str(c.get("tag", "")).strip(),
                )
                if chave not in chaves:
                    chaves.add(chave)
                    unicos.append(c)
            candidatos = unicos
            resultado["evidencia"] = candidatos

            if len(candidatos) == 1:
                texto_status = str(candidatos[0].get("texto", "") or "").strip()
                m = re.match(r"^\s*Status\s*:\s*(.+?)\s*$", texto_status, re.I)
                status = _normalizar_valor_status_portal(m.group(1) if m else "")

                if status:
                    assinatura = (status, texto_status)
                    if assinatura == anterior:
                        resultado.update(status=status, motivo="")
                        return resultado
                    anterior = assinatura
                    resultado["motivo"] = "aguardando segunda leitura igual do campo Status"
                else:
                    anterior = None
                    resultado["motivo"] = (
                        f"valor do campo Status não reconhecido: {texto_status!r}"
                    )
            elif len(candidatos) > 1:
                anterior = None
                resultado["motivo"] = (
                    f"foram encontrados {len(candidatos)} campos Status visíveis; "
                    "leitura ambígua"
                )
            else:
                anterior = None
                resultado["motivo"] = "campo visual 'Status:' não localizado"

        except Exception as erro:
            anterior = None
            resultado["motivo"] = f"não foi possível ler o campo Status: {erro}"

        if time.monotonic() >= limite:
            return resultado
        time.sleep(0.35)


class StatusPORTALIndeterminado(RuntimeError):
    """Bloqueia SAP/escrita PORTAL quando o status visual não pode ser comprovado."""


def confirmar_status_cabecalho_portal(page, frame, requisicao, timeout=8000):
    """Confirma requisição + Status em regiões independentes da tela PORTAL."""
    requisicao = str(requisicao).strip()

    # Usa a validação de cabeçalho do TESTE 44, que já estava funcionando.
    validar_requisicao_portal_aberta(
        frame,
        requisicao,
        timeout=min(max(int(timeout), 2500), 8000),
    )

    leitura = ler_status_requisicao_portal(frame, requisicao, timeout=timeout)
    if leitura["status"]:
        evidencia = leitura["evidencia"][0] if leitura["evidencia"] else {}
        texto_elemento = str(evidencia.get("texto", "") or "").strip()
        print(f"[PORTAL STATUS] Requisição confirmada: {requisicao}")
        print(f"[PORTAL STATUS] Elemento localizado: {texto_elemento!r}")
        print(f"[PORTAL STATUS] Valor normalizado: {leitura['status']}")
        print("[PORTAL STATUS] 2 leituras consecutivas iguais -> confirmado")
        return leitura["status"]

    mensagem = f"Status PORTAL inconclusivo: {leitura['motivo']}"
    print(f"[PORTAL] Requisição {requisicao}: {mensagem}. Nenhuma nova escrita autorizada.")

    evidencias = []
    try:
        pasta = criar_pasta_requisicao(requisicao)
        base = os.path.join(
            pasta,
            f"status_portal_indeterminado_{requisicao}_{ID_RODADA}_{datetime.now():%H%M%S_%f}",
        )
        with open(base + ".txt", "w", encoding="utf-8") as arquivo:
            arquivo.write(
                f"Requisição: {requisicao}\n"
                f"{mensagem}\n"
                "Fonte: campo visual Status no PortalPainelDemo, após validação separada do cabeçalho\n"
                f"Candidatos: {leitura['evidencia']!r}\n"
            )
        evidencias.append(os.path.basename(base + ".txt"))
        try:
            page.screenshot(path=base + ".png", full_page=True)
            evidencias.append(os.path.basename(base + ".png"))
        except Exception as erro_print:
            print(f"[PORTAL] Não foi possível salvar print do status: {erro_print}")
    except Exception as erro_evidencia:
        print(f"[PORTAL] Não foi possível salvar evidência do status: {erro_evidencia}")

    if evidencias:
        mensagem += " | Evidências: " + ", ".join(evidencias)
    raise StatusPORTALIndeterminado(mensagem)


def status_requisicao_portal_nao_acionavel(status):
    """Estados confirmados em que o robô não deve escrever no PORTAL/SAP."""
    return normalizar_texto_busca(status) in {
        "cancelada", "encerrada", "concluida", "finalizada"
    }


def tentar_assumir_responsabilidade_portal(page, frame, requisicao):
    """Assume quando o botão existe; quando não existe, segue normalmente.

    TESTE 46:
    - status é confirmado pelo campo visual local antes de qualquer escrita;
    - ausência de 'Assumir Responsabilidade' não é erro operacional;
    - após clicar, Em andamento/Em resolução, desaparecimento do botão ou
      alerta de carregamento confirmam avanço suficiente.
    """
    botao = frame.get_by_role("img", name="Assumir Responsabilidade").first
    inicio = len(PORTAL_DIALOG_STATE.get("eventos", []))

    try:
        status_antes = confirmar_status_cabecalho_portal(page, frame, requisicao)
        if status_requisicao_portal_nao_acionavel(status_antes):
            return False, f"PORTAL_NAO_ACIONAVEL::{status_antes or 'estado encerrado'}"

        disponivel = botao.count() > 0 and botao.is_visible(timeout=1200)
        if not disponivel:
            mensagem = (
                "Assumir Responsabilidade não está disponível; seguindo normalmente "
                "conforme regra operacional."
            )
            if status_antes:
                mensagem += f" Status atual: {status_antes}."
            print(f"[PORTAL] {mensagem}")
            return True, mensagem

        botao.click(timeout=12000)
        page.wait_for_timeout(700)
        verificar_dialogs_portal_desde(inicio)
        clicar_ok_portal(frame, timeout=2500)

        limite = time.monotonic() + 20.0
        status_depois = ""
        botao_sumiu = False
        alerta_carregamento = False

        while time.monotonic() < limite:
            verificar_dialogs_portal_desde(inicio)
            eventos = PORTAL_DIALOG_STATE.get("eventos", [])[inicio:]
            alerta_carregamento = any(
                e.get("tipo") == "alert" and e.get("aceito") and
                re.search(
                    r"aguarde.*carregamento.*dados adicionais|responsabilidade assumida|sucesso",
                    normalizar_texto_busca(e.get("mensagem", "")),
                    re.I,
                )
                for e in eventos
            )

            leitura = ler_status_requisicao_portal(frame, requisicao, timeout=1000)
            status_depois = leitura.get("status", "")
            if status_requisicao_portal_nao_acionavel(status_depois):
                return False, f"PORTAL_NAO_ACIONAVEL::{status_depois or 'estado encerrado'}"

            try:
                botao_sumiu = not botao.count() or not botao.is_visible(timeout=700)
            except Exception:
                botao_sumiu = True

            status_norm = normalizar_texto_busca(status_depois)
            if status_norm in {"em andamento", "em resolucao"} or botao_sumiu or alerta_carregamento:
                break

            page.wait_for_timeout(500)

        # Reconfirma o status ao final. Se a tela estiver realmente ambígua,
        # preserva a barreira de segurança e não escreve às cegas.
        status_final = confirmar_status_cabecalho_portal(page, frame, requisicao)
        if status_requisicao_portal_nao_acionavel(status_final):
            return False, f"PORTAL_NAO_ACIONAVEL::{status_final}"

        if not (
            normalizar_texto_busca(status_final) in {"em andamento", "em resolucao", "em aberto"}
            or botao_sumiu
            or alerta_carregamento
        ):
            return False, (
                "Assumir Responsabilidade foi acionado, mas o PORTAL não apresentou "
                "mudança de estado confirmável dentro do tempo."
            )

        mensagem = "Responsabilidade assumida/estado operacional confirmado."
        if status_final:
            mensagem += f" Status atual: {status_final}."
        print(f"[PORTAL] {mensagem}")
        return True, mensagem

    except StatusPORTALIndeterminado:
        raise
    except Exception as erro:
        return False, f"Não foi possível confirmar responsabilidade no PORTAL: {erro}"

def preparar_requisicao_portal_para_anexo_encerramento(page, requisicao):
    """Confirma/reabre a requisição correta, fecha modais e verifica a responsabilidade."""
    restaurar_browser_portal(page)
    page.wait_for_timeout(700)

    # TESTE 35:
    # requisições originadas de planilha podem voltar do SAP com a janela
    # de seleção/download ainda aberta no PORTAL. Ela deve ser removida antes de
    # qualquer clique em Assumir Responsabilidade / Dados Adicionais / Documentos.
    fechar_modal_residual_portal(
        page,
        motivo=f"ao retornar do SAP para a requisição {requisicao}",
        timeout=7000,
    )

    frame = frame_principal_portal(page)

    try:
        validar_requisicao_portal_aberta(
            frame,
            requisicao,
            timeout=4000,
        )
        print(
            f"[PORTAL] Requisição {requisicao} permaneceu aberta após o SAP. "
            "Não será necessário reabrir pela consulta direta."
        )
    except Exception as e_validacao:
        print(
            f"[PORTAL] Requisição {requisicao} não está mais aberta após retornar do SAP. "
            "Reabrindo pela consulta direta..."
        )
        print(f"[PORTAL] Detalhe da validação anterior: {e_validacao}")

        # Caminho robusto já validado nos testes anteriores: usa o campo global
        # 'Consultar requisição...' e somente recorre à fila como fallback interno.
        frame = abrir_requisicao_portal(page, requisicao)
        page.wait_for_timeout(1200)
        validar_requisicao_portal_aberta(
            frame,
            requisicao,
            timeout=20000,
        )
        print(
            f"[PORTAL] Requisição {requisicao} reaberta e confirmada para "
            "anexo/encerramento."
        )

        fechar_modal_residual_portal(
            page,
            motivo=f"após reabrir a requisição {requisicao}",
            timeout=5000,
        )
        frame = frame_principal_portal(page)

    # Confirma que estamos de fato na tela operacional da requisição.
    # O PORTAL pode carregar o cabeçalho antes das abas/botões; por isso aguardamos
    # algum elemento funcional da tela de detalhe.
    elementos_validacao = [
        frame.get_by_role("img", name="Assumir Responsabilidade"),
        frame.get_by_role("img", name="Concluir Requisição"),
        frame_portalmain(frame).get_by_text("Documentos", exact=True),
        frame_portalmain(frame).get_by_text("Dados Adicionais", exact=True),
        frame_portalmain(frame).locator(FRAME_EDITOR_TEXTO),
    ]

    tela_detalhe_ok = False
    ultimo_erro = ""
    limite = time.monotonic() + 20
    while time.monotonic() < limite and not tela_detalhe_ok:
        for elemento in elementos_validacao:
            try:
                if elemento.first.is_visible(timeout=900):
                    tela_detalhe_ok = True
                    break
            except Exception as e:
                ultimo_erro = str(e)
        if not tela_detalhe_ok:
            page.wait_for_timeout(500)

    if not tela_detalhe_ok:
        raise Exception(
            f"A requisição {requisicao} foi confirmada pelo cabeçalho, mas a tela "
            f"operacional não terminou de carregar. Detalhe: {ultimo_erro}"
        )

    print(
        f"[PORTAL] Requisição {requisicao} confirmada na tela de detalhes. "
        "Preparando anexo e conclusão."
    )

    fechar_modal_residual_portal(
        page,
        motivo=f"antes das ações finais da requisição {requisicao}",
        timeout=5000,
    )
    frame = frame_principal_portal(page)

    assumida, mensagem = tentar_assumir_responsabilidade_portal(page, frame, requisicao)
    if not assumida:
        raise RuntimeError(mensagem)

    # A confirmação de assumir responsabilidade também pode recriar/fechar
    # overlays. Retornamos sempre um FrameLocator atual.
    fechar_modal_residual_portal(
        page,
        motivo=f"após assumir responsabilidade da requisição {requisicao}",
        timeout=3000,
    )
    return frame_principal_portal(page)


def listar_pdfs_validos_da_requisicao(pasta_requisicao, consultas_sap=None):
    """
    Lista somente PDFs ligados explicitamente às OVs desta requisição.

    AV12 remove o fallback que anexava qualquer PDF encontrado na pasta. Cada
    arquivo precisa ser exatamente <pasta da requisição>/<OV>.pdf e ter
    assinatura PDF válida.
    """
    pdfs = []
    pasta_esperada = os.path.abspath(str(pasta_requisicao or ""))

    if consultas_sap:
        for consulta in consultas_sap:
            ordem = str(consulta.get("Ordem", "") or "").strip()
            caminho = str(consulta.get("ArquivoPDF", "") or "").strip()
            status = str(consulta.get("StatusSAP", "") or "").upper()
            caminho_esperado = os.path.abspath(
                os.path.join(pasta_esperada, f"{ordem}.pdf")
            )

            if not ordem or not caminho:
                continue
            if os.path.abspath(caminho).lower() != caminho_esperado.lower():
                print(
                    f"[PORTAL] PDF bloqueado por associação divergente. OV {ordem}; "
                    f"esperado: {caminho_esperado}; recebido: {caminho}."
                )
                continue
            if not pdf_valido(caminho):
                continue
            if status_sap_ok_para_medicao(status):
                pdfs.append(caminho)

    # Remove duplicados mantendo ordem.
    pdfs_unicos = []
    vistos = set()
    for pdf in pdfs:
        chave = os.path.abspath(pdf).lower()
        if chave not in vistos:
            vistos.add(chave)
            pdfs_unicos.append(pdf)

    return pdfs_unicos


def todas_ovs_com_pdf_ok(consultas_sap):
    """
    Segurança: evita concluir o PORTAL quando alguma OV falhou no SAP/PDF.
    """
    if not consultas_sap:
        return False, "Nenhuma OV localizada/processada para a requisição."

    pendencias = []
    for consulta in consultas_sap:
        ordem = consulta.get("Ordem", "")
        status = str(consulta.get("StatusSAP", "") or "").upper()
        caminho = str(consulta.get("ArquivoPDF", "") or "").strip()

        ok_status = status_sap_ok_para_medicao(status)
        ok_nome = bool(caminho) and os.path.basename(caminho).lower() == f"{ordem}.pdf".lower()
        ok_pdf = bool(ok_nome and pdf_valido(caminho))

        if not (ok_status and ok_pdf):
            pendencias.append(f"OV {ordem}: status={status or 'VAZIO'}, pdf={'OK' if ok_pdf else 'NÃO'}")

    if pendencias:
        return False, "Pendências para conclusão: " + "; ".join(pendencias)

    return True, "Todas as OVs possuem PDF válido."


def montar_texto_encerramento_portal(requisicao, consultas_sap, pdfs):
    """Monta texto simples de conclusão para o PORTAL."""
    linhas = [TEXTO_ENCERRAMENTO_PADRAO, "", f"Requisição PORTAL: {requisicao}"]

    if consultas_sap:
        linhas.append("OVs processadas:")
        for consulta in consultas_sap:
            ordem = consulta.get("Ordem", "")
            status = consulta.get("StatusSAP", "")
            origem = identificar_origem_sap(consulta)
            nome_pdf = os.path.basename(str(consulta.get("ArquivoAnexoPORTAL") or consulta.get("ArquivoPDF", "") or ""))
            detalhe = f"- {ordem} - {status}"
            if origem:
                detalhe += f" - {origem}"
            if nome_pdf:
                detalhe += f" - {nome_pdf}"
            linhas.append(detalhe)

    linhas.append("")
    linhas.append(f"Total de arquivo(s) anexado(s): {len(pdfs)}")

    return "\n".join(linhas)


def preencher_dados_adicionais_portal(page, frame, qtd_documentos, requisicao):
    """Preenche Total e Processados por rótulo, com IDs conhecidos como fallback.

    Valida os valores e grava Prioridade=Não, conforme regra operacional."""
    total = str(int(qtd_documentos))

    print(
        f"[PORTAL] Preenchendo Dados Adicionais da requisição {requisicao} "
        "por rótulo/estrutura, sem depender de ID fixo..."
    )

    # Reobtém o frame atual após o retorno do SAP.
    frame = frame_principal_portal(page)
    validar_requisicao_portal_aberta(
        frame,
        requisicao,
        timeout=30000,
    )

    portalmain = frame_portalmain(frame)

    # Abre Dados Adicionais.
    try:
        portalmain.get_by_text(
            "Dados Adicionais",
            exact=True,
        ).click(timeout=15000)
    except Exception:
        portalmain.get_by_text(
            "Dados Adicionais"
        ).first.click(timeout=15000)

    page.wait_for_timeout(1000)

    # Usa o mesmo padrão estável do robô de Locação: FrameLocator.
    portalmain.locator(
        FRAME_DADOS_ADICIONAIS
    ).wait_for(
        state="attached",
        timeout=30000,
    )
    dados = portalmain.frame_locator(
        FRAME_DADOS_ADICIONAIS
    )

    # Em vez de esperar IDs, espera o conteúdo funcional da tela.
    dados.get_by_text(
        re.compile(r"Total\s+de\s+Documentos", re.I)
    ).first.wait_for(
        state="visible",
        timeout=30000,
    )

    dados.get_by_text(
        re.compile(r"Documentos\s+processados", re.I)
    ).first.wait_for(
        state="visible",
        timeout=30000,
    )

    print("[PORTAL] Tela de Dados Adicionais materializada pelos rótulos.")

    # Diagnóstico: mostra quais controles Dad_* existem de verdade nesta atividade.
    try:
        mapa = dados.locator("[id^='Dad_']").evaluate_all("""els => els.slice(0,100).map(el => ({
            tag: el.tagName, id: el.id || '', name: el.getAttribute('name') || '',
            type: el.getAttribute('type') || '', value: el.value || ''
        }))""")

        if mapa:
            print(
                "[PORTAL] Controles Dad_* encontrados nesta atividade: "
                + " | ".join(
                    f"{x['tag']} id='{x['id']}' name='{x['name']}' "
                    f"type='{x['type']}'"
                    for x in mapa
                )
            )
        else:
            print("[PORTAL] Nenhum controle com ID Dad_* foi encontrado no frame.")
    except Exception as e_diag:
        print(f"[PORTAL] Diagnóstico de IDs Dad_* não pôde ser concluído: {e_diag}")

    def _campo_por_rotulo(padrao_rotulo, nome_campo, ids_fallback=()):
        """
        Localiza o input pela posição estrutural em relação ao rótulo visível.
        Só depois tenta IDs conhecidos.
        """
        erros = []

        # 1) Caminho principal: texto do rótulo -> primeiro input seguinte.
        try:
            rotulo = dados.get_by_text(padrao_rotulo).first
            rotulo.wait_for(state="visible", timeout=10000)

            campo = rotulo.locator(
                "xpath=following::input[not(@type='hidden')][1]"
            )

            if campo.count() > 0:
                print(
                    f"[PORTAL] Campo '{nome_campo}' localizado estruturalmente "
                    "a partir do rótulo."
                )
                return campo.first
        except Exception as e:
            erros.append(f"rótulo/following-input: {e}")

        # 2) Fallback por label HTML, caso exista associação for/id.
        try:
            label = dados.locator("label").filter(
                has_text=padrao_rotulo
            ).first
            if label.count() > 0:
                for_attr = label.get_attribute("for")
                if for_attr:
                    campo = dados.locator(f"#{for_attr}")
                    if campo.count() > 0:
                        print(
                            f"[PORTAL] Campo '{nome_campo}' localizado por "
                            f"label[for='{for_attr}']."
                        )
                        return campo.first
        except Exception as e:
            erros.append(f"label-for: {e}")

        # 3) Fallback por IDs conhecidos de versões/processos já observados.
        for seletor in ids_fallback:
            try:
                campo = dados.locator(seletor)
                if campo.count() > 0:
                    print(
                        f"[PORTAL] Campo '{nome_campo}' localizado pelo fallback "
                        f"de ID {seletor}."
                    )
                    return campo.first
            except Exception as e:
                erros.append(f"{seletor}: {e}")

        raise RuntimeError(
            f"Não consegui localizar o campo '{nome_campo}' por rótulo nem "
            f"pelos IDs de contingência. Detalhes: {' | '.join(erros[-4:])}"
        )

    campo_total = _campo_por_rotulo(
        re.compile(r"Total\s+de\s+Documentos", re.I),
        "Total de Documentos",
        ids_fallback=(
            "#Dad_6908_1",
            "#Dad_4846_1",
        ),
    )

    campo_processados = _campo_por_rotulo(
        re.compile(r"Documentos\s+processados", re.I),
        "Documentos processados",
        ids_fallback=(
            "#Dad_7329_1",
            "#Dad_7234_1",
        ),
    )

    # Preenche usando Playwright normal.
    campo_total.click(timeout=15000)
    campo_total.fill(total, timeout=15000)

    campo_processados.click(timeout=15000)
    campo_processados.fill(total, timeout=15000)

    valor_total = str(
        campo_total.input_value(timeout=5000) or ""
    ).strip()
    valor_processados = str(
        campo_processados.input_value(timeout=5000) or ""
    ).strip()

    if valor_total != total:
        raise RuntimeError(
            "Total de Documentos não confirmou o valor preenchido. "
            f"Esperado={total}; observado={valor_total!r}."
        )

    if valor_processados != total:
        raise RuntimeError(
            "Documentos processados não confirmou o valor preenchido. "
            f"Esperado={total}; observado={valor_processados!r}."
        )

    print(
        "[PORTAL] Quantidades confirmadas: "
        f"Total={valor_total}; Processados={valor_processados}."
    )

    # Prioridade = Não.
    try:
        radio_nao = dados.get_by_role(
            "radio",
            name="Não",
            exact=True,
        )
        radio_nao.check(timeout=10000)
    except Exception as e_role:
        print(
            "[PORTAL] Rádio 'Não' não respondeu pelo role; tentando estrutura "
            f"dos radios. Detalhe: {e_role}"
        )
        radios = dados.locator("input[type='radio']")
        qtd_radios = radios.count()
        if qtd_radios < 2:
            raise RuntimeError(
                f"Esperava ao menos 2 radios em 'Com prioridade?', mas "
                f"encontrei {qtd_radios}."
            )
        # Ordem visível confirmada no PORTAL: Sim | Não | Não Informado
        radios.nth(1).check(timeout=8000)

    print("[PORTAL] 'Com prioridade?' marcado como Não.")

    # Salva dentro do próprio frameDadosExtrasDemo.
    print("[PORTAL] Salvando Dados Adicionais...")
    salvar = dados.get_by_role(
        "button",
        name="Salvar",
        exact=True,
    )
    salvar.click(timeout=20000)

    page.wait_for_timeout(TEMPO_PORTAL_ACAO)

    # Confirma a mensagem de sucesso/Ok.
    confirmou = fechar_confirmacao_gravar_portal(
        page,
        frame_principal=frame,
        timeout=25000,
        aguardar_aparecer=True,
    )

    if not confirmou:
        frame_atual = frame_principal_portal(page)
        if not clicar_ok_portal(
            frame_atual,
            timeout=15000,
            page=None,
        ):
            raise RuntimeError(
                "Dados Adicionais foram preenchidos e Salvar foi acionado, "
                "mas não consegui confirmar o Ok do PORTAL."
            )

    print(
        f"[PORTAL] Dados Adicionais salvos com sucesso para {requisicao}: "
        f"Total={total}; Processados={total}; Prioridade=Não."
    )

    # Frame atual para anexação/conclusão.
    return frame_principal_portal(page)


def ler_estado_pdf_na_grade_portal(andoc, pdf):
    """Distingue título editável do upload e link de documento já gravado.

    No print do TESTE 30, o nome do novo PDF é o VALUE de um input text.
    get_by_text() não lê esse valor. Também não aceitamos o nome presente
    apenas no input file: selecionar um arquivo não comprova seu upload.
    """
    nome = os.path.basename(pdf)
    return andoc.locator("body").evaluate(
        r"""(body, nome) => {
            const norm = s => String(s || '').trim().toLowerCase();
            const visivel = el => {
                const css = getComputedStyle(el);
                return !!el.getClientRects().length &&
                    css.visibility !== 'hidden' && css.display !== 'none';
            };
            const esperado = norm(nome);
            if (!esperado.endsWith('.pdf')) throw new Error('Arquivo esperado não é PDF.');
            const campos = [...body.querySelectorAll(
                'input[type="text"], input:not([type]), textarea'
            )].filter(visivel);
            const titulos = campos.map(el => el.value || '');
            const links = [...body.querySelectorAll('a')].filter(visivel);
            const gravado = links.some(el => {
                const corresponde = [el.innerText, el.title].some(
                    valor => norm(valor) === esperado
                );
                if (!corresponde) return false;
                // Uma linha ainda editável não vale como documento salvo.
                const linha = el.closest('tr');
                return !linha || ![...linha.querySelectorAll(
                    'input[type="text"], input:not([type]), textarea'
                )].some(visivel);
            });
            const carregado = [...body.querySelectorAll('*')].some(el =>
                visivel(el) && /^Carregado\s*!?$/i.test((el.textContent || '').trim())
            );
            return {
                titulo_editavel: titulos.some(valor => norm(valor) === esperado),
                gravado,
                carregado,
                titulos: titulos.slice(0, 20),
                links: links.map(el => (el.innerText || el.title || '').trim()).slice(0, 30)
            };
        }""",
        nome,
        timeout=2000,
    )


def localizar_input_upload_portal(page, andoc):
    """Busca somente em Documentos da requisição atual e nos seus subframes."""
    contextos = [andoc]
    for fr in page.frames:
        ancestral = fr.parent_frame
        while ancestral is not None:
            if ancestral.name == "frameDocumentosDemo":
                pai = ancestral.parent_frame
                if pai is not None and pai.name == "PortalPainelDemo":
                    contextos.append(fr)
                break
            ancestral = ancestral.parent_frame
    for contexto in contextos:
        try:
            inputs = contexto.locator("input[type='file']")
            for i in range(inputs.count()):
                campo = inputs.nth(i)
                if campo.is_enabled(timeout=500):
                    return campo
        except PlaywrightError:
            continue
    return None


def localizar_abertura_documentos_portal(andoc, pdf):
    """Reconhece a grade observada sem seletor: #icoEnviar / Enviar Documento.

    Só usa esse botão na requisição indicada pela pasta do PDF, sem arquivo
    ou título em edição e sem transferência pendente. Ele abre a seleção;
    o envio/gravação só prossegue após o arquivo e o upload serem confirmados.
    """
    pasta = os.path.dirname(os.path.abspath(pdf))
    if os.path.basename(pasta) == 'anexos_portal':
        pasta = os.path.dirname(pasta)
    requisicao = os.path.basename(pasta)
    if not re.fullmatch(r'\d{6,10}', requisicao):
        return None
    nome = os.path.basename(pdf)
    identidade = re.match(r'^NC_PORTAL(\d+)_OV\d+_DOC\d+\.pdf$', nome, re.I)
    if identidade and identidade.group(1).lstrip('0') != requisicao.lstrip('0'):
        raise RuntimeError('A identidade do PDF difere da pasta da requisição; abertura do upload bloqueada.')
    candidatos = andoc.get_by_role('button', name='Enviar Documento', exact=True)
    visiveis = [candidatos.nth(i) for i in range(candidatos.count())
                if candidatos.nth(i).is_visible(timeout=500)]
    if not visiveis:
        return None
    if len(visiveis) != 1:
        raise RuntimeError('Mais de um botão Enviar Documento visível; abertura de upload ambígua.')
    botao = visiveis[0]
    if botao.get_attribute('id') != 'icoEnviar' or not botao.is_enabled(timeout=500):
        return None
    estado = andoc.locator('body').evaluate(r'''body => {
        const vis = e => !!e.getClientRects().length && getComputedStyle(e).visibility !== 'hidden';
        return {
            texto: body.innerText || '',
            inputsArquivo: body.querySelectorAll('input[type=file]').length,
            editaveis: [...body.querySelectorAll('input[type=text],input:not([type]),textarea')].filter(vis).length,
            pendente: [...body.querySelectorAll('[aria-busy=true],[role=progressbar]')].some(vis)
                || /carregado\s*!|carregando|enviando|aguarde/i.test(body.innerText || '')
        };
    }''')
    if not re.match(r'\s*0*' + re.escape(requisicao) + r'\s*-', estado.get('texto', '')):
        raise RuntimeError(f'Documentos não confirmou o cabeçalho da requisição {requisicao}; upload bloqueado.')
    if estado.get('inputsArquivo') or estado.get('editaveis') or estado.get('pendente'):
        raise RuntimeError('Documentos já contém um controle/arquivo em edição ou upload pendente; '
                           'Enviar Documento não será acionado novamente.')
    return botao


def selecionar_arquivo_upload_portal(page, andoc, pdf, registrar=True):
    """Seleciona uma vez pelo input ou seletor de arquivo, antes de qualquer envio.

    False significa que nenhum controle foi acionado. Depois de um clique ou
    set_input_files incerto, lança erro para não repetir o envio do arquivo.
    """
    campo = localizar_input_upload_portal(page, andoc)
    # A grade pode aparecer antes do componente de seleção do arquivo.
    limite_input = time.monotonic() + (12.0 if registrar else 0.0)
    while campo is None and time.monotonic() < limite_input:
        page.wait_for_timeout(300)
        campo = localizar_input_upload_portal(page, andoc)
    if campo is not None:
        campo.set_input_files(pdf, timeout=15000)
        return True, "Arquivo definido no input[type=file] de Documentos."

    padrao = re.compile(
        r"^(?:Choose\s+Files?|Browse(?:\.\.\.)?|"
        r"(?:Selecionar|Escolher|Procurar|Anexar|Adicionar)\s*(?:arquivo|documento)?s?)$",
        re.I,
    )
    botao = None
    for candidatos in (
        andoc.get_by_role("button", name=padrao),
        andoc.locator("label, a, [role='button']").filter(has_text=padrao),
    ):
        for i in range(min(candidatos.count(), 20)):
            candidato = candidatos.nth(i)
            if candidato.is_visible(timeout=500) and candidato.is_enabled(timeout=500):
                botao = candidato
                break
        if botao is not None:
            break
    if botao is None:
        botao = localizar_abertura_documentos_portal(andoc, pdf)
        if botao is not None and registrar:
            print('[PORTAL] Grade de Documentos confirmada sem arquivo em edição; '
                  'abrindo o seletor pelo botão Enviar Documento uma única vez.')
    if botao is None:
        return False, 'Seletor de arquivo e abertura de upload não confirmados na grade de Documentos.'

    if registrar:
        print("[PORTAL] Acionando o seletor de arquivo uma única vez.")
    clicou = False
    try:
        with page.expect_file_chooser(timeout=2500) as fc_info:
            botao.click(timeout=5000)
            clicou = True
    except PlaywrightError as erro:
        if not clicou:
            raise RuntimeError(
                "Clique de upload sem confirmação. Não será repetido automaticamente."
            ) from erro
    else:
        fc_info.value.set_files(pdf, timeout=15000)
        return True, "Arquivo definido pelo seletor aberto no controle de upload."

    # Alguns seletores abrem um formulário antes de expor o input de arquivo.
    # Não reabre Documentos aqui, pois isso fecharia o formulário recém-aberto.
    limite = time.monotonic() + 15.0
    while time.monotonic() < limite:
        campo = localizar_input_upload_portal(page, andoc)
        if campo is not None:
            campo.set_input_files(pdf, timeout=15000)
            return True, "Arquivo definido no formulário do seletor de arquivo."
        page.wait_for_timeout(300)
    raise RuntimeError(
        "Controle de upload acionado uma vez, mas nenhum seletor/formulário foi confirmado. "
        "Anexo interrompido sem repetir o clique; requisição permanece aberta."
    )


def anexar_pdfs_no_portal(page, frame, pdfs):
    """
    Anexa PDFs no PORTAL.

    TESTE 32 - conferência do upload e do documento gravado:
    1. Abre a aba "Documentos" no PortalPainelDemo.
    2. Aguarda o iframe frameDocumentosDemo.
    3. Aguarda input[type='file'] ou abre um seletor de arquivo uma única vez.
       Aceita seletor de arquivo ou formulário dentro de Documentos.
    4. Lê o VALUE do título e aguarda "Carregado!" antes de Salvar.
    5. Aguarda o Ok tardio, sem repetir Salvar, e confere o link na grade.
    6. PDFs já presentes são conferidos antes de tentar novo upload.
    """
    if not pdfs:
        return False, "Nenhum PDF disponível para anexar."

    for pdf in pdfs:
        if not os.path.exists(pdf):
            return False, f"PDF não localizado para anexo: {pdf}"

    def abrir_tela_anexo():
        """Abre a aba Documentos e o iframe frameDocumentosDemo do PORTAL."""
        fechar_modal_residual_portal(
            page,
            motivo="antes de abrir Documentos para anexar PDF",
            timeout=5000,
        )

        portalmain = frame_portalmain(frame_principal_portal(page))

        # Fluxo validado pelo Codegen da fila Operações de Faturamento:
        # PortalPainelDemo -> Documentos -> frameDocumentosDemo -> Choose File / input[type=file].
        abriu = False
        erros = []
        for acao in (
            lambda: portalmain.get_by_title("Documentos").click(timeout=12000),
            lambda: portalmain.get_by_text("Documentos", exact=True).click(timeout=12000),
            lambda: portalmain.get_by_text("Documentos").first.click(timeout=12000),
        ):
            try:
                acao()
                abriu = True
                break
            except Exception as e:
                erros.append(str(e))

        if not abriu:
            raise Exception(
                "Não consegui abrir a aba Documentos no PORTAL. "
                + (erros[-1] if erros else "")
            )

        page.wait_for_timeout(1200)
        portalmain.locator(FRAME_DOCUMENTOS).wait_for(
            state="attached", timeout=20000
        )
        andoc = portalmain.frame_locator(FRAME_DOCUMENTOS)
        try:
            andoc.locator("body").wait_for(timeout=10000)
        except Exception:
            pass
        return andoc

    def diagnosticar_controles_upload(andoc):
        """Registra somente controles visíveis quando o input file não materializa."""
        try:
            dados = andoc.locator("body").evaluate("""body => {
                const visivel = el => {
                    const r = el.getBoundingClientRect();
                    const css = getComputedStyle(el);
                    return r.width > 0 && r.height > 0 && css.display !== 'none' &&
                           css.visibility !== 'hidden' && !el.closest('[aria-hidden="true"]');
                };
                const seletor = 'button, input, a, label, [role="button"], img';
                const controles = [...body.querySelectorAll(seletor)]
                    .filter(visivel)
                    .slice(0, 40)
                    .map(el => ({
                        tag: el.tagName,
                        type: el.getAttribute('type') || '',
                        text: (el.innerText || el.value || el.title || el.alt || '').trim().slice(0, 120),
                        id: el.id || '',
                        name: el.getAttribute('name') || ''
                    }));
                const dropzones = [...body.querySelectorAll('*')]
                    .filter(el => visivel(el) && /arraste|solte|drag|drop/i.test(el.innerText || ''))
                    .slice(0, 10)
                    .map(el => ({tag:el.tagName, text:(el.innerText || '').trim().slice(0,160), id:el.id || '', cls:el.className || ''}));
                return {
                    texto: (body.innerText || '').trim().slice(0, 1000),
                    controles,
                    dropzones
                };
            }""")
            print(f"[PORTAL DIAG UPLOAD] {dados}")
            return dados
        except Exception as erro:
            print(f"[PORTAL DIAG UPLOAD] Não foi possível ler controles visíveis: {erro}")
            return {}

    def aguardar_upload_pronto(andoc, pdf, timeout=60000):
        """
        Aguarda o componente de upload confirmar que o arquivo terminou de
        carregar antes de clicar em Salvar.

        TESTE 32: título em input.value + faixa "Carregado!", na mesma
        leitura. Ver somente o nome por oito segundos não comprova upload.
        Retorna False se surgir duplicidade; o chamador conferirá a grade.
        """
        nome_pdf = os.path.basename(pdf)
        inicio = time.monotonic()
        limite = inicio + max(0, timeout) / 1000.0
        ultimo_log = -999
        estado = {}
        ultimo_erro = ""

        while time.monotonic() < limite:
            decorrido = time.monotonic() - inicio
            if PORTAL_DIALOG_STATE.get("arquivo_ja_anexado"):
                return False

            try:
                estado = ler_estado_pdf_na_grade_portal(andoc, pdf)
            except Exception as e:
                estado = {}
                ultimo_erro = str(e)

            if estado.get("titulo_editavel") and estado.get("carregado"):
                print(
                    f"[PORTAL] Upload do PDF confirmado pelo componente: {nome_pdf} "
                    f"(Carregado!)."
                )
                return True

            segundo = int(decorrido)
            if segundo - ultimo_log >= 5:
                ultimo_log = segundo
                print(
                    f"[PORTAL] Aguardando upload do PDF terminar... "
                    f"{decorrido:.1f}s/{timeout/1000:.0f}s "
                    f"(título em input={'SIM' if estado.get('titulo_editavel') else 'NÃO'}; "
                    f"Carregado={'SIM' if estado.get('carregado') else 'NÃO'})"
                )

            page.wait_for_timeout(500)

        raise Exception(
            f"O arquivo {nome_pdf} foi definido no input, mas o componente "
            "não confirmou título + Carregado! dentro do tempo. "
            f"Último estado: {estado}. Detalhe: {ultimo_erro}"
        )

    def fechar_confirmacao_anexo_portal(timeout=75000, historico_inicio=0):
        """
        Fecha especificamente o modal de sucesso do anexo.

        TESTE 31:
        O print automático do TESTE 30 mostra um spinner, ainda sem modal.
        O print seguinte mostra a confirmação aberta. Isso é compatível
        com confirmação tardia; os prints não medem a duração exata.

        Aqui:
        - aguardamos até 75 s;
        - reacquirimos framePrincipalDemo em TODA iteração;
        - tentamos primeiro o caminho exato mostrado na tela/Codegen;
        - usamos o helper genérico apenas como fallback;
        - nunca clicamos Salvar uma segunda vez, evitando duplicidade.
        """
        inicio = time.monotonic()
        limite = inicio + max(0, timeout) / 1000.0
        ultimo_log = -999
        ultimo_erro = ""

        while time.monotonic() < limite:
            decorrido = time.monotonic() - inicio

            # O aviso pode chegar só DEPOIS de Salvar, não apenas ao selecionar
            # o arquivo. O handler rejeita nova versão; conferiremos a grade.
            if PORTAL_DIALOG_STATE.get("arquivo_ja_anexado"):
                return False

            # Aceitar um alert não significa que ele informou sucesso.
            # Só vale uma mensagem positiva emitida por ESTE salvamento.
            novas_mensagens = PORTAL_DIALOG_STATE.get("historico", [])[historico_inicio:]
            if any(PADRAO_SUCESSO_GRAVAR_PORTAL.search(msg or "") for msg in novas_mensagens):
                print("[PORTAL] Sucesso do anexo confirmado por alert nativo desta operação.")
                return True

            # Caminho principal: modal observado dentro do framePrincipalDemo.
            try:
                principal_atual = frame_principal_portal(page)

                mensagem = principal_atual.get_by_text(
                    PADRAO_SUCESSO_GRAVAR_PORTAL
                ).first

                if mensagem.is_visible(timeout=700):
                    print(
                        "[PORTAL] Modal de sucesso do anexo localizado no "
                        "framePrincipalDemo."
                    )

                    candidatos_ok = [
                        principal_atual.get_by_text("Ok", exact=True),
                        principal_atual.get_by_role(
                            "button",
                            name=re.compile(r"^\s*Ok\s*$", re.I),
                        ),
                        principal_atual.locator(
                            "input[type='button'][value='Ok' i], "
                            "input[type='submit'][value='Ok' i]"
                        ),
                    ]

                    for candidato in candidatos_ok:
                        try:
                            alvo = candidato.first
                            if not alvo.is_visible(timeout=600):
                                continue

                            try:
                                alvo.click(timeout=7000)
                                modo = "clique normal"
                            except Exception:
                                try:
                                    alvo.click(timeout=5000, force=True)
                                    modo = "clique forçado"
                                except Exception:
                                    alvo.evaluate("el => el.click()")
                                    modo = "clique DOM"

                            # Confirma que o modal realmente saiu.
                            try:
                                mensagem.wait_for(
                                    state="hidden",
                                    timeout=10000,
                                )
                            except Exception:
                                try:
                                    if mensagem.is_visible(timeout=500):
                                        ultimo_erro = (
                                            "Ok foi acionado, mas o modal de "
                                            "sucesso permaneceu visível."
                                        )
                                        continue
                                except Exception:
                                    pass

                            print(
                                f"[PORTAL] Ok do anexo confirmado ({modo}); "
                                "modal 'Gravar' fechado."
                            )
                            return True

                        except Exception as e_ok:
                            ultimo_erro = str(e_ok)

            except Exception as e_principal:
                ultimo_erro = str(e_principal)

            # Fallback: varredura de todos os frames, mas sem esperar mais
            # 25 segundos dentro do helper a cada iteração.
            try:
                if fechar_confirmacao_gravar_portal(
                    page,
                    frame_principal=frame_principal_portal(page),
                    timeout=0,
                    aguardar_aparecer=False,
                ):
                    print(
                        "[PORTAL] Modal de sucesso do anexo fechado pelo "
                        "fallback genérico."
                    )
                    return True
            except Exception as e_fallback:
                ultimo_erro = str(e_fallback)

            segundo = int(decorrido)
            if segundo - ultimo_log >= 5:
                ultimo_log = segundo
                print(
                    "[PORTAL] Aguardando confirmação final do anexo "
                    f"'Operação efetuada com sucesso'... "
                    f"{decorrido:.1f}s/{timeout/1000:.0f}s"
                )

            page.wait_for_timeout(500)

        raise Exception(
            "Salvar do anexo foi acionado, mas o modal de sucesso/Ok não foi "
            f"fechado em {int(timeout/1000)} s. Último detalhe: "
            f"{ultimo_erro or 'sem detalhe adicional'}"
        )

    def salvar_anexo(andoc, pdf):
        """
        Salva um PDF sem repetir o clique em Salvar.

        A confirmação do PORTAL pode demorar dezenas de segundos enquanto o
        componente exibe o spinner. Esperamos o próprio modal final.
        """
        if not aguardar_upload_pronto(andoc, pdf, timeout=TEMPO_PORTAL_UPLOAD_PRONTO):
            return False
        if PORTAL_DIALOG_STATE.get("arquivo_ja_anexado"):
            return False

        print(f"[PORTAL] Salvando anexo no PORTAL: {os.path.basename(pdf)}")
        historico_inicio = len(PORTAL_DIALOG_STATE.get("historico", []))
        andoc.get_by_role("button", name="Salvar", exact=True).click(timeout=20000)

        # Não fazemos sleep fixo longo: o polling abaixo aguarda o evento real.
        page.wait_for_timeout(500)

        return fechar_confirmacao_anexo_portal(
            timeout=TEMPO_PORTAL_CONFIRMACAO_ANEXO,
            historico_inicio=historico_inicio,
        )

    def conferir_pdf_gravado(pdf):
        """Reabre Documentos após o Ok e exige o link exato do PDF salvo.

        Título editável e faixa Carregado! não comprovam gravação. Na tela
        observada, o documento persistido tem um link na lista de anexos.
        """
        nome = os.path.basename(pdf)
        andoc_atual = abrir_tela_anexo()
        inicio = time.monotonic()
        limite = inicio + TEMPO_PORTAL_CONFERENCIA_ANEXO / 1000.0
        ultimo_log = -999
        estado = {}
        ultimo_erro = ""
        while time.monotonic() < limite:
            try:
                estado = ler_estado_pdf_na_grade_portal(andoc_atual, pdf)
                if estado.get("gravado"):
                    print(f"[PORTAL] PDF gravado conferido na lista de Documentos: {nome}")
                    return True
            except Exception as e:
                ultimo_erro = str(e)
            decorrido = time.monotonic() - inicio
            if int(decorrido) - ultimo_log >= 5:
                ultimo_log = int(decorrido)
                print(f"[PORTAL] Conferindo link gravado de {nome}... {decorrido:.1f}s")
            page.wait_for_timeout(500)
        raise RuntimeError(
            f"Não foi possível confirmar o link gravado de {nome} em Documentos. "
            "O robô não repetirá Salvar nem concluirá a requisição. "
            f"Último estado: {estado}. Detalhe: {ultimo_erro}"
        )

    def fechar_tela_anexo_sem_salvar(andoc):
        """Fecha/cancela a tela de anexo quando o arquivo já existe no PORTAL, evitando salvar duplicado."""
        tentativas = [
            lambda: andoc.get_by_role("button", name=re.compile(r"Cancelar|Fechar|Voltar", re.I)).click(timeout=4000),
            lambda: andoc.locator("input[type='button'][value*='Cancelar' i]").first.click(timeout=4000),
            lambda: andoc.locator("input[type='button'][value*='Fechar' i]").first.click(timeout=4000),
            lambda: page.keyboard.press("Escape"),
        ]

        for tentativa in tentativas:
            try:
                tentativa()
                page.wait_for_timeout(1000)
                return True
            except Exception:
                continue
        return False

    anexados = 0
    ja_existiam = 0
    erros = []

    for pdf in pdfs:
        try:
            print(f"[PORTAL] Anexando PDF: {pdf}")

            # Nunca inicia outro anexo com a confirmação do anterior bloqueando a tela.
            if fechar_confirmacao_gravar_portal(
                page,
                frame_principal=frame,
                timeout=0,
                aguardar_aparecer=False,
            ):
                print("[PORTAL] Confirmação pendente do anexo anterior removida antes de continuar.")

            # Limpa flag antes de cada arquivo para detectar duplicidade deste PDF específico.
            PORTAL_DIALOG_STATE["arquivo_ja_anexado"] = False
            PORTAL_DIALOG_STATE["ultimo_dialog"] = ""

            andoc = abrir_tela_anexo()

            # Uma rodada anterior pode ter gravado o PDF e perdido o Ok.
            # Confere o documento persistido ANTES de selecionar o arquivo.
            if ler_estado_pdf_na_grade_portal(andoc, pdf).get("gravado"):
                ja_existiam += 1
                anexados += 1
                print(f"[PORTAL] PDF já gravado na lista; upload dispensado: {os.path.basename(pdf)}")
                continue

            ok_set, msg_set = selecionar_arquivo_upload_portal(page, andoc, pdf)
            if not ok_set:
                # TESTE 44: input ausente pode significar frame/tela de Documentos
                # ainda incompleto ou uma confirmação conhecida sobreposta. Antes
                # do filechooser, reacquire a tela uma única vez e espera o input real.
                print(
                    "[PORTAL] Controle de upload não materializado. Reabrindo Documentos "
                    "e readquirindo frameDocumentosDemo uma única vez antes do fallback."
                )
                try:
                    fechar_confirmacao_gravar_portal(
                        page,
                        frame_principal=frame_principal_portal(page),
                        timeout=2500,
                        aguardar_aparecer=False,
                    )
                except Exception:
                    pass

                andoc = abrir_tela_anexo()

                # Uma resposta tardia pode ter gravado o arquivo apesar da tela
                # anterior vazia. Confere a lista antes de qualquer novo upload.
                try:
                    if ler_estado_pdf_na_grade_portal(andoc, pdf).get("gravado"):
                        ja_existiam += 1
                        anexados += 1
                        print(
                            f"[PORTAL] PDF já apareceu na lista após recuperar Documentos; "
                            f"novo upload dispensado: {os.path.basename(pdf)}"
                        )
                        continue
                except Exception:
                    pass

                limite_input = time.monotonic() + 10.0
                ultimo_msg_input = msg_set
                while time.monotonic() < limite_input:
                    ok_set, ultimo_msg_input = selecionar_arquivo_upload_portal(
                        page,
                        andoc,
                        pdf,
                        registrar=False,
                    )
                    if ok_set:
                        msg_set = ultimo_msg_input
                        print(
                            "[PORTAL] Controle de upload disponível após reabrir Documentos; "
                            "arquivo selecionado com confirmação."
                        )
                        break
                    page.wait_for_timeout(500)

                if not ok_set:
                    msg_set = ultimo_msg_input
                    diagnosticar_controles_upload(andoc)
                    print(f"[PORTAL] Última verificação do controle de upload. Motivo: {msg_set}")
                    ok_set, msg_set = selecionar_arquivo_upload_portal(page, andoc, pdf)

            # Se o PORTAL avisou que o arquivo já estava anexado, NÃO salva novamente.
            # Só conta como OK depois de conferir o documento na grade.
            if PORTAL_DIALOG_STATE.get("arquivo_ja_anexado"):
                fechar_tela_anexo_sem_salvar(andoc)
                conferir_pdf_gravado(pdf)
                ja_existiam += 1
                anexados += 1
                print(f"[PORTAL] PDF já existia no PORTAL; não foi anexado como nova versão: {os.path.basename(pdf)}")
                continue

            if not ok_set:
                raise Exception(msg_set)

            print(f"[PORTAL] {msg_set}")

            # A espera pelo estado do upload substitui os 15 s fixos. A mesma
            # checagem de duplicidade continua durante o upload e após Salvar.
            confirmou_novo = salvar_anexo(andoc, pdf)
            if not confirmou_novo:
                fechar_tela_anexo_sem_salvar(andoc)
                conferir_pdf_gravado(pdf)
                ja_existiam += 1
                anexados += 1
                print(f"[PORTAL] Duplicidade rejeitada e PDF existente conferido: {os.path.basename(pdf)}")
                continue

            conferir_pdf_gravado(pdf)
            anexados += 1
            print(f"[PORTAL] PDF anexado com sucesso: {os.path.basename(pdf)}")

        except Exception as e_um:
            erros.append(f"{os.path.basename(pdf)}: {e_um}")
            print(f"[PORTAL] Erro ao anexar {pdf}: {e_um}")

            # Diagnóstico funcional sem depender de um print isolado.
            try:
                andoc_atual = frame_portalmain(frame_principal_portal(page)).frame_locator(
                    FRAME_DOCUMENTOS
                )
                print(f"[PORTAL DIAG ANEXO] {ler_estado_pdf_na_grade_portal(andoc_atual, pdf)}")
            except Exception as e_diag:
                print(f"[PORTAL DIAG ANEXO] Grade indisponível: {e_diag}")
            print(f"[PORTAL DIAG ANEXO] Último dialog: {PORTAL_DIALOG_STATE.get('ultimo_dialog', '')!r}")

            # Evidência para ajuste fino, caso o componente de anexo varie.
            try:
                caminho_print = os.path.join(
                    os.path.dirname(pdf),
                    f"erro_anexo_{os.path.basename(pdf)}.png",
                )
                page.screenshot(path=caminho_print, full_page=True)
                print(f"[PORTAL] Print do erro de anexo salvo em: {caminho_print}")
            except Exception:
                pass

            # Remove somente uma confirmação conhecida de sucesso que tenha
            # acabado de chegar. Isso não converte a falha em sucesso.
            try:
                fechar_confirmacao_gravar_portal(
                    page,
                    frame_principal=frame,
                    timeout=3000,
                    aguardar_aparecer=True,
                )
            except Exception:
                pass

            # Um salvamento sem confirmação pode continuar no servidor.
            # Não mistura a resposta tardia com o próximo PDF desta requisição.
            print("[PORTAL] Lote de anexos interrompido; requisição permanecerá aberta para conferência.")
            break

    if anexados == len(pdfs):
        if ja_existiam:
            return True, f"{anexados} PDF(s) OK no PORTAL; {ja_existiam} já existia(m) e foram pulados para evitar duplicidade."
        return True, f"{anexados} PDF(s) anexado(s) individualmente."

    return False, f"Anexados/OK {anexados}/{len(pdfs)}. Já existiam: {ja_existiam}. Erros: {' | '.join(erros)}"

def concluir_requisicao_portal(
    page,
    frame,
    texto_encerramento,
    pasta_evidencias=None,
):
    """
    Conclui a requisição no PORTAL com texto de encerramento.

    TESTE 10:
    - não exige cegamente um "Ok" antes do editor: primeiro verifica se o
      iframe fraAcaoCONC já foi aberto;
    - aceita tanto o fluxo com confirmação intermediária quanto o fluxo em que
      o handler global já tratou a confirmação;
    - localiza o Rich Text por mais de um seletor e valida que o texto entrou;
    - salva a conclusão e confirma o Ok final quando ele existir;
    - grava logs de cada etapa para deixar claro exatamente onde o PORTAL parou.
    """

    def _frame_conclusao_disponivel(portalmain, timeout_ms=3000):
        try:
            portalmain.locator("iframe[name='fraAcaoCONC']").wait_for(
                state="attached",
                timeout=timeout_ms,
            )
            return True
        except Exception:
            return False

    def _obter_editor_conclusao(conc):
        seletores = [
            "iframe[title='Editor de Rich Text, areaDescricao']",
            "iframe[title*='areaDescricao']",
            "iframe[title*='Editor de Rich Text']",
        ]

        erros = []
        for seletor in seletores:
            try:
                iframe_editor = conc.locator(seletor).first
                iframe_editor.wait_for(state="attached", timeout=7000)
                editor = conc.frame_locator(seletor).first
                corpo = editor.locator("body")
                corpo.wait_for(state="visible", timeout=7000)
                return corpo, seletor
            except Exception as e:
                erros.append(f"{seletor}: {e}")

        raise Exception(
            "Não localizei o editor Rich Text da conclusão. "
            + " | ".join(erros[-2:])
        )

    def _preencher_editor_conclusao(editor_body, texto):
        # 1) fill direto no body contenteditable (fluxo observado no codegen).
        try:
            editor_body.click(timeout=12000)
            page.wait_for_timeout(300)
            editor_body.fill(texto)
            page.wait_for_timeout(600)
            atual = str(editor_body.inner_text(timeout=3000) or "").strip()
            if atual:
                return True, "Mensagem preenchida com body.fill()."
        except Exception as e_fill:
            erro_fill = str(e_fill)
        else:
            erro_fill = "body.fill executou, mas o editor permaneceu vazio"

        # 2) fallback JS para CKEditor/contenteditable.
        try:
            editor_body.click(timeout=8000)
            editor_body.evaluate(
                """
                (el, texto) => {
                    el.focus();
                    el.innerHTML = '';
                    el.innerText = texto;
                    el.dispatchEvent(new InputEvent('input', {
                        bubbles: true,
                        inputType: 'insertText',
                        data: texto
                    }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
                }
                """,
                texto,
            )
            page.wait_for_timeout(600)
            atual = str(editor_body.inner_text(timeout=3000) or "").strip()
            if atual:
                return True, "Mensagem preenchida com evaluate()."
            return False, "Evaluate executou, mas o editor continuou vazio."
        except Exception as e_eval:
            return False, (
                "Não consegui preencher a mensagem de conclusão. "
                f"Fill: {erro_fill} | Evaluate: {e_eval}"
            )

    def _clicar_salvar_conclusao(conc):
        tentativas = [
            lambda: conc.get_by_role("button", name="Salvar").click(timeout=20000),
            lambda: conc.get_by_role(
                "button", name=re.compile(r"^\s*Salvar\s*$", re.I)
            ).click(timeout=15000),
            lambda: conc.locator("input[type='button'][value='Salvar']").first.click(timeout=12000),
            lambda: conc.locator("button:has-text('Salvar')").first.click(timeout=12000),
            lambda: conc.get_by_text("Salvar", exact=True).click(timeout=12000),
        ]

        erros = []
        for tentativa in tentativas:
            try:
                tentativa()
                return True, "Clique no Salvar da conclusão realizado."
            except Exception as e:
                erros.append(str(e))

        return False, (
            "Não consegui clicar no Salvar da conclusão. Erros: "
            + " | ".join(erros[-2:])
        )

    try:
        print("[PORTAL] Iniciando conclusão da requisição...")

        portalmain = frame_portalmain(frame)
        historico_antes = len(PORTAL_DIALOG_STATE.get("historico", []))
        eventos_conclusao = len(PORTAL_DIALOG_STATE.get("eventos", []))

        print("[PORTAL] Clicando em 'Concluir Requisição'...")
        frame.get_by_role("img", name="Concluir Requisição").click(timeout=20000)
        page.wait_for_timeout(800)
        verificar_dialogs_portal_desde(eventos_conclusao)

        # Alguns cenários exibem um Ok intermediário; em outros o tab de
        # conclusão já abre diretamente (ou o handler global aceita o dialog).
        if _frame_conclusao_disponivel(portalmain, timeout_ms=2500):
            print("[PORTAL] Tela de conclusão abriu diretamente; Ok intermediário não foi necessário.")
        else:
            print("[PORTAL] Tela de conclusão ainda não abriu; procurando confirmação Ok intermediária...")
            clicou_ok = clicar_ok_portal(frame, timeout=7000)
            if clicou_ok:
                print("[PORTAL] Ok intermediário da conclusão confirmado.")
            elif len(PORTAL_DIALOG_STATE.get("historico", [])) > historico_antes:
                print("[PORTAL] Confirmação intermediária foi tratada pelo handler global de dialogs.")
            else:
                print("[PORTAL] Nenhum Ok intermediário visível; aguardando abertura direta da conclusão.")

        portalmain.locator("iframe[name='fraAcaoCONC']").wait_for(
            state="attached",
            timeout=30000,
        )
        conc = portalmain.frame_locator("iframe[name='fraAcaoCONC']")
        print("[PORTAL] iframe fraAcaoCONC localizado.")

        corpo, seletor_editor = _obter_editor_conclusao(conc)
        print(f"[PORTAL] Editor Rich Text localizado por: {seletor_editor}")

        ok_msg, msg_msg = _preencher_editor_conclusao(corpo, texto_encerramento)
        if not ok_msg:
            raise Exception(msg_msg)

        print(f"[PORTAL] Mensagem de conclusão inserida. {msg_msg}")
        page.wait_for_timeout(700)

        # Confirma explicitamente que há texto antes de salvar.
        texto_no_editor = str(corpo.inner_text(timeout=5000) or "").strip()
        if not texto_no_editor:
            raise Exception("O editor de conclusão continuou vazio após o preenchimento.")
        print(
            "[PORTAL] Texto confirmado no editor de conclusão: "
            + repr(texto_no_editor[:120])
        )

        print("[PORTAL] Clicando no Salvar da conclusão...")
        eventos_antes_salvar = len(PORTAL_DIALOG_STATE.get("eventos", []))
        ok_salvar, msg_salvar = _clicar_salvar_conclusao(conc)
        if not ok_salvar:
            raise Exception(msg_salvar)
        print(f"[PORTAL] {msg_salvar}")
        page.wait_for_timeout(1200)

        verificar_dialogs_portal_desde(eventos_antes_salvar)

        # Depois do Salvar, o PORTAL pode mostrar um Ok HTML ou um dialog que o
        # handler global já aceitou. Aceitamos ambos os comportamentos.
        if clicar_ok_portal(frame, timeout=10000):
            print("[PORTAL] Ok final da conclusão confirmado na interface.")
        elif any(e.get("tipo") == "alert" and e.get("aceito") and
                 re.search(r"sucesso|conclu[ií]da|gravado|salvo", e.get("mensagem", ""), re.I)
                 for e in PORTAL_DIALOG_STATE.get("eventos", [])[eventos_antes_salvar:]):
            print("[PORTAL] Confirmação final da conclusão tratada pelo handler global de dialogs.")
        else:
            # Última evidência: se a própria tela de conclusão saiu, o salvar
            # efetivamente avançou mesmo sem um Ok clicável.
            try:
                ainda_aberta = _frame_conclusao_disponivel(portalmain, timeout_ms=1500)
            except Exception:
                ainda_aberta = False

            if ainda_aberta:
                raise Exception(
                    "O Salvar foi acionado, mas não apareceu confirmação final e "
                    "a tela de conclusão permaneceu aberta."
                )
            print("[PORTAL] Tela de conclusão foi fechada após o Salvar; conclusão considerada confirmada.")

        verificar_dialogs_portal_desde(eventos_conclusao)
        print("[PORTAL] Requisição concluída com sucesso.")
        return True, "Requisição concluída com sucesso."

    except Exception as e:
        print(f"[PORTAL] ERRO na conclusão da requisição: {e}")
        try:
            pasta_destino = pasta_evidencias or PASTA_RESUMO
            caminho_print = os.path.join(
                pasta_destino,
                f"erro_conclusao_portal_{ID_RODADA}.png",
            )
            page.screenshot(path=caminho_print, full_page=True)
            print(f"[PORTAL] Print do erro de conclusão salvo em: {caminho_print}")
        except Exception:
            pass

        return False, f"Erro ao concluir requisição no PORTAL: {e}"

def anexar_pdfs_e_encerrar_portal(page, requisicao, pasta_requisicao, consultas_sap):
    """
    Novo comportamento:
    - anexa todos os PDFs válidos que foram obtidos;
    - se qualquer OV estiver pendente/erro, NÃO conclui a requisição;
    - uma OV com erro não impede anexar PDFs válidos das demais OVs.
    """
    resultado = {
        "StatusPORTAL": "NAO_EXECUTADO",
        "MensagemPORTAL": "",
        "QtdPDFsAnexados": 0,
    }

    if not ENCERRAR_PORTAL_APOS_PDF:
        resultado["StatusPORTAL"] = "DESABILITADO"
        resultado["MensagemPORTAL"] = "ENCERRAR_PORTAL_APOS_PDF=False."
        return resultado

    if not consultas_sap:
        resultado["StatusPORTAL"] = "SEM_OV_LOCALIZADA"
        resultado["MensagemPORTAL"] = "Nenhuma OV localizada/processada; finalização PORTAL não executada."
        return resultado

    pdfs = listar_pdfs_validos_da_requisicao(pasta_requisicao, consultas_sap)
    ok_todos, msg_validacao = todas_ovs_com_pdf_ok(consultas_sap)

    if pdfs:
        print(
            f"[PORTAL] Requisição {requisicao}: {len(pdfs)} PDF(s) válido(s) "
            "localizado(s). A anexação será executada imediatamente antes de "
            "seguir para a próxima requisição."
        )
    else:
        print(
            f"[PORTAL] Requisição {requisicao}: nenhum PDF válido foi gerado/localizado; "
            "não há arquivo para anexar nesta etapa."
        )

    if not pdfs:
        if existe_erro_determinacao_conta_na_requisicao(consultas_sap):
            resultado["StatusPORTAL"] = "PENDENTE_ERRO_SAP"
            resultado["MensagemPORTAL"] = montar_mensagem_erro_determinacao_conta_portal(consultas_sap)
        elif any(
            str(c.get("StatusSAP", "") or "").upper() == STATUS_NOTA_CREDITO_PENDENTE
            for c in (consultas_sap or [])
        ):
            resultado["StatusPORTAL"] = "PENDENTE_SAP"
            resultado["MensagemPORTAL"] = msg_validacao
        else:
            resultado["StatusPORTAL"] = "SEM_PDF"
            resultado["MensagemPORTAL"] = msg_validacao or "Nenhum PDF válido localizado para anexar."
        return resultado

    try:
        # Preserva <OV>.pdf na pasta original e usa identidade completa na cópia de envio.
        pdfs_envio = []
        por_caminho = {os.path.normcase(os.path.abspath(c.get("ArquivoPDF", ""))): c for c in consultas_sap if c.get("ArquivoPDF")}
        for pdf in pdfs:
            consulta = por_caminho[os.path.normcase(os.path.abspath(pdf))]
            if classificar_atividade_suporte_venda(consulta.get("Atividade", "")) == ATIVIDADE_NOTA_CREDITO:
                pdfs_envio.append(preparar_pdf_credito_para_portal(requisicao, consulta, pdf))
            else:
                pdfs_envio.append(pdf)
        pdfs = pdfs_envio
        inicio_dialogs = len(PORTAL_DIALOG_STATE.get("eventos", []))
        frame = preparar_requisicao_portal_para_anexo_encerramento(page, requisicao)

        # Dados Adicionais só representam processamento concluído. Se alguma
        # OV estiver pendente/erro, anexamos os PDFs válidos, mas NÃO gravamos
        # Total/Processados como se a requisição estivesse completa.
        if ok_todos:
            frame = preencher_dados_adicionais_portal(
                page,
                frame,
                len(consultas_sap or []),
                requisicao,
            )
        else:
            print(
                "[PORTAL] Existem pendências na requisição; Dados Adicionais não "
                "serão finalizados nesta execução."
            )

        verificar_dialogs_portal_desde(inicio_dialogs)
        ok_anexo, msg_anexo = anexar_pdfs_no_portal(page, frame, pdfs)
        verificar_dialogs_portal_desde(inicio_dialogs)
        if not ok_anexo:
            resultado["StatusPORTAL"] = "ERRO_ANEXO"
            resultado["MensagemPORTAL"] = msg_anexo + (' | ' + msg_validacao if not ok_todos else '')
            return resultado

        resultado["QtdPDFsAnexados"] = len(pdfs)

        # PDFs corretos foram anexados, mas a requisição permanece aberta se
        # existir qualquer OV sem sucesso/PDF válido.
        if not ok_todos:
            if existe_erro_determinacao_conta_na_requisicao(consultas_sap):
                resultado["StatusPORTAL"] = "PENDENTE_ERRO_SAP"
            else:
                resultado["StatusPORTAL"] = "PENDENTE_PDF"
            resultado["MensagemPORTAL"] = msg_anexo + " | Requisição mantida aberta. " + msg_validacao
            print(
                f"[PORTAL] {len(pdfs)} PDF(s) anexado(s) à requisição {requisicao}, "
                "mas ela NÃO será concluída porque ainda existem pendências."
            )
            return resultado

        texto = montar_texto_encerramento_portal(requisicao, consultas_sap, pdfs)
        ok_conc, msg_conc = concluir_requisicao_portal(
            page,
            frame,
            texto,
            pasta_evidencias=pasta_requisicao,
        )

        if ok_conc:
            resultado["StatusPORTAL"] = "CONCLUIDO"
            resultado['RodadaEncerramentoPORTAL'] = str(globals().get('ID_RODADA') or '')
            resultado["MensagemPORTAL"] = msg_anexo + " | " + msg_conc
        else:
            resultado["StatusPORTAL"] = "ERRO_CONCLUSAO"
            resultado["MensagemPORTAL"] = msg_anexo + " | " + msg_conc

        return resultado

    except StatusPORTALIndeterminado as erro_status:
        resultado["StatusPORTAL"] = "STATUS_PORTAL_INDETERMINADO"
        resultado["MensagemPORTAL"] = str(erro_status)
        return resultado
    except Exception as e:
        mensagem_erro = str(e)
        if "PORTAL_NAO_ACIONAVEL::" in mensagem_erro:
            status_atual = mensagem_erro.split("PORTAL_NAO_ACIONAVEL::", 1)[1].strip()
            resultado["StatusPORTAL"] = "IGNORADA_STATUS_PORTAL"
            resultado["MensagemPORTAL"] = (
                f"Requisição não alterada porque o PORTAL está em estado não acionável: "
                f"{status_atual or 'encerrado/cancelado'}."
            )
            print(f"[PORTAL] {resultado['MensagemPORTAL']}")
            return resultado

        resultado["StatusPORTAL"] = "ERRO_PORTAL"
        resultado["MensagemPORTAL"] = f"Erro no fluxo de anexo/encerramento PORTAL: {e}"
        print(f"[PORTAL] ERRO no fluxo de anexo/encerramento da requisição {requisicao}: {e}")
        try:
            caminho_print = os.path.join(
                pasta_requisicao,
                f"erro_anexo_encerramento_portal_{requisicao}_{ID_RODADA}.png",
            )
            page.screenshot(path=caminho_print, full_page=True)
            print(f"[PORTAL] Evidência do erro salva em: {caminho_print}")
        except Exception:
            pass
        return resultado


# =========================
# MEDIÇÃO DE TEMPO
# =========================

def formatar_duracao(segundos):
    """Formata segundos em HH:MM:SS para facilitar leitura no Excel."""
    try:
        segundos_int = int(round(float(segundos)))
    except Exception:
        return ""

    horas = segundos_int // 3600
    minutos = (segundos_int % 3600) // 60
    seg = segundos_int % 60
    return f"{horas:02d}:{minutos:02d}:{seg:02d}"


def campos_situacao_conciliacao(item):
    """Visão operacional; ausência de confirmação nunca significa 'não existe'."""
    tipo = classificar_atividade_suporte_venda(item.get('Atividade', ''))
    origem = str(item.get('OrigemOVs') or '')
    fase = str(item.get('StatusConciliacao') or '')
    numero = str(item.get('RequisicaoConciliacao') or '').strip()
    req = str(item.get('Requisicao') or '').strip()
    numero_valido = bool(re.fullmatch(r'\d{6,10}', numero) and int(numero)
                         and numero.lstrip('0') != req.lstrip('0'))
    obrigatoria = origem == 'PORTAL_NC_DUAS_PLANILHAS' or bool(fase)
    nao_aplica = (not obrigatoria and (tipo == ATIVIDADE_DEVOLUCAO_REFATURAMENTO
                  or (tipo == ATIVIDADE_NOTA_CREDITO and origem in {
                      'PORTAL_DESCRICAO', 'PORTAL_DESCRICAO_TABELA', 'PORTAL_DESCRICAO_CANDIDATA',
                      'PORTAL_PLANILHA', 'PORTAL_PLANILHA_LOCAL'})))
    confirmado = numero_valido and fase in {'CONCLUIDA', 'CRIADA_ANEXO_PENDENTE'}
    if confirmado:
        criada = 'SIM'
    elif fase in {'ENVIO_INICIADO', 'PENDENTE_CONFERENCIA', 'CONCLUIDA', 'CRIADA_ANEXO_PENDENTE'} or numero:
        criada = 'NÃO CONFIRMADA'
    elif obrigatoria:
        criada = 'NÃO'
    elif nao_aplica:
        criada = 'NÃO SE APLICA'
    else:
        criada = 'NÃO AVALIADA'
    completo = (item.get('StatusPORTAL') == 'CONCLUIDO'
                and (nao_aplica or (confirmado and fase == 'CONCLUIDA')))
    rodada = str(globals().get('ID_RODADA') or '')
    encerrada_agora = bool(rodada and item.get('RodadaEncerramentoPORTAL') == rodada
                          and item.get('StatusPORTAL') == 'CONCLUIDO')
    return {'Conciliação criada?': criada,
            'PORTAL da conciliação': numero if confirmado else '',
            'Processo concluído?': 'SIM' if completo else 'NÃO',
            'PORTAL encerrado nesta rodada?': 'SIM' if encerrada_agora else 'NÃO'}


def imprimir_situacao_conciliacao(item):
    campos = campos_situacao_conciliacao(item)
    print(f"[CONCILIAÇÃO] PORTAL origem={item.get('Requisicao', '')} | "
          f"Criada={campos['Conciliação criada?']} | "
          f"PORTAL da conciliação={campos['PORTAL da conciliação'] or '—'} | "
          f"Processo concluído={campos['Processo concluído?']} | "
          f"Etapa={item.get('StatusConciliacao') or item.get('Etapa', '')}")


def status_sap_ok_para_medicao(status):
    """Único status de sucesso com PDF emitido por este fluxo."""
    return str(status or "").upper().strip() == "OK_VA03"


def requisicao_valida_para_medicao(resultado_portal, consultas_sap):
    """
    Só conta a requisição para medição de produtividade quando:
    - PORTAL foi concluído;
    - existe pelo menos uma OV;
    - todas as OVs tiveram status de sucesso/PDF válido;
    - nenhuma OV ficou com erro SAP, pendência ou erro de determinação de conta.
    """
    if str(resultado_portal.get("StatusPORTAL", "") or "").upper().strip() != "CONCLUIDO":
        return False

    if resultado_portal.get('OrigemOVs') == 'PORTAL_NC_DUAS_PLANILHAS':
        if resultado_portal.get('StatusConciliacao') != 'CONCLUIDA' or resultado_portal.get('RetomadaNC'):
            return False

    if not consultas_sap:
        return False

    for consulta in consultas_sap:
        if not status_sap_ok_para_medicao(consulta.get("StatusSAP", "")):
            return False

    return True


def montar_campos_tempo_requisicao(inicio_req_dt, fim_req_dt, tempo_valido):
    """Monta campos padronizados de tempo para gravar no relatório."""
    try:
        tempo_segundos = round((fim_req_dt - inicio_req_dt).total_seconds(), 2)
    except Exception:
        tempo_segundos = ""

    return {
        "InicioRequisicao": inicio_req_dt.strftime("%d/%m/%Y %H:%M:%S") if inicio_req_dt else "",
        "FimRequisicao": fim_req_dt.strftime("%d/%m/%Y %H:%M:%S") if fim_req_dt else "",
        "TempoReqSegundos": tempo_segundos,
        "TempoReqFormatado": formatar_duracao(tempo_segundos) if tempo_segundos != "" else "",
        "TempoValidoProdutividade": "SIM" if tempo_valido else "NÃO",
    }


# =========================
# EXCEL FINAL
# =========================

def identificar_origem_sap(consulta):
    """
    Identifica a origem principal do processamento da OV para facilitar auditoria.

    Regra operacional AV5:
    - VF01 é usada para faturar/tentar gravar.
    - VA03 é o caminho padrão para localizar e salvar o PDF.
    """
    status = str(consulta.get("StatusSAP", "") or "").upper()
    mensagem = str(consulta.get("Mensagem", "") or "").upper()

    if status == "PDF_JA_EXISTE":
        return "PDF_JA_EXISTE"

    if status == STATUS_ERRO_DETERMINACAO_CONTA:
        return "SAP_FLUXO"

    if status.startswith("ERRO_VF01") or status.startswith("VF01_"):
        return "VF01"

    if "VA03" in status or "VA03" in mensagem:
        return "VA03"

    if status in {"OK", "SEM_FATURAMENTO", "ERRO_SAP", "PENDENTE_PDF", "ERRO"}:
        return "VF01"

    if "VF01" in status or "VF01" in mensagem:
        return "VF01"

    return ""


def montar_linhas_detalhe_ovs(resultados):
    """Uma linha por OV: A=OV, B=Requisição, C=Status SAP, D=Atividade."""
    linhas = []
    for item in resultados:
        consultas = item.get("ConsultasSAP") or [{
            "StatusSAP": "SEM_OV_ENCONTRADA" if item.get("ExtracaoConcluida") else
                         item.get("StatusPORTAL", "NAO_PROCESSADA"),
            "Mensagem": item.get("MensagemPORTAL", ""),
        }]
        for consulta in consultas:
            linhas.append({
                "OV SAP": consulta.get("Ordem", ""),
                "Requisição PORTAL": item.get("Requisicao", ""),
                "Status SAP": consulta.get("StatusSAP", ""),
                "Atividade PORTAL": item.get("Atividade", ""),
                "Origem SAP": identificar_origem_sap(consulta),
                "NF SAP": consulta.get("NumeroNF", ""),
                "Arquivo PDF": consulta.get("ArquivoPDF", ""),
                "Arquivo Anexo PORTAL": consulta.get("ArquivoAnexoPORTAL", ""),
                "Mensagem SAP": consulta.get("Mensagem", ""),
                "Pasta da Requisição": item.get("PastaRequisicao", ""),
                "Início Requisição": item.get("InicioRequisicao", ""),
                "Fim Requisição": item.get("FimRequisicao", ""),
                "Tempo Total Requisição (seg)": item.get("TempoReqSegundos", ""),
                "Tempo Total Requisição": item.get("TempoReqFormatado", ""),
                "Tempo válido para produtividade": item.get("TempoValidoProdutividade", ""),
                "Status PORTAL": item.get("StatusPORTAL", ""),
                **campos_situacao_conciliacao(item),
                "Retomada SAP": "SIM" if consulta.get('RetomadaSAP') else "NÃO",
                "Qtd PDFs Anexados": item.get("QtdPDFsAnexados", ""),
                "Mensagem PORTAL": item.get("MensagemPORTAL", ""),
            })
    return linhas


def ajustar_largura_colunas_excel(writer, sheet_name, df):
    """
    Ajuste simples de largura para deixar o retorno mais legível.
    Usa o engine openpyxl disponível no pandas para o arquivo gerado pelo próprio script.
    """
    try:
        worksheet = writer.sheets[sheet_name]

        for idx, coluna in enumerate(df.columns, start=1):
            valores = [str(coluna)]
            if not df.empty:
                valores.extend(df[coluna].astype(str).fillna("").head(200).tolist())

            largura = max(len(v) for v in valores) + 2

            # Limites para não criar colunas gigantes quando houver descrição/mensagem longa.
            if "Descrição" in str(coluna) or "Mensagem" in str(coluna):
                largura = min(max(largura, 25), 80)
            elif "Arquivo" in str(coluna) or "Pasta" in str(coluna):
                largura = min(max(largura, 25), 70)
            else:
                largura = min(max(largura, 12), 35)

            letra = worksheet.cell(row=1, column=idx).column_letter
            worksheet.column_dimensions[letra].width = largura

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

    except Exception as e:
        print(f"[AVISO EXCEL] Não consegui ajustar largura da aba {sheet_name}: {e}")


def montar_excel_resultado(resultados):
    """Resumo com colunas fixas; detalhes de todas as OVs na segunda aba.

    Atualiza o mesmo arquivo da rodada por substituição atômica. Se estiver
    aberto no Excel, preserva-o e salva uma cópia alternativa com os novos dados.
    """
    linhas = []
    for item in resultados:
        consultas = item.get("ConsultasSAP") or []
        linhas.append({
            "Requisição PORTAL": item.get("Requisicao", ""),
            "Atividade PORTAL": item.get("Atividade", ""),
            "Descrição PORTAL": item.get("Descricao", ""),
            "Pasta da Requisição": item.get("PastaRequisicao", ""),
            "Início Requisição": item.get("InicioRequisicao", ""),
            "Fim Requisição": item.get("FimRequisicao", ""),
            "Tempo Total Requisição (seg)": item.get("TempoReqSegundos", ""),
            "Tempo Total Requisição": item.get("TempoReqFormatado", ""),
            "Tempo válido para produtividade": item.get("TempoValidoProdutividade", ""),
            "Status PORTAL": item.get("StatusPORTAL", ""),
            **campos_situacao_conciliacao(item),
            "OVs retomadas": sum(bool(c.get('RetomadaSAP')) for c in consultas),
            "Qtd PDFs Anexados": item.get("QtdPDFsAnexados", 0),
            "Validação prévia NC": item.get("ValidacaoNC", ""),
            "Status conciliação NC": item.get("StatusConciliacao", ""),
            "Requisição conciliação NC": item.get("RequisicaoConciliacao", ""),
            "Mensagem conciliação NC": item.get("MensagemConciliacao", ""),
            "Planilha conciliação NC": item.get("PlanilhaConciliacao", ""),
            "Retomada NC": "SIM" if item.get("RetomadaNC") else "",
            "Mensagem PORTAL": item.get("MensagemPORTAL", ""),
            "OVs localizadas": ", ".join(str(c.get("Ordem", "")) for c in consultas),
            "Qtd OVs": len(consultas),
            "Qtd OVs com sucesso SAP": sum(status_sap_ok_para_medicao(c.get("StatusSAP")) for c in consultas),
            "Origem OVs": item.get("OrigemOVs", ""),
            "Etapa": item.get("Etapa", ""),
        })
    abas = {"Resumo_Requisicoes": pd.DataFrame(linhas),
            "Detalhe_OVs": pd.DataFrame(montar_linhas_detalhe_ovs(resultados))}
    os.makedirs(PASTA_RESUMO, exist_ok=True)
    destino = montar_caminho_excel_resultado()
    temporario = destino + ".tmp.xlsx"
    with pd.ExcelWriter(temporario, engine="openpyxl") as writer:
        for nome, df in abas.items():
            df.to_excel(writer, sheet_name=nome, index=False)
            ajustar_largura_colunas_excel(writer, nome, df)
    try:
        os.replace(temporario, destino)
    except PermissionError:
        destino = destino[:-5] + "_alternativo_" + datetime.now().strftime("%H%M%S_%f") + ".xlsx"
        os.replace(temporario, destino)
    print(f"[RELATÓRIO] Progresso salvo em: {destino}")
    return destino


# =========================
# INÍCIO
# =========================


def obter_ovs_requisicao(page, frame, requisicao, pasta_requisicao, descricao, atividade, contexto_credito=None):
    """NC distingue a modalidade pela lista; a NC simples prioriza a OV na descrição."""
    if classificar_atividade_suporte_venda(atividade) == ATIVIDADE_NOTA_CREDITO:
        if contexto_credito is None:
            raise ValidacaoCreditoError("Contexto da conferência de NC não inicializado.")
        ovs, origem, validacao = preparar_validacao_credito(
            page, frame, requisicao, pasta_requisicao, descricao, atividade)
        contexto_credito.clear()
        contexto_credito.update(validacao or {})
        return ovs, origem
    ovs = extrair_ovs_suporte_venda(descricao)
    if ovs:
        return ovs, "PORTAL_DESCRICAO"
    ovs = extrair_ovs_tabela_descricao(descricao)
    if ovs:
        print(
            f"[PORTAL] Tabela BP/cliente/faturamento/OV: {len(ovs)} candidato(s). "
            "Cada OV será confirmada na VA03 antes de executar a VF01."
        )
        return ovs, "PORTAL_DESCRICAO_TABELA"
    ovs = extrair_candidatos_ov_descricao(descricao, atividade)
    if ovs:
        print(f'[PORTAL] {len(ovs)} candidato(s) da descrição; confirmação obrigatória na VA03 antes da VF01.')
        return ovs, 'PORTAL_DESCRICAO_CANDIDATA'
    # Uma cópia local pode ser de uma versão anterior do anexo. Quando a
    # Locação depende do Excel, obter o arquivo atual no PORTAL nesta rodada.
    # Falha de download não autoriza substituir o anexo por um cache antigo.
    print(f'[PORTAL EXCEL] Requisição {requisicao}: consultando o anexo atual no PORTAL; '
          'cópias locais de rodadas anteriores não serão usadas para extrair OVs.')
    _salvos, ovs, referencias = baixar_planilhas_anexas_portal(
        page, frame, requisicao, pasta_requisicao, atividade=atividade,
    )
    origem = "PORTAL_PLANILHA"
    if ovs:
        for ov in ovs:
            print(f"[PORTAL EXCEL] {origem}: OV {ov} -> {'; '.join(referencias.get(ov, []))}")
        return ovs, origem
    return [], "SEM_OV_LOCALIZADA"

def buscar_ov_planilha_apos_rejeicao(page, requisicao, pasta, atividade, rejeitada):
    """Recuperação limitada de Locação após uma descrição rejeitada pelo SAP.

    Baixa os anexos atuais da MESMA requisição. Não usa arquivos soltos nem
    números fixos; exige uma única OV obtida por rótulo de planilha.
    """
    if FIORI_PENDENCIA_GRAVACAO:
        raise RuntimeError('Há resultado SAP incerto; recuperação de OV bloqueada.')
    frame = abrir_requisicao_portal(page, requisicao)
    validar_requisicao_portal_aberta(frame, requisicao, timeout=15000)
    status = confirmar_status_cabecalho_portal(page, frame, requisicao)
    if not str(status or '').strip() or status_requisicao_portal_nao_acionavel(status):
        raise RuntimeError(f'PORTAL não acionável para recuperar a OV: {status}.')
    inspecao = {}
    arquivos, _ovs, _refs = baixar_planilhas_anexas_portal(
        page, frame, requisicao, pasta, atividade=atividade, baixar_todas=True, inspecao=inspecao)
    if (not inspecao.get('completa') or not arquivos
            or inspecao.get('qtd_excel') != len(arquivos)):
        raise RuntimeError('Não foi possível obter todos os Excels atuais da requisição. '
                           + str(inspecao.get('erro') or 'Nenhum Excel disponível.'))
    candidatos = {}
    for caminho in arquivos:
        # Uma planilha protegida/ilegível não pode ser ignorada silenciosamente.
        excel = _abrir_excel_file(caminho)
        try:
            for aba in excel.sheet_names:
                if not _aba_excel_deve_ser_ignorada(aba):
                    _ler_planilha_dataframe(caminho, aba)
        finally:
            excel.close()
        ovs, referencias = extrair_ovs_de_planilha(caminho, atividade=atividade)
        for ov in ovs:
            if _normalizar_numero_documento_sap(ov) == _normalizar_numero_documento_sap(rejeitada):
                continue
            refs = [r for r in referencias.get(ov, [])
                    if 'campo associado ao rótulo' in r or 'rótulo/contexto' in r]
            if refs:
                candidatos.setdefault(_normalizar_numero_documento_sap(ov), []).extend(refs)
    if len(candidatos) != 1:
        raise RuntimeError(f'Esperada uma OV alternativa inequívoca nas planilhas; encontradas {len(candidatos)}. '
                           'Conferir a descrição e os anexos da requisição.')
    nova = next(iter(candidatos))
    referencias = '; '.join(candidatos[nova])
    print(f'[PORTAL OV RECUPERAÇÃO] PORTAL {requisicao}: número da descrição {rejeitada} rejeitado; '
          f'OV candidata {nova} encontrada em {referencias}. Será confirmada na VA03 antes da VF01.')
    return nova, referencias


def executar_rodada(context, resultados):
    page = context.new_page()
    configurar_dialogs_seguros(page)

    print("Acessando o PORTAL...")

    page.goto(
        PORTAL_URL,
        wait_until="domcontentloaded",
        timeout=60000
    )

    print("Aguardando login manual no PORTAL...")

    if not aguardar_login_manual_portal(page):
        raise RuntimeError("Login manual do PORTAL não confirmado.")

    page.wait_for_timeout(2000)

    if os.environ.get("PORTAL_PERMITIR_SCREENSHOT_LOGIN", "").upper() == "SIM":
        try:
            caminho_debug_login = os.path.join(PASTA_RESUMO, f"debug_login_{ID_RODADA}.png")
            page.screenshot(path=caminho_debug_login, full_page=True)
            print(f"Print diagnóstico salvo localmente: {caminho_debug_login}")
        except Exception as e:
            print(f"Não consegui salvar screenshot de diagnóstico: {e}")

    print("\nAbrindo o SAP Fiori em uma segunda aba...")
    page_fiori = context.new_page()
    page_fiori.goto(
        FIORI_URL,
        wait_until="domcontentloaded",
        timeout=TEMPO_FIORI_NAVEGACAO,
    )

    if not aguardar_login_fiori_sem_enter(page_fiori):
        print("[ERRO] Login do SAP Fiori não confirmado. Encerrando a execução.")
        raise RuntimeError("Login do SAP Fiori não confirmado.")

    print(
        "[FIORI] Validação inicial concluída. Durante a fila, o robô fará "
        "somente recuperação automática de navegação, sem novo login."
    )

    page.bring_to_front()

    # V01: captura somente as duas atividades-alvo da fila Operações de Faturamento.
    try:
        requisicoes = capturar_requisicoes_suporte_venda(page)
    except Exception as e:
        print(f"\n[ERRO] Não consegui acessar/capturar a fila Equipe de Operações - Faturamento: {e}")
        try:
            page.screenshot(
                path=os.path.join(
                    PASTA_RESUMO,
                    f"erro_captura_fila_suporte_venda_{ID_RODADA}.png",
                ),
                full_page=True
            )
        except Exception:
            pass
        raise RuntimeError("Falha ao capturar a fila do PORTAL.") from e

    execucao_cancelada_usuario = False

    requisicoes = incluir_pendencias_conciliacao_nc(requisicoes)

    if not requisicoes:
        print("\n[AVISO] Nenhuma requisição das atividades-alvo foi localizada na fila Operações de Faturamento.")
    else:
        # Mantém a melhoria da AV8: antes de iniciar o SAP, permite processar
        # toda a fila ou somente as requisições escolhidas pelo usuário.
        requisicoes, execucao_cancelada_usuario = selecionar_requisicoes_para_processar(
            requisicoes
        )

        if not execucao_cancelada_usuario:
            print(
                f"\nSucesso! {len(requisicoes)} requisição(ões) selecionada(s) "
                "será(ão) processada(s), mantendo a ordem original da fila."
            )

    resultados.extend({
        "Requisicao": req,
        "Atividade": ATIVIDADE_POR_REQUISICAO.get(str(req).strip(), ""),
        "Descricao": "", "ConsultasSAP": [], "StatusPORTAL": "NAO_PROCESSADA",
        "QtdPDFsAnexados": 0, "MensagemPORTAL": "Aguardando processamento nesta rodada.",
    } for req in requisicoes)
    montar_excel_resultado(resultados)

    for registro in resultados:
        req = registro["Requisicao"]
        # Se uma gravação SAP anterior ficou incerta, não faz sentido abrir novas
        # requisições: nenhuma nova escrita é permitida até conferência manual.
        if FIORI_PENDENCIA_GRAVACAO:
            print(
                "[FIORI] Fila interrompida: existe resultado VF01 incerto. "
                "O Excel será gerado com o que já foi processado."
            )
            for pendente in resultados:
                if pendente["StatusPORTAL"] == "NAO_PROCESSADA":
                    pendente.update(StatusPORTAL="NAO_PROCESSADA_ESTADO_SAP_INCERTO",
                                    MensagemPORTAL="Fila interrompida após resultado SAP incerto; requisição não aberta.")
            break

        pasta_req = criar_pasta_requisicao(req)
        inicio_req_dt = datetime.now()
        atividade = ""
        texto = ""
        sequencias = []
        consultas_sap = registro["ConsultasSAP"]
        origem_ovs = ""
        contexto_credito = {}
        etapa_requisicao = "PORTAL_ABERTURA"
        trava_req = None
        registro.update(PastaRequisicao=pasta_req, StatusPORTAL="EM_ANDAMENTO",
                        MensagemPORTAL="Processamento iniciado.",
                        InicioRequisicao=inicio_req_dt.strftime("%d/%m/%Y %H:%M:%S"))

        try:
            trava_req = TravaRequisicao(pasta_req)
            trava_req.__enter__()
            print("\n----------------------------------------")
            print(f"Consultando PORTAL: {req}")
            print(f"Pasta da requisição: {pasta_req}")

            page.bring_to_front()
            atividade = ATIVIDADE_POR_REQUISICAO.get(str(req).strip(), "")
            if not atividade:
                raise RuntimeError(f"Atividade da requisição {req} não foi preservada na captura da fila.")

            print(f"[PORTAL] Atividade da requisição {req}: {atividade}")
            controle_nc = carregar_controle_nc(pasta_req, req)
            if controle_nc and controle_nc['Fase'] != 'VALIDACAO_REPROVADA':
                if classificar_atividade_suporte_venda(atividade) != ATIVIDADE_NOTA_CREDITO:
                    raise ConciliacaoNCError('A atividade atual diverge do controle NC; conferir antes de retomar.')
                texto = controle_nc.get('Descricao', '')
                contexto_credito = controle_nc['Validacao']
                sequencias = [contexto_credito['Ordem']]
                origem_ovs = 'PORTAL_NC_DUAS_PLANILHAS'
                etapa_requisicao = 'NC_RETOMADA'
                registro.update(retomar_nc_duas_planilhas(page, req, pasta_req, consultas_sap))
                etapa_requisicao = registro['Etapa']
                continue
            frame = abrir_requisicao_portal(page, req)

            print("Aguardando abrir a requisição no PORTAL...")

            validar_requisicao_portal_aberta(
                frame,
                req,
                timeout=15000,
            )

            etapa_requisicao = "PORTAL_VALIDACAO_STATUS"
            status_req_portal = confirmar_status_cabecalho_portal(page, frame, req)
            if status_requisicao_portal_nao_acionavel(status_req_portal):
                etapa_requisicao = "PORTAL_STATUS_NAO_ACIONAVEL"
                registro.update(
                    StatusPORTAL="IGNORADA_STATUS_PORTAL",
                    MensagemPORTAL=(
                        f"Requisição não processada porque o PORTAL está em estado "
                        f"não acionável: {status_req_portal}. Nenhuma escrita SAP/PORTAL foi executada."
                    ),
                    Etapa=etapa_requisicao,
                )
                print(
                    f"[PORTAL] Requisição {req} ignorada com segurança: "
                    f"Status={status_req_portal}. SAP não será executado."
                )
                continue

            iframe_texto = (
                frame.frame_locator(FRAME_PORTAL_MAIN)
                .frame_locator(FRAME_EDITOR_TEXTO)
            )

            texto = iframe_texto.locator("body").inner_text(timeout=10000)
            etapa_requisicao = "PORTAL_EXTRACAO_OV"

            sequencias, origem_ovs = obter_ovs_requisicao(
                page, frame, req, pasta_req, texto, atividade, contexto_credito=contexto_credito)
            print(f"[PORTAL] Requisição {req}: {len(sequencias)} OV(s); origem={origem_ovs}; OVs={sequencias}")

            salvar_dados_requisicao(
                requisicao=req,
                descricao=texto,
                ordens=sequencias,
            )

            consultas_sap.extend({
                "Ordem": ordem, "StatusSAP": "NAO_PROCESSADA", "NumeroNF": "",
                "ArquivoPDF": "", "Mensagem": "OV identificada; SAP ainda não executado.",
            } for ordem in sequencias)
            registro.update(Descricao=texto, Atividade=atividade, OrigemOVs=origem_ovs,
                            ExtracaoConcluida=True, Etapa="PORTAL_EXTRACAO_OV")
            montar_excel_resultado(resultados)

            if not sequencias:
                registro.update(
                    StatusPORTAL="SEM_OV_LOCALIZADA", QtdPDFsAnexados=0,
                    MensagemPORTAL="Nenhuma OV extraída com segurança da descrição ou dos Excels locais/anexos tentados. SAP e finalização PORTAL não executados.",
                )
                print(f"[PORTAL] Requisição {req}: SEM_OV_LOCALIZADA. {registro['MensagemPORTAL']}")
                continue

            if contexto_credito and origem_ovs == 'PORTAL_NC_DUAS_PLANILHAS':
                etapa_requisicao = 'NC_PROCESSAMENTO'
                registro.update(executar_nc_duas_planilhas(
                    page, page_fiori, req, pasta_req, atividade, texto, contexto_credito, consultas_sap))
                etapa_requisicao = registro['Etapa']
                continue

            if sequencias and not FIORI_PENDENCIA_GRAVACAO:
                etapa_requisicao = "SAP_PREPARACAO"
                preparar_fiori_nova_requisicao(
                    page_fiori=page_fiori,
                    requisicao=req,
                )

            etapa_requisicao = "SAP_PROCESSAMENTO"
            for indice_ov, ordem in enumerate(sequencias):
                registro["Etapa"] = etapa_requisicao
                consultas_sap[indice_ov].update(StatusSAP="EM_PROCESSAMENTO", Mensagem="Consulta SAP iniciada.")
                try:
                    resultado_sap = processar_ov_com_retomada(
                        page_fiori=page_fiori,
                        requisicao=req,
                        ordem=ordem,
                        atividade=atividade,
                        pasta_requisicao=pasta_req,
                        validar_tabela=(origem_ovs in {"PORTAL_DESCRICAO_TABELA", "PORTAL_DESCRICAO_CANDIDATA"}),
                        validacao_credito=contexto_credito or None,
                    )
                    validar_associacao_resultado_sap(
                        requisicao=req,
                        ordem=ordem,
                        resultado_sap=resultado_sap,
                    )
                except Exception as erro_ov:
                    # Uma falha técnica inesperada em uma OV não pode cancelar
                    # as demais OVs da requisição nem o restante da fila.
                    evidencia = salvar_evidencia_fiori(
                        page_fiori,
                        prefixo="erro_processamento_ov_fiori",
                        requisicao=req,
                        ordem=ordem,
                    )
                    mensagem_erro = f"Erro inesperado ao processar a OV: {erro_ov}"
                    if evidencia:
                        mensagem_erro += f" Evidência: {os.path.basename(evidencia)}"

                    resultado_sap = {
                        "Ordem": str(ordem).strip(),
                        "StatusSAP": "VF01_RESULTADO_NAO_CONFIRMADO" if FIORI_PENDENCIA_GRAVACAO else "ERRO_PROCESSAMENTO_OV",
                        "NumeroNF": "",
                        "ArquivoPDF": "",
                        "Mensagem": mensagem_erro,
                    }
                    print(
                        f"[FIORI] OV {ordem}: falha tratada; "
                        "seguindo para a próxima OV."
                    )

                consultas_sap[indice_ov] = resultado_sap
                montar_excel_resultado(resultados)
                salvar_dados_requisicao(
                    requisicao=req,
                    descricao=texto,
                    ordens=sequencias,
                    consultas_sap=consultas_sap,
                )

                if FIORI_PENDENCIA_GRAVACAO:
                    print(
                        f"[FIORI] Requisição {req}: estado SAP incerto detectado; "
                        "não serão processadas outras OVs desta requisição."
                    )
                    break

            # Recuperação SOMENTE do caso observado: Locação, uma OV informada
            # na descrição, rejeição explícita sem gravação. Não altera a
            # prioridade normal das fontes nem a validação das duas planilhas.
            if (classificar_atividade_suporte_venda(atividade) == ATIVIDADE_DEVOLUCAO_REFATURAMENTO
                    and origem_ovs == 'PORTAL_DESCRICAO' and len(sequencias) == 1
                    and len(consultas_sap) == 1 and not FIORI_PENDENCIA_GRAVACAO
                    and consultas_sap[0].get('StatusSAP') == 'DOCUMENTO_INEXISTENTE'
                    and consultas_sap[0].get('RejeicaoSemGravacao') is True):
                etapa_requisicao = 'PORTAL_RECUPERACAO_OV'
                rejeitada = sequencias[0]
                try:
                    nova, referencias = buscar_ov_planilha_apos_rejeicao(
                        page, req, pasta_req, atividade, rejeitada)
                except Exception as erro_recuperacao:
                    motivo = f'Número {rejeitada} da descrição rejeitado pelo SAP. Recuperação pendente: {erro_recuperacao}'
                    consultas_sap[0]['Mensagem'] += ' | ' + motivo
                    registro.update(StatusPORTAL='PENDENTE_CORRECAO_OV', MensagemPORTAL=motivo, QtdPDFsAnexados=0)
                    print(f'[PORTAL OV RECUPERAÇÃO] {motivo}')
                    continue
                correcao = (f'Número da descrição {rejeitada} rejeitado sem gravação. '
                            f'OV recuperada da planilha: {nova}. Origem: {referencias}')
                sequencias[:] = [nova]
                consultas_sap[:] = [{'Ordem': nova, 'StatusSAP': 'EM_PROCESSAMENTO',
                                    'NumeroNF': '', 'ArquivoPDF': '', 'Mensagem': correcao}]
                origem_ovs = 'PORTAL_PLANILHA_APOS_REJEICAO_DESCRICAO'
                registro.update(OrigemOVs=origem_ovs, CorrecaoOV=correcao, Etapa='SAP_RECUPERACAO_OV')
                salvar_dados_requisicao(req, texto, sequencias, consultas_sap=consultas_sap)
                montar_excel_resultado(resultados)
                etapa_requisicao = 'SAP_RECUPERACAO_OV'
                # Consulta a identidade da OV candidata. Não confere preço/BP
                # de NC simples e não há segunda recuperação automática.
                recuperada = processar_ov_com_retomada(page_fiori, req, nova, atividade, pasta_req,
                                                     validar_tabela=True)
                validar_associacao_resultado_sap(req, nova, recuperada)
                recuperada['Mensagem'] = correcao + ' | ' + str(recuperada.get('Mensagem', ''))
                consultas_sap[:] = [recuperada]
                salvar_dados_requisicao(req, texto, sequencias, consultas_sap=consultas_sap)
                montar_excel_resultado(resultados)

            if any(c.get("StatusSAP") == "NC_VALIDACAO_REPROVADA" for c in consultas_sap):
                etapa_requisicao = "NC_VALIDACAO_PREVIA"
                motivo = next(c["Mensagem"] for c in consultas_sap if c.get("StatusSAP") == "NC_VALIDACAO_REPROVADA")
                registro.update(StatusPORTAL="PENDENTE_VALIDACAO_NC", MensagemPORTAL=motivo,
                                QtdPDFsAnexados=0, ValidacaoNC=motivo)
                print(f"[NC PRÉ-VALIDAÇÃO] Requisição {req} pulada: {motivo}")
                continue
            if contexto_credito:
                registro["ValidacaoNC"] = contexto_credito.get("ResultadoVA03") or "; ".join(c.get("ValidacaoNC", "") for c in consultas_sap)

            etapa_requisicao = "PORTAL_FINALIZACAO"
            resultado_portal = anexar_pdfs_e_encerrar_portal(
                page=page,
                requisicao=req,
                pasta_requisicao=pasta_req,
                consultas_sap=consultas_sap
            )
            registro.update(resultado_portal)
            if registro.get('CorrecaoOV'):
                registro['MensagemPORTAL'] = str(registro.get('MensagemPORTAL', '')) + ' | ' + registro['CorrecaoOV']

        except RequisicaoEmExecucaoError as erro_trava:
            etapa_requisicao = 'REQUISICAO_EM_EXECUCAO'
            registro.update(StatusPORTAL='REQUISICAO_EM_EXECUCAO', MensagemPORTAL=str(erro_trava))
            print(f'[PORTAL TRAVA] PORTAL {req}: {erro_trava}')
        except DocumentosPORTALIndisponiveis as erro_documentos:
            etapa_requisicao = 'PORTAL_DOCUMENTOS'
            registro.update(StatusPORTAL='ERRO_CARREGAMENTO_DOCUMENTOS', MensagemPORTAL=str(erro_documentos))
            print(f'[PORTAL DOCUMENTOS] PORTAL {req}: {erro_documentos}')
        except ConciliacaoNCError as erro_nc:
            etapa_requisicao = 'NC_PENDENTE_CONFERENCIA'
            registro.update(StatusPORTAL='PENDENTE_NC', MensagemPORTAL=str(erro_nc),
                            StatusConciliacao='PENDENTE_CONFERENCIA', MensagemConciliacao=str(erro_nc))
            print(f'[NC CONTROLE] PORTAL {req}: {erro_nc}')
        except IdentificacaoModalidadeNCError as erro_modalidade:
            etapa_requisicao = 'NC_IDENTIFICACAO_MODALIDADE'
            registro.update(StatusPORTAL='PENDENTE_IDENTIFICACAO_NC',
                            MensagemPORTAL=str(erro_modalidade), QtdPDFsAnexados=0)
            print(f'[NC MODALIDADE] PORTAL {req}: {erro_modalidade}')
        except ValidacaoCreditoError as erro_validacao:
            etapa_requisicao = "NC_VALIDACAO_PREVIA"
            motivo = f"Requisição pulada antes da VF01: {erro_validacao}"
            registro.update(StatusPORTAL="PENDENTE_VALIDACAO_NC", MensagemPORTAL=motivo,
                            QtdPDFsAnexados=0, ValidacaoNC=motivo)
            print(f"[NC PRÉ-VALIDAÇÃO] PORTAL {req}: {motivo}")
        except StatusPORTALIndeterminado as erro_status:
            etapa_requisicao = "PORTAL_VALIDACAO_STATUS"
            registro.update(StatusPORTAL="STATUS_PORTAL_INDETERMINADO",
                            MensagemPORTAL=str(erro_status), Etapa=etapa_requisicao)
        except KeyboardInterrupt:
            registro.update(StatusPORTAL="INTERROMPIDO", MensagemPORTAL=f"Interrupção manual na etapa {etapa_requisicao}.")
            raise
        except Exception as e:
            erro_sap = str(etapa_requisicao).startswith("SAP_")
            tipo_erro = "SAP" if erro_sap else "PORTAL"
            print(f"[-] {req}: erro na etapa {etapa_requisicao} ({tipo_erro})")
            print(f"Detalhe do erro: {e}")

            caminho_print = os.path.join(
                pasta_req,
                f"erro_{tipo_erro.lower()}_{req}_{ID_RODADA}.png",
            )
            try:
                pagina_print = page_fiori if erro_sap else page
                pagina_print.screenshot(path=caminho_print, full_page=True)
                print(f"Print do erro salvo em: {caminho_print}")
            except Exception as erro_print:
                print(f"Não consegui salvar o print do erro: {erro_print}")

            registro.update(
                StatusPORTAL="ERRO_SAP_REQUISICAO" if erro_sap else "ERRO_PORTAL",
                QtdPDFsAnexados=0, MensagemPORTAL=f"{etapa_requisicao}: {e}",
            )
        finally:
            try:
                registro.update(Descricao=texto, Atividade=atividade, Etapa=etapa_requisicao,
                                OrigemOVs=origem_ovs)
                for consulta in consultas_sap:
                    if consulta.get("StatusSAP") not in {"NAO_PROCESSADA", "EM_PROCESSAMENTO"}:
                        continue
                    em_andamento = consulta["StatusSAP"] == "EM_PROCESSAMENTO"
                    if FIORI_PENDENCIA_GRAVACAO:
                        status = "VF01_RESULTADO_NAO_CONFIRMADO" if em_andamento else "NAO_PROCESSADA_ESTADO_SAP_INCERTO"
                    else:
                        status = "INTERROMPIDA" if em_andamento else "NAO_PROCESSADA"
                    consulta.update(StatusSAP=status, Mensagem=registro.get("MensagemPORTAL", "") or
                                    f"Processamento encerrado na etapa {etapa_requisicao}.")
                registro.update(montar_campos_tempo_requisicao(
                    inicio_req_dt, datetime.now(), requisicao_valida_para_medicao(registro, consultas_sap)))
                if trava_req is not None and trava_req.arquivo is not None:
                    salvar_dados_requisicao(req, texto, sequencias, consultas_sap, registro)
                imprimir_situacao_conciliacao(registro)
                montar_excel_resultado(resultados)
            finally:
                if trava_req is not None:
                    trava_req.__exit__(None, None, None)


def imprimir_resumo_final_execucao(resultados, caminho_relatorio=""):
    """Resumo operacional com motivos separados do TESTE 46."""
    concluidas = [
        item for item in (resultados or [])
        if str(item.get("StatusPORTAL", "") or "").upper().strip() == "CONCLUIDO"
    ]
    ignoradas = [
        item for item in (resultados or [])
        if str(item.get("StatusPORTAL", "") or "").upper().strip() == "IGNORADA_STATUS_PORTAL"
    ]
    processos_concluidos = [item for item in concluidas
        if campos_situacao_conciliacao(item)['Processo concluído?'] == 'SIM']
    nao_concluidas = [
        item for item in (resultados or [])
        if item not in processos_concluidos and item not in ignoradas
    ]

    print("\n" + "=" * 72)
    print("RESUMO FINAL DA RODADA")
    print("=" * 72)
    print(f"Requisições registradas: {len(resultados or [])}")
    novas = sum(campos_situacao_conciliacao(item)['PORTAL encerrado nesta rodada?'] == 'SIM'
                for item in resultados or [])
    print(f"PORTALs encerrados nesta rodada: {novas}")
    print(f"PORTALs originais com encerramento confirmado (inclui anteriores): {len(concluidas)}")
    print(f"Processos totalmente concluídos: {len(processos_concluidos)}")
    print(f"Ignoradas por status encerrado/cancelado: {len(ignoradas)}")
    print(f"Pendentes / com falha: {len(nao_concluidas)}")
    for status in ("STATUS_PORTAL_INDETERMINADO", "SEM_OV_LOCALIZADA", "PENDENTE_VALIDACAO_NC", "ERRO_ANEXO"):
        quantidade = sum(item.get("StatusPORTAL") == status for item in nao_concluidas)
        print(f"  {status}: {quantidade}")

    if ignoradas:
        print("Ignoradas com segurança:")
        for item in ignoradas:
            req = str(item.get("Requisicao", "") or "<sem número>")
            mensagem = re.sub(r"\s+", " ", str(item.get("MensagemPORTAL", "") or "")).strip()
            print(f"- {req}: {mensagem[:220] or 'status PORTAL não acionável'}")

    if nao_concluidas:
        print("Pendências:")
        for item in nao_concluidas:
            req = str(item.get("Requisicao", "") or "<sem número>")
            status_portal = str(item.get("StatusPORTAL", "") or "SEM_STATUS")
            consultas = item.get("ConsultasSAP") or []
            status_sap = []
            for consulta in consultas:
                status = str(consulta.get("StatusSAP", "") or "").strip()
                if status and status != "OK_VA03" and status not in status_sap:
                    status_sap.append(status)
            motivo = status_portal
            if status_sap:
                motivo += " | SAP: " + ", ".join(status_sap)
            mensagem = re.sub(r"\s+", " ", str(item.get("MensagemPORTAL", "") or "")).strip()
            if mensagem:
                motivo += " | " + mensagem[:220]
            print(f"- {req}: {motivo}")

    conciliacoes = [r for r in resultados or [] if r.get('StatusConciliacao')]
    if conciliacoes:
        concluidas_nc = sum(r['StatusConciliacao'] == 'CONCLUIDA'
                           and campos_situacao_conciliacao(r)['Conciliação criada?'] == 'SIM' for r in conciliacoes)
        print(f'Conciliações NC confirmadas: {concluidas_nc}/{len(conciliacoes)}')
        for r in conciliacoes:
            campos = campos_situacao_conciliacao(r)
            print(f"- PORTAL {r['Requisicao']} -> Criada={campos['Conciliação criada?']}; "
                  f"PORTAL conciliação={campos['PORTAL da conciliação'] or 'sem número confirmado'}: "
                  f"{r['StatusConciliacao']} | {r.get('MensagemConciliacao', '')}")

    if caminho_relatorio:
        print(f"Relatório final: {caminho_relatorio}")
    print("=" * 72)

def main():
    """Importar este arquivo não abre navegador, não cria pastas e não inicia faturamento."""
    # Proteção para impedir ações em portais reais ao executar um exemplo público.
    if os.environ.get("AUTORIZAR_EXECUCAO_REAL", "").strip().upper() != "SIM":
        print("Versão pública: execução real bloqueada por padrão.")
        print("Configure um ambiente autorizado e habilite AUTORIZAR_EXECUCAO_REAL=SIM.")
        return 2
    necessarias = {
        "ERP_FIORI_URL": FIORI_URL,
        "PORTAL_URL": PORTAL_URL,
        "PORTAL_URL_MENU": PORTAL_URL_MENU,
        "PORTAL_URL_ATENDIMENTO": PORTAL_URL_ATENDIMENTO,
    }
    faltantes = [chave for chave, valor in necessarias.items() if not valor]
    if faltantes:
        print("Configurações de ambiente ausentes: " + ", ".join(faltantes))
        return 2

    os.makedirs(PASTA_REQUISICOES, exist_ok=True)
    os.makedirs(PASTA_RESUMO, exist_ok=True)
    caminho_log = iniciar_log_execucao()
    print(f"VERSÃO TÉCNICA: {VERSAO_ROBO}")
    print('[NC CONCILIAÇÃO] Modo automático: após validação, faturamento, PDF e encerramento '
          'confirmados, cria a conciliação de Glosa com a planilha validada na mesma execução, '
          'sem confirmação manual. O erro legado de número inválido admite uma única tentativa '
          'adicional registrada quando o menu de consulta não é reconhecido; VF01 não é repetida.')
    print(f"ARQUIVO EXECUTADO: {os.path.abspath(__file__)}")
    print(f"NAVEGADOR: {NOME_NAVEGADOR} (channel={CANAL_NAVEGADOR})")
    print(f"Pasta de requisições: {PASTA_REQUISICOES}")
    print(f"Pasta de resumo: {PASTA_RESUMO}")
    print(f"Log desta rodada: {caminho_log}")
    resultados = []
    codigo_saida = 0
    with sync_playwright() as p:
        browser = p.chromium.launch(channel=CANAL_NAVEGADOR, headless=False,
                                    args=["--start-maximized"])
        context = browser.new_context(accept_downloads=True, no_viewport=True)
        try:
            executar_rodada(context, resultados)
        except KeyboardInterrupt:
            codigo_saida = 130
            print("\nInterrupção manual: preservando resultados desta rodada.")
        except Exception as erro:
            codigo_saida = 1
            print(f"[ERRO EXECUÇÃO] {type(erro).__name__}: {erro}")
        finally:
            for pendente in resultados:
                if pendente.get("StatusPORTAL") == "NAO_PROCESSADA":
                    pendente["MensagemPORTAL"] = "Requisição não iniciada antes do encerramento da rodada."
                    if FIORI_PENDENCIA_GRAVACAO:
                        pendente["StatusPORTAL"] = "NAO_PROCESSADA_ESTADO_SAP_INCERTO"
            caminho_relatorio_final = ""
            try:
                caminho_relatorio_final = montar_excel_resultado(resultados)
            except Exception as erro_relatorio:
                codigo_saida = 1
                print(f"[ERRO RELATÓRIO] Não foi possível salvar o Excel: {erro_relatorio}")
            imprimir_resumo_final_execucao(resultados, caminho_relatorio_final)
            try:
                input("\nChrome aberto para conferência. Pressione ENTER para fechar...")
            except (KeyboardInterrupt, EOFError):
                pass
            context.close()
            browser.close()
    return codigo_saida


if __name__ == "__main__":
    sys.exit(main())
