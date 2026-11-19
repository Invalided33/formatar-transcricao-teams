"""Formata as falas do Teams e remove figuras DrawingML incorporadas no corpo do Word."""

import argparse
import posixpath
import re
from copy import copy
from io import BytesIO
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as ET
from xml.parsers import expat
from zipfile import BadZipFile, ZipFile

TIPOS = "[Content_Types].xml"
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
ET.register_namespace("w", NS["w"])
TEMPO = re.compile(r"(?:\d+:)?\d+:[0-5]\d")
TAG_INICIAL = re.compile(rb"""<(?:[^>"']|"[^"]*"|'[^']*')+>""")


def falas(raiz):
    """Reconhece tanto a exportação do Teams quanto as cópias geradas pelo script."""
    for paragrafo in raiz.findall(".//w:p", NS):
        todos = paragrafo.findall("w:r", NS)
        trechos = [r for r in todos if r.find("w:t", NS) is not None]
        if len(trechos) < 3:
            continue
        hora_primeiro = bool(TEMPO.fullmatch(trechos[0].findtext("w:t", "", NS).strip()))
        inicio_nome = 1 if hora_primeiro else 0
        nomes = []
        for trecho in trechos[inicio_nome:]:
            if trecho.find("w:rPr/w:b", NS) is None:
                break
            if nomes and TEMPO.fullmatch(trecho.findtext("w:t", "", NS).strip()):
                break
            nomes.append(trecho)
        proximo = inicio_nome + len(nomes)
        if not nomes or proximo >= len(trechos):
            continue
        if hora_primeiro:
            horario, primeira_linha = trechos[0], trechos[proximo]
            fim_cabecalho = nomes[-1]
        elif TEMPO.fullmatch(trechos[proximo].findtext("w:t", "", NS).strip()):
            if proximo + 1 == len(trechos):
                continue
            horario, primeira_linha = trechos[proximo:proximo + 2]
            fim_cabecalho = horario
        else:
            # Em uma versão anterior o horário vinha depois da última linha de fala.
            horario, primeira_linha = trechos[-1], trechos[proximo]
            fim_cabecalho = nomes[-1]
        texto_hora = horario.find("w:t", NS)
        if texto_hora is None or not TEMPO.fullmatch((texto_hora.text or "").strip()):
            continue
        corpo = [r for r in todos[todos.index(fim_cabecalho) + 1:] if r is not horario]
        # O Word pode armazenar a quebra em um run próprio antes da primeira palavra.
        inicio_fala = corpo[:corpo.index(primeira_linha) + 1]
        if hora_primeiro or any(r.find("w:br", NS) is not None for r in inicio_fala):
            yield nomes, horario, texto_hora, corpo


def juntar_linhas(trechos):
    """Remove as quebras da fala e ajusta o espaço no início e entre os trechos."""
    w = "{" + NS["w"] + "}"
    anterior = ""
    quebras = set()
    substituicoes = {}
    entre_textos = False
    for trecho in trechos:
        for no in trecho:
            if no.tag in (w + "br", w + "cr") and no.get(w + "type", "textWrapping") == "textWrapping":
                quebras.add(no)
                entre_textos = True
            elif no.tag == w + "t" and no.text:
                valor = no.text
                if not anterior:
                    valor = valor.lstrip()
                elif entre_textos and not anterior[-1].isspace() and not valor[0].isspace():
                    valor = " " + valor
                if valor != no.text:
                    substituicoes[no] = valor
                anterior = valor
                entre_textos = False
            elif no.tag == w + "tab":
                anterior = "\t"
                entre_textos = False
    return quebras, substituicoes


def figuras_do_corpo(raiz, relacoes):
    """Enumera as figuras DrawingML incorporadas no corpo principal do Word."""
    ids_de_imagem = {
        r.get("Id") for r in relacoes
        if r.get("Type") == NS["r"] + "/image" and r.get("TargetMode") != "External"
    }
    corpo = raiz.find("w:body", NS)
    if corpo is None:
        raise ValueError("O arquivo não contém o corpo de uma transcrição do Teams em Word.")
    for desenho in corpo.findall(".//w:r/w:drawing", NS):
        if len(desenho) != 1:
            continue
        quadro = desenho[0]
        if quadro.tag not in (f"{{{NS['wp']}}}inline", f"{{{NS['wp']}}}anchor"):
            continue
        objeto = quadro.find("a:graphic/a:graphicData", NS)
        if objeto is None or objeto.get("uri") != NS["pic"]:
            continue
        imagem = objeto.find("pic:pic/pic:blipFill/a:blip", NS)
        if imagem is not None and imagem.get(f"{{{NS['r']}}}embed") in ids_de_imagem:
            yield desenho


def id_do_desenho(desenho):
    propriedades = desenho.find("*/wp:docPr", NS)
    return propriedades.get("id") if propriedades is not None else None


def listar_figuras(entrada):
    with ZipFile(entrada) as pacote:
        documento, arquivo_rels = partes_principais(pacote)
        raiz = ET.fromstring(pacote.read(documento))
        relacoes = ET.fromstring(pacote.read(arquivo_rels)) if arquivo_rels in pacote.namelist() else ()
        alvos = {r.get("Id"): r.get("Target") for r in relacoes}
        contextos = {}
        for p in raiz.findall(".//w:p", NS):
            for desenho in p.findall(".//w:drawing", NS):
                contextos[desenho] = " ".join("".join(p.itertext()).split())[:80]
        for desenho in figuras_do_corpo(raiz, relacoes):
            ref = desenho.find(".//pic:blipFill/a:blip", NS).get(f"{{{NS['r']}}}embed")
            print(f"ID {id_do_desenho(desenho)}  {alvos[ref]}  {contextos.get(desenho, '')}")


def intervalos_xml(xml, raiz, selecionados):
    """Obtém posições no XML original para editar sem regravar a formatação."""
    elementos = iter(raiz.iter())
    pilha = []
    intervalos = {}
    parser = expat.ParserCreate()

    def abrir(nome, atributos):
        elemento = next(elementos)
        pilha.append((elemento, parser.CurrentByteIndex) if elemento in selecionados else None)

    def fechar(nome):
        registro = pilha.pop()
        if registro is None:
            return
        elemento, inicio = registro
        fim = parser.CurrentByteIndex
        # Para <tag/>, o Expat já aponta para depois da tag. Para <tag>…</tag>,
        # ele aponta para o início de </tag>.
        if not TAG_INICIAL.match(xml, inicio).group().endswith(b"/>"):
            fim = xml.index(b">", fim) + 1
        intervalos[elemento] = (inicio, fim)

    parser.StartElementHandler = abrir
    parser.EndElementHandler = fechar
    parser.Parse(xml, True)
    return intervalos


def aplicar_edicoes(xml, edicoes):
    partes = []
    anterior = 0
    for inicio, fim, substituto in sorted(edicoes):
        partes.append(xml[anterior:inicio])
        partes.append(substituto)
        anterior = fim
    partes.append(xml[anterior:])
    return b"".join(partes)


def cortar_elementos(xml, raiz, selecionados):
    if not selecionados:
        return xml
    intervalos = intervalos_xml(xml, raiz, selecionados)
    return aplicar_edicoes(xml, [(i, f, b"") for i, f in intervalos.values()])


def formatar_horario(valor):
    segundos = 0
    for campo in valor.split(":"):
        segundos = segundos * 60 + int(campo)
    horas, resto = divmod(segundos, 3600)
    minutos, segundos = divmod(resto, 60)
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"


def horario_em_negrito(horario, texto, quebra_nome):
    propriedades = horario.find("w:rPr", NS)
    if propriedades is None:
        propriedades = ET.Element(ET.QName(NS["w"], "rPr"))
        horario.insert(0, propriedades)
    for tag in ("b", "bCs"):
        negrito = propriedades.find(f"w:{tag}", NS)
        if negrito is None:
            ET.SubElement(propriedades, ET.QName(NS["w"], tag))
        else:
            negrito.attrib.pop(f"{{{NS['w']}}}val", None)
    if quebra_nome is not None:
        horario.insert(list(horario).index(texto), quebra_nome)
    texto.text = formatar_horario(texto.text.strip()) + "\u00a0"
    horario.tail = None  # A cauda do elemento já está preservada no XML de origem.
    return ET.tostring(horario, encoding="utf-8")


def editar_falas(xml, raiz, figuras):
    entradas = list(falas(raiz))
    if not entradas and not figuras:
        raise ValueError("Nenhuma fala com nome e horário reconhecida no formato esperado do Teams.")

    elementos = figuras.copy()
    quebras = set()
    substituicoes = {}
    for nomes, horario, _, corpo in entradas:
        nome = nomes[0]
        elementos.update((nome, horario))
        elementos.update(nome.findall("w:br", NS))
        internas, ajustes = juntar_linhas(corpo)
        quebras.update(internas)
        substituicoes.update(ajustes)
        textos_nome = [t for r in nomes for t in r.findall("w:t", NS)]
        for texto_nome in reversed(textos_nome):
            locutor = (texto_nome.text or "").rstrip().rstrip(":").rstrip()
            legenda = locutor + ": " if locutor else ""
            if legenda != (texto_nome.text or ""):
                substituicoes[texto_nome] = legenda
            if locutor:
                break
    posicoes = intervalos_xml(xml, raiz, elementos | quebras | substituicoes.keys())
    edicoes = [(*posicoes[no], b"") for no in figuras | quebras]
    for no, valor in substituicoes.items():
        inicio, fim = posicoes[no]
        no.text = valor
        no.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        no.tail = None
        edicoes.append((inicio, fim, ET.tostring(no, encoding="utf-8")))

    for nomes, horario, texto, _ in entradas:
        nome = nomes[0]
        inicio, fim = posicoes[horario]
        quebra = nome.find("w:br", NS)
        if quebra is not None:
            edicoes.append((*posicoes[quebra], b""))
        trecho = horario_em_negrito(horario, texto, quebra)
        inicio_nome = posicoes[nome][0]
        edicoes.extend(((inicio, fim, b""), (inicio_nome, inicio_nome, trecho)))
    return aplicar_edicoes(xml, edicoes), figuras, len(entradas)


def caminho_alvo(arquivo_rels, relacao):
    """Resolve o Target relativo à parte do DOCX que contém a referência."""
    alvo = unquote(relacao.get("Target"))
    pasta = posixpath.dirname(posixpath.dirname(arquivo_rels))
    return posixpath.normpath(posixpath.join(pasta, alvo)).lstrip("/")


def partes_principais(pacote):
    """Descobre o corpo do Word pelo tipo da relação, mesmo se o arquivo for renomeado."""
    origem = "_rels/.rels"
    for r in ET.fromstring(pacote.read(origem)):
        if r.get("Type") == NS["r"] + "/officeDocument" and r.get("TargetMode") != "External":
            documento = caminho_alvo(origem, r)
            pasta, nome = posixpath.split(documento)
            return documento, posixpath.join(pasta, "_rels", nome + ".rels")
    raise ValueError("O pacote não contém o documento principal no formato de exportação do Teams.")


def formatar_transcricao(entrada, saida, ids_figuras=None):
    with ZipFile(entrada) as original:
        documento, arquivo_rels = partes_principais(original)
        xml = original.read(documento)
        raiz = ET.fromstring(xml)
        xml_rels = original.read(arquivo_rels) if arquivo_rels in original.namelist() else None
        rels = ET.fromstring(xml_rels) if xml_rels is not None else ()
        figuras = set(figuras_do_corpo(raiz, rels))
        if ids_figuras is not None:
            selecionadas = set()
            for id_figura in set(ids_figuras):
                ocorrencias = [d for d in figuras if id_do_desenho(d) == id_figura]
                if len(ocorrencias) != 1:
                    raise ValueError(f"O ID {id_figura} identifica {len(ocorrencias)} figuras neste arquivo.")
                selecionadas.add(ocorrencias[0])
            figuras = selecionadas
        novo_xml, figuras, total_horarios = editar_falas(xml, raiz, figuras)
        # Uma figura pode conter uma representação adicional, como um SVG ao lado do PNG.
        ids = {no.get(f"{{{NS['r']}}}embed") for figura in figuras for no in figura.iter()
               if f"{{{NS['r']}}}embed" in no.attrib}
        # Uma imagem também pode ilustrar outra parte do documento.
        for elemento in ET.fromstring(novo_xml).iter():
            for atributo, valor in elemento.attrib.items():
                if atributo.startswith(f"{{{NS['r']}}}"):
                    ids.discard(valor)

        remover_rels = {r for r in rels if r.get("Id") in ids
                        and r.get("Type") == NS["r"] + "/image"
                        and r.get("TargetMode") != "External"}
        midias = {caminho_alvo(arquivo_rels, r) for r in remover_rels}
        for arquivo in original.namelist():
            if not arquivo.endswith(".rels"):
                continue
            if arquivo == arquivo_rels:
                outras = [r for r in rels if r not in remover_rels]
            else:
                outras = ET.fromstring(original.read(arquivo))
            for r in outras:
                if r.get("TargetMode") != "External":
                    midias.discard(caminho_alvo(arquivo, r))

        alterados = {documento: novo_xml}
        if remover_rels:
            alterados[arquivo_rels] = cortar_elementos(xml_rels, rels, remover_rels)
        xml_tipos = original.read(TIPOS)
        tipos = ET.fromstring(xml_tipos)
        exclusivos = {t for t in tipos if unquote(t.get("PartName", "")).lstrip("/") in midias}
        if exclusivos:
            alterados[TIPOS] = cortar_elementos(xml_tipos, tipos, exclusivos)

        pacote = BytesIO()
        with ZipFile(pacote, "w") as novo:
            novo.comment = original.comment
            for item in original.infolist():
                if item.filename not in midias:
                    dados = alterados.get(item.filename)
                    novo.writestr(copy(item), original.read(item) if dados is None else dados)

    # O modo exclusivo também protege o original caso os caminhos coincidam.
    with saida.open("xb") as destino:
        try:
            destino.write(pacote.getbuffer())
            destino.flush()
        except OSError:
            destino.close()
            saida.unlink()
            raise
    return len(figuras), total_horarios


def main():
    parser = argparse.ArgumentParser(
        description=__doc__,
        epilog='Uso normal: python formatar_transcricao.py "reuniao.docx".',
    )
    parser.add_argument("entrada", type=Path, help="Transcrição original em Word")
    parser.add_argument("-o", "--saida", type=Path, help="Caminho do novo DOCX")
    selecao = parser.add_mutually_exclusive_group()
    selecao.add_argument("--listar-figuras", action="store_true", help="Mostra os IDs das figuras e o texto próximo")
    selecao.add_argument("--remover-ids", nargs="+", metavar="ID",
                         help="Remove só as figuras com estes IDs. Por padrão, remove todas as figuras do corpo.")
    args = parser.parse_args()
    saida = args.saida or args.entrada.with_stem(args.entrada.stem + "_formatado")
    try:
        if args.listar_figuras:
            listar_figuras(args.entrada)
            return
        total, horarios = formatar_transcricao(args.entrada, saida, args.remover_ids)
    except FileExistsError:
        parser.exit(1, f"O arquivo de saída já existe: {saida}\n")
    except (OSError, BadZipFile, ValueError, KeyError, ET.ParseError) as erro:
        parser.exit(1, f"Falha ao processar o documento: {erro}\n")
    print(f"{total} imagens removidas e {horarios} horários formatados. Arquivo: {saida}")


if __name__ == "__main__":
    main()
