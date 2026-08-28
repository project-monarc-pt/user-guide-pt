#!/usr/bin/env python3
"""Verificação de integridade de um .docx gerado por adoc2docx.rb.

Deteta os erros que levam o Word a pedir a recuperação do ficheiro e que não
se notam ao ver o documento no LibreOffice: partes sem content type, relações
por resolver, marcadores repetidos ou com nomes ilegais, campos
desequilibrados, estilos ou numerações inexistentes e propriedades do
documento fora do esquema.

    python3 tools/docx/verificar.py _build/user-guide.docx

Valida também cada parte pelos esquemas oficiais (ECMA-376 Parte 4,
Transitional), que são descarregados e guardados em _build/xsd na primeira
utilização. Para usar uma cópia já existente, indique a pasta:

    python3 tools/docx/verificar.py _build/user-guide.docx /caminho/para/xsd
"""
import io
import os
import re
import sys
import urllib.request
import zipfile
from collections import Counter

ECMA_PART4 = ('https://www.ecma-international.org/wp-content/uploads/'
              'ECMA-376-4_5th_edition_december_2016.zip')
XML_XSD = 'https://www.w3.org/2001/xml.xsd'


def obter_esquemas(cache):
    """Descarrega (uma vez) os esquemas Transitional para `cache`.

    Devolve a pasta, ou None se não houver rede — nesse caso só correm as
    verificações de pacote, que não precisam de esquema.
    """
    if os.path.exists(os.path.join(cache, 'wml.xsd')):
        return cache
    try:
        os.makedirs(cache, exist_ok=True)
        with urllib.request.urlopen(ECMA_PART4, timeout=180) as r:
            parte4 = zipfile.ZipFile(io.BytesIO(r.read()))
        nome = next(n for n in parte4.namelist() if 'XMLSchema-Transitional' in n)
        zipfile.ZipFile(io.BytesIO(parte4.read(nome))).extractall(cache)
        with urllib.request.urlopen(XML_XSD, timeout=60) as r:
            open(os.path.join(cache, 'xml.xsd'), 'wb').write(r.read())
        # os esquemas usam xml:space sem importar o namespace XML
        imp = ('  <xsd:import namespace="http://www.w3.org/XML/1998/namespace"'
               ' schemaLocation="xml.xsd"/>\n')
        for f in os.listdir(cache):
            if not f.endswith('.xsd') or f == 'xml.xsd':
                continue
            caminho = os.path.join(cache, f)
            texto = open(caminho, encoding='utf-8').read()
            if 'XML/1998/namespace' not in texto or 'schemaLocation="xml.xsd"' in texto:
                continue
            m = re.search(r'<xsd:schema[^>]*>\n', texto)
            if m:
                open(caminho, 'w', encoding='utf-8').write(
                    texto[:m.end()] + imp + texto[m.end():])
        return cache
    except Exception as exc:                                  # sem rede, etc.
        print('aviso: esquemas indisponíveis (%s); só verificações de pacote'
              % str(exc)[:80])
        return None

REL_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CT_NS = '{http://schemas.openxmlformats.org/package/2006/content-types}'
PARTES_WML = ['word/document.xml', 'word/styles.xml', 'word/numbering.xml',
              'word/settings.xml', 'word/footnotes.xml', 'word/header1.xml',
              'word/header2.xml', 'word/footer1.xml', 'word/footer2.xml']


def verificar(caminho, xsd_dir=None):
    from lxml import etree

    zf = zipfile.ZipFile(caminho)
    nomes = [n for n in zf.namelist() if not n.endswith('/')]
    ler = lambda n: zf.read(n).decode('utf-8')
    problemas = []

    # ---- content types --------------------------------------------------
    ct = etree.fromstring(zf.read('[Content_Types].xml'))
    defaults = {d.get('Extension').lower() for d in ct.iter(CT_NS + 'Default')}
    overrides = {o.get('PartName').lstrip('/') for o in ct.iter(CT_NS + 'Override')}
    for n in nomes:
        if n == '[Content_Types].xml':
            continue
        ext = n.rsplit('.', 1)[-1].lower() if '.' in n else ''
        if n not in overrides and ext not in defaults:
            problemas.append('sem content type: %s' % n)

    # ---- relações -------------------------------------------------------
    for rel in [n for n in nomes if n.endswith('.rels')]:
        base = os.path.dirname(os.path.dirname(rel))
        for r in etree.fromstring(zf.read(rel)).iter(REL_NS + 'Relationship'):
            if r.get('TargetMode') == 'External':
                continue
            alvo = os.path.normpath(os.path.join(base, r.get('Target')))
            if alvo not in nomes:
                problemas.append('%s: alvo inexistente -> %s' % (rel, r.get('Target')))

    for parte in [n for n in nomes if n.startswith('word/') and n.endswith('.xml')]:
        rel = 'word/_rels/%s.rels' % os.path.basename(parte)
        declaradas = set()
        if rel in nomes:
            declaradas = {r.get('Id') for r in
                          etree.fromstring(zf.read(rel)).iter(REL_NS + 'Relationship')}
        usadas = set(re.findall(r'r:(?:id|embed|link)="([^"]+)"', ler(parte)))
        for u in sorted(usadas - declaradas):
            problemas.append('%s: usa %s sem relação declarada' % (parte, u))

    doc = ler('word/document.xml')

    # ---- marcadores (nome: letras, dígitos e _; até 40 caracteres) ------
    inicios = re.findall(r'<w:bookmarkStart w:id="(\d+)" w:name="([^"]+)"/>', doc)
    fins = re.findall(r'<w:bookmarkEnd w:id="(\d+)"/>', doc)
    if len(inicios) != len(fins):
        problemas.append('marcadores: %d inícios vs %d fins' % (len(inicios), len(fins)))
    for nome, n in Counter(n for _i, n in inicios).items():
        if n > 1:
            problemas.append('nome de marcador repetido: %s (%dx)' % (nome, n))
        if len(nome) > 40 or not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', nome):
            problemas.append('nome de marcador inválido: %s' % nome)
    for ident, n in Counter(i for i, _n in inicios).items():
        if n > 1:
            problemas.append('id de marcador repetido: %s (%dx)' % (ident, n))

    # ---- ids de imagem únicos dentro de cada parte ----------------------
    for parte in ['word/document.xml', 'word/header1.xml']:
        if parte not in nomes:
            continue
        for ident, n in Counter(re.findall(r'<wp:docPr id="(\d+)"', ler(parte))).items():
            if n > 1:
                problemas.append('%s: wp:docPr id repetido %s (%dx)' % (parte, ident, n))

    # ---- estilos e numerações referidos existem -------------------------
    estilos = ler('word/styles.xml')
    definidos = set(re.findall(r'<w:style [^>]*w:styleId="([^"]+)"', estilos))
    for parte in [p for p in PARTES_WML if p in nomes] + ['word/footnotes.xml']:
        raw = ler(parte)
        usados = set()
        for tag in ('pStyle', 'rStyle', 'tblStyle'):
            usados |= set(re.findall(r'<w:%s w:val="([^"]+)"/>' % tag, raw))
        for s in sorted(usados - definidos):
            problemas.append('%s: estilo inexistente %s' % (parte, s))
    herdados = set(re.findall(r'<w:basedOn w:val="([^"]+)"/>', estilos)) | \
        set(re.findall(r'<w:next w:val="([^"]+)"/>', estilos))
    for s in sorted(herdados - definidos):
        problemas.append('styles.xml: basedOn/next para estilo inexistente %s' % s)

    num = ler('word/numbering.xml')
    nums = set(re.findall(r'<w:num w:numId="(\d+)"', num))
    abstratos = set(re.findall(r'<w:abstractNum w:abstractNumId="(\d+)"', num))
    for a in sorted(set(re.findall(r'<w:abstractNumId w:val="(\d+)"/>', num)) - abstratos):
        problemas.append('numbering.xml: abstractNumId inexistente %s' % a)
    for parte in ('word/document.xml', 'word/styles.xml'):
        usados = set(re.findall(r'<w:numId w:val="(\d+)"/>', ler(parte))) - {'0'}
        for n in sorted(usados - nums):
            problemas.append('%s: numId inexistente %s' % (parte, n))

    # ---- campos equilibrados e células com parágrafo --------------------
    for parte in [p for p in PARTES_WML if p in nomes]:
        raw = ler(parte)
        b, e = raw.count('w:fldCharType="begin"'), raw.count('w:fldCharType="end"')
        if b != e:
            problemas.append('%s: campos desequilibrados (begin=%d, end=%d)' % (parte, b, e))
    vazias = len(re.findall(r'<w:tc>(?:(?!<w:p[ >]).)*?</w:tc>', doc, re.S))
    if vazias:
        problemas.append('%d células de tabela sem parágrafo' % vazias)

    # ---- validação pelos esquemas oficiais (opcional) -------------------
    if xsd_dir:
        wml = etree.XMLSchema(etree.parse(os.path.join(xsd_dir, 'wml.xsd')))
        for parte in [p for p in PARTES_WML if p in nomes]:
            arvore = etree.fromstring(zf.read(parte))
            if not wml.validate(arvore):
                for err in list(wml.error_log)[:5]:
                    problemas.append('%s: %s' % (parte, err.message[:180]))
        app = os.path.join(xsd_dir, 'shared-documentPropertiesExtended.xsd')
        if os.path.exists(app) and 'docProps/app.xml' in nomes:
            sch = etree.XMLSchema(etree.parse(app))
            if not sch.validate(etree.fromstring(zf.read('docProps/app.xml'))):
                for err in list(sch.error_log)[:5]:
                    problemas.append('docProps/app.xml: %s' % err.message[:180])
    return problemas


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit('uso: verificar.py <ficheiro.docx> [pasta-dos-xsd]')
    destino = sys.argv[1]
    cache = sys.argv[2] if len(sys.argv) > 2 else \
        os.path.join(os.path.dirname(destino) or '.', 'xsd')
    encontrados = verificar(destino, obter_esquemas(cache))
    for p in encontrados:
        print('  - %s' % p)
    print('%d problema(s) em %s' % (len(encontrados), sys.argv[1]))
    sys.exit(1 if encontrados else 0)
