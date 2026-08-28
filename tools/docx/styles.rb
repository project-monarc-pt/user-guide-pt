# frozen_string_literal: true
#
# word/styles.xml e word/numbering.xml
#
# Todos os estilos usados pelo documento são estilos reais do Word: quem editar
# o ficheiro no futuro altera o aspecto num único sítio (Estilos > Modificar) e
# a mudança propaga-se a todo o guia.

require_relative 'ooxml'

module Ooxml
  module Styles
    O = Ooxml

    # numId reservados
    NUM_HEADINGS = 1
    NUM_BULLET   = 2
    NUM_DYNAMIC_BASE = 10   # listas do corpo: um numId por lista

    # ---------------------------------------------------------------- helpers
    def self.rpr(bold: false, italic: false, color: O::DARK, pt: O::BASE_PT,
                 font: O::FONT, style: nil, valign: nil, underline: nil,
                 caps: false)
      x = +'<w:rPr>'
      x << %(<w:rStyle w:val="#{style}"/>) if style
      x << %(<w:rFonts w:ascii="#{font}" w:hAnsi="#{font}" w:cs="#{font}"/>) if font
      x << '<w:b/><w:bCs/>' if bold
      x << '<w:i/><w:iCs/>' if italic
      x << '<w:caps/>' if caps
      x << %(<w:color w:val="#{color}"/>) if color
      x << %(<w:sz w:val="#{O.pt2half(pt)}"/><w:szCs w:val="#{O.pt2half(pt)}"/>) if pt
      x << %(<w:u w:val="#{underline}"/>) if underline
      x << %(<w:vertAlign w:val="#{valign}"/>) if valign
      x << %(<w:lang w:val="#{O::LANG}"/>)
      x << '</w:rPr>'
    end

    # Ordem dos filhos de w:pPr é imposta pelo schema (ECMA-376).
    def self.ppr(style: nil, keep_next: false, keep_lines: false,
                 page_break: false, num: nil, ilvl: 0, shading: nil,
                 borders: nil, tabs: nil, before: nil, after: nil,
                 line: nil, line_rule: 'auto', ind: nil, contextual: false,
                 jc: nil, outline: nil, rpr: nil)
      x = +'<w:pPr>'
      x << %(<w:pStyle w:val="#{style}"/>) if style
      x << '<w:keepNext/>' if keep_next
      x << '<w:keepLines/>' if keep_lines
      x << '<w:pageBreakBefore/>' if page_break
      if num
        x << %(<w:numPr><w:ilvl w:val="#{ilvl}"/><w:numId w:val="#{num}"/></w:numPr>)
      end
      x << borders if borders
      x << %(<w:shd w:val="clear" w:color="auto" w:fill="#{shading}"/>) if shading
      x << tabs if tabs
      if before || after || line
        x = x.dup
        s = +'<w:spacing'
        s << %( w:before="#{before}") if before
        s << %( w:after="#{after}") if after
        s << %( w:line="#{line}" w:lineRule="#{line_rule}") if line
        s << '/>'
        x << s
      end
      x << ind if ind
      x << '<w:contextualSpacing/>' if contextual
      x << %(<w:jc w:val="#{jc}"/>) if jc
      x << %(<w:outlineLvl w:val="#{outline}"/>) unless outline.nil?
      x << rpr if rpr
      x << '</w:pPr>'
    end

    def self.ind(left: nil, right: nil, hanging: nil, first_line: nil)
      a = +'<w:ind'
      a << %( w:left="#{left}") if left
      a << %( w:right="#{right}") if right
      a << %( w:hanging="#{hanging}") if hanging
      a << %( w:firstLine="#{first_line}") if first_line
      a << '/>'
      a
    end

    def self.style(id, name, type: 'paragraph', based: 'Normal', nxt: nil,
                   priority: 9, qformat: true, custom: false, ppr: nil, rpr: nil,
                   extra: nil)
      x = +%(<w:style w:type="#{type}" w:styleId="#{id}")
      x << ' w:customStyle="1"' if custom
      x << '>'
      x << %(<w:name w:val="#{name}"/>)
      x << %(<w:basedOn w:val="#{based}"/>) if based
      x << %(<w:next w:val="#{nxt}"/>) if nxt
      x << %(<w:uiPriority w:val="#{priority}"/>)
      x << '<w:qFormat/>' if qformat
      x << ppr if ppr
      x << rpr if rpr
      x << extra if extra
      x << '</w:style>'
    end

    def self.borders(color:, sz:, sides: %w[top left bottom right], inside_v: nil,
                     inside_h: nil, tag: 'w:tblBorders')
      x = +"<#{tag}>"
      %w[top left bottom right].each do |s|
        x << if sides.include?(s)
               %(<w:#{s} w:val="single" w:sz="#{sz}" w:space="0" w:color="#{color}"/>)
             else
               %(<w:#{s} w:val="nil"/>)
             end
      end
      if tag == 'w:tblBorders'
        x << (inside_h ? %(<w:insideH w:val="single" w:sz="#{inside_h[1]}" w:space="0" w:color="#{inside_h[0]}"/>) : '<w:insideH w:val="nil"/>')
        x << (inside_v ? %(<w:insideV w:val="single" w:sz="#{inside_v[1]}" w:space="0" w:color="#{inside_v[0]}"/>) : '<w:insideV w:val="nil"/>')
      end
      x << "</#{tag}>"
    end

    # -------------------------------------------------------------- styles.xml
    def self.styles_xml
      body_line = ((O::LINE_PT / O::BASE_PT) * 240).round   # 274 → entrelinha 12 pt
      leader_tab = %(<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="#{O.mm2twip(O::CONTENT_W_MM)}"/></w:tabs>)

      s = +''
      # -- Normal: corpo do guia (Arial 10,5 justificado, Preto Cyber) --------
      s << style('Normal', 'Normal', based: nil, priority: 1,
                 ppr: ppr(after: O.pt2twip(O::LINE_PT), line: body_line, jc: 'both'),
                 rpr: rpr)
      s << style('DefaultParagraphFont', 'Default Paragraph Font',
                 type: 'character', based: nil, priority: 1, qformat: false)
      s << style('TableNormal', 'Normal Table', type: 'table', based: nil,
                 priority: 99, qformat: false,
                 extra: '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/>' \
                        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/>' \
                        '<w:left w:w="108" w:type="dxa"/>' \
                        '<w:bottom w:w="0" w:type="dxa"/>' \
                        '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>')
      s << style('NoList', 'No List', type: 'numbering', based: nil,
                 priority: 99, qformat: false)

      # -- Títulos: teal, negrito, numeração automática ----------------------
      (1..5).each do |lvl|
        pt = O::H_PT[lvl]
        s << style("Heading#{lvl}", "heading #{lvl}", nxt: 'Normal', priority: 9,
                   ppr: ppr(keep_next: true, keep_lines: true,
                            page_break: lvl == 1,
                            num: NUM_HEADINGS, ilvl: lvl - 1,
                            before: lvl == 1 ? 0 : O.pt2twip(O::LINE_PT * 1.2),
                            after: O.pt2twip(O::LINE_PT * 0.6),
                            line: 264, jc: 'left', outline: lvl - 1),
                   rpr: rpr(bold: true, color: O::TEAL, pt: pt))
      end

      # Título de frontmatter (Índice / Lista de Figuras): aspecto de Título 1
      # mas sem número e fora do índice.
      s << style('TituloFrontmatter', 'Título Frontmatter', nxt: 'Normal',
                 custom: true, priority: 10,
                 ppr: ppr(keep_next: true, keep_lines: true, page_break: true,
                          after: O.pt2twip(O::LINE_PT * 0.6), line: 264, jc: 'left'),
                 rpr: rpr(bold: true, color: O::TEAL, pt: O::H_PT[1]))

      # -- Corpo: variantes --------------------------------------------------
      s << style('ListParagraph', 'List Paragraph', nxt: 'ListParagraph',
                 priority: 34,
                 ppr: ppr(after: O.pt2twip(O::LINE_PT * 0.5), line: body_line,
                          jc: 'both'))
      s << style('ImagemMonarc', 'Imagem', nxt: 'Caption', custom: true, priority: 11,
                 ppr: ppr(keep_next: true, before: O.pt2twip(O::LINE_PT * 0.5),
                          after: O.pt2twip(4), line: 240, jc: 'center'))
      s << style('Caption', 'caption', nxt: 'Normal', priority: 35,
                 ppr: ppr(keep_lines: true, before: O.pt2twip(2),
                          after: O.pt2twip(O::LINE_PT), line: 240, jc: 'left'),
                 rpr: rpr(italic: true, pt: O::CAPTION_PT))
      s << style('CaixaMonarc', 'Caixa de destaque', nxt: 'CaixaMonarc',
                 custom: true, priority: 12,
                 ppr: ppr(after: O.pt2twip(O::LINE_PT * 0.4), line: body_line,
                          contextual: true, jc: 'both'))
      s << style('TabelaTexto', 'Texto de tabela', nxt: 'TabelaTexto',
                 custom: true, priority: 13,
                 ppr: ppr(after: 0, line: body_line, jc: 'left'))

      # -- Estilos de carácter ----------------------------------------------
      s << style('Hyperlink', 'Hyperlink', type: 'character',
                 based: 'DefaultParagraphFont', priority: 99,
                 rpr: rpr(color: O::TEAL, pt: nil))
      s << style('CodigoInline', 'Código inline', type: 'character',
                 based: 'DefaultParagraphFont', custom: true, priority: 20,
                 rpr: rpr(color: O::SLATE, font: O::FONT_MONO, pt: O::BASE_PT * 0.95))

      # -- Cabeçalho / rodapé ------------------------------------------------
      s << style('Header', 'header', nxt: 'Header', priority: 99, qformat: false,
                 ppr: ppr(after: 0, line: 240, jc: 'left'),
                 rpr: rpr(color: O::TEAL, pt: O::SMALL_PT))
      s << style('Footer', 'footer', nxt: 'Footer', priority: 99, qformat: false,
                 ppr: ppr(after: 0, line: 240, jc: 'center'),
                 rpr: rpr(color: O::TEAL, pt: O::SMALL_PT))

      # -- Índice e lista de figuras (preenchidos por campos do Word) --------
      s << style('TOCHeading', 'TOC Heading', based: 'TituloFrontmatter',
                 nxt: 'Normal', priority: 39, qformat: false)
      (1..3).each do |lvl|
        s << style("TOC#{lvl}", "toc #{lvl}", nxt: 'Normal', priority: 39,
                   qformat: false,
                   ppr: ppr(tabs: leader_tab, after: O.pt2twip(8), line: 240,
                            ind: (lvl > 1 ? ind(left: O.pt2twip(12) * (lvl - 1), right: O.pt2twip(24)) : ind(right: O.pt2twip(24))),
                            jc: 'left'))
      end
      s << style('TableofFigures', 'table of figures', nxt: 'Normal',
                 priority: 99, qformat: false,
                 ppr: ppr(tabs: leader_tab, after: O.pt2twip(6), line: 240,
                          ind: ind(right: O.pt2twip(24)), jc: 'left'))

      # -- Notas de rodapé ---------------------------------------------------
      s << style('FootnoteText', 'footnote text', nxt: 'FootnoteText',
                 priority: 99, qformat: false,
                 ppr: ppr(after: 0, line: 240, jc: 'left'),
                 rpr: rpr(pt: O::SMALL_PT))
      s << style('FootnoteReference', 'footnote reference', type: 'character',
                 based: 'DefaultParagraphFont', priority: 99, qformat: false,
                 rpr: rpr(valign: 'superscript', pt: nil, color: nil))

      # -- Capa / contracapa -------------------------------------------------
      s << style('CapaTitulo', 'Capa título', nxt: 'CapaSubtitulo', custom: true,
                 priority: 14,
                 ppr: ppr(after: O.pt2twip(6), line: 240, jc: 'left'),
                 rpr: rpr(bold: true, color: 'FFFFFF', pt: O::COVER_TITLE_PT))
      s << style('CapaSubtitulo', 'Capa subtítulo', nxt: 'CapaTexto',
                 custom: true, priority: 15,
                 ppr: ppr(after: O.pt2twip(6), line: 240, jc: 'left'),
                 rpr: rpr(color: 'FFFFFF', pt: O::BASE_PT * 1.25))
      s << style('CapaTexto', 'Capa texto', nxt: 'CapaTexto', custom: true,
                 priority: 16,
                 ppr: ppr(after: O.pt2twip(3), line: 240, jc: 'center'),
                 rpr: rpr(color: 'FFFFFF', pt: O::BASE_PT))

      # -- Tabelas -----------------------------------------------------------
      cell_mar = format(
        '<w:tblCellMar><w:top w:w="%d" w:type="dxa"/><w:left w:w="%d" w:type="dxa"/>' \
        '<w:bottom w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tblCellMar>',
        O.pt2twip(3), O.pt2twip(4), O.pt2twip(3), O.pt2twip(4)
      )
      s << style('TabelaMonarc', 'Tabela MONARC', type: 'table',
                 based: 'TableNormal', custom: true, priority: 59,
                 extra: '<w:tblPr>' +
                        borders(color: O::BORDER, sz: O.pt2eighth(0.5),
                                inside_v: [O::BORDER, O.pt2eighth(0.5)],
                                inside_h: [O::BORDER, O.pt2eighth(0.5)]) +
                        cell_mar + '</w:tblPr>' +
                        '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>')
      s << style('TabelaSemBordas', 'Tabela sem bordas', type: 'table',
                 based: 'TableNormal', custom: true, priority: 60,
                 extra: '<w:tblPr>' +
                        borders(color: O::BORDER, sz: 0, sides: []) +
                        cell_mar + '</w:tblPr>')

      doc_defaults =
        '<w:docDefaults><w:rPrDefault>' +
        rpr.sub('<w:rPr>', '<w:rPr>').sub('</w:rPr>', '</w:rPr>') +
        '</w:rPrDefault><w:pPrDefault>' +
        ppr(after: O.pt2twip(O::LINE_PT), line: body_line, jc: 'both') +
        '</w:pPrDefault></w:docDefaults>'

      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">#{doc_defaults}#{s}</w:styles>
      XML
    end

    # ------------------------------------------------------------ numbering.xml
    def self.lvl(ilvl, fmt:, text:, pstyle: nil, left:, hanging:, suff: nil,
                 start: 1, rpr: nil, jc: 'left')
      x = +%(<w:lvl w:ilvl="#{ilvl}">)
      x << %(<w:start w:val="#{start}"/>)
      x << %(<w:numFmt w:val="#{fmt}"/>)
      x << %(<w:pStyle w:val="#{pstyle}"/>) if pstyle
      x << %(<w:suff w:val="#{suff}"/>) if suff
      x << %(<w:lvlText w:val="#{text}"/>)
      x << %(<w:lvlJc w:val="#{jc}"/>)
      x << '<w:pPr>' << ind(left: left, hanging: hanging) << '</w:pPr>'
      x << rpr if rpr
      x << '</w:lvl>'
    end

    # Recuos das listas do corpo, medidos no PDF de referência.
    LIST_IND = [[482, 284], [1032, 250], [1580, 250], [2128, 250],
                [2676, 250], [3224, 250], [3772, 250], [4320, 250], [4868, 250]].freeze

    def self.heading_abstract(signifier = 'Capítulo', numlevels = 3)
      x = +'<w:abstractNum w:abstractNumId="0"><w:nsid w:val="0A1B2C31"/>' \
          '<w:multiLevelType w:val="multilevel"/><w:name w:val="NumeracaoTitulos"/>'
      # Níveis numerados até :sectnumlevels: (3 por omissão); os restantes sem
      # número. O separador é um espaço inseparável dentro do próprio lvlText,
      # com suff="nothing": assim nenhum leitor o pode substituir pela
      # tabulação que o OOXML assume por omissão (o Google Docs fazia-o, com um
      # espaço grande e diferente de nível para nível).
      sig = signifier.to_s.empty? ? '' : "#{signifier} "
      (0..4).each do |i|
        numbered = i < numlevels
        text = if !numbered then ''
               elsif i.zero? then "#{sig}%1.&#160;"
               else (1..(i + 1)).map { |k| "%#{k}." }.join + '&#160;'
               end
        x << lvl(i, fmt: numbered ? 'decimal' : 'none', text: text,
                 pstyle: "Heading#{i + 1}", left: 0, hanging: 0, suff: 'nothing')
      end
      (5..8).each do |i|
        x << lvl(i, fmt: 'none', text: '', left: 0, hanging: 0, suff: 'nothing')
      end
      x << '</w:abstractNum>'
    end

    def self.bullet_abstract
      sym = '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
      x = +'<w:abstractNum w:abstractNumId="1"><w:nsid w:val="0A1B2C32"/>' \
          '<w:multiLevelType w:val="hybridMultilevel"/><w:name w:val="Marcas"/>'
      chars = ["\uF0B7", "\uF0A7", "\uF0A8"] * 3
      9.times do |i|
        l, h = LIST_IND[i]
        x << lvl(i, fmt: 'bullet', text: chars[i], left: l, hanging: h, rpr: sym)
      end
      x << '</w:abstractNum>'
    end

    # Cada lista do corpo recebe um numId próprio para respeitar [start=N]
    # sem deixar de ser uma lista automática do Word.
    def self.list_abstract(id, formats)
      x = +%(<w:abstractNum w:abstractNumId="#{id}"><w:nsid w:val="#{format('%08X', 0x10000000 + id)}"/>)
      x << '<w:multiLevelType w:val="hybridMultilevel"/>'
      9.times do |i|
        l, h = LIST_IND[i]
        fmt = formats[i] || (i.zero? ? 'decimal' : %w[lowerLetter lowerRoman][(i - 1) % 2])
        text = "%#{i + 1}."
        x << lvl(i, fmt: fmt, text: text, left: l, hanging: h)
      end
      x << '</w:abstractNum>'
    end

    def self.numbering_xml(dynamic, signifier = 'Capítulo', numlevels = 3)
      abstracts = +''
      abstracts << heading_abstract(signifier, numlevels) << bullet_abstract
      nums = +%(<w:num w:numId="#{NUM_HEADINGS}"><w:abstractNumId w:val="0"/></w:num>)
      nums << %(<w:num w:numId="#{NUM_BULLET}"><w:abstractNumId w:val="1"/></w:num>)
      dynamic.each do |d|
        abstracts << list_abstract(d[:abstract_id], d[:formats])
        nums << %(<w:num w:numId="#{d[:num_id]}"><w:abstractNumId w:val="#{d[:abstract_id]}"/>)
        # startOverride reproduz o [start=N] do AsciiDoc mantendo a lista viva.
        nums << %(<w:lvlOverride w:ilvl="#{d[:ilvl]}"><w:startOverride w:val="#{d[:start]}"/></w:lvlOverride>)
        nums << '</w:num>'
      end
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">#{abstracts}#{nums}</w:numbering>
      XML
    end
  end
end
