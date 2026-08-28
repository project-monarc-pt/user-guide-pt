# frozen_string_literal: true
#
# Conversor Asciidoctor → WordprocessingML.
#
# O texto em linha é primeiro codificado com marcas privadas (negrito, itálico,
# monoespaçado, XML em bruto) e só depois transformado em <w:r>, o que permite
# aninhar formatação corretamente. Tudo o que o Word sabe automatizar fica em
# campos: TOC (índice), TOC \c (lista de figuras), SEQ (numeração de figuras),
# REF/NOTEREF (remissões) e PAGE (rodapé).

require 'asciidoctor'
require 'asciidoctor/converter'
require_relative 'ooxml'
require_relative 'styles'
require_relative 'package'

class DocxConverter < Asciidoctor::Converter::Base
  O = Ooxml
  S = Ooxml::Styles

  # Marcas de formatação (área de uso privado do Unicode).
  B_ON    = "\uE001"; B_OFF   = "\uE002"
  I_ON    = "\uE003"; I_OFF   = "\uE004"
  M_ON    = "\uE005"; M_OFF   = "\uE006"
  SUP_ON  = "\uE007"; SUP_OFF = "\uE008"
  SUB_ON  = "\uE009"; SUB_OFF = "\uE00A"
  RAW_A   = "\uE010"; RAW_B   = "\uE011"
  BREAK   = "\uE012"
  TOGGLES = {
    B_ON => [:bold, 1], B_OFF => [:bold, -1],
    I_ON => [:italic, 1], I_OFF => [:italic, -1],
    M_ON => [:mono, 1], M_OFF => [:mono, -1],
    SUP_ON => [:sup, 1], SUP_OFF => [:sup, -1],
    SUB_ON => [:sub, 1], SUB_OFF => [:sub, -1]
  }.freeze
  MARK_CHARS = (TOGGLES.keys + [BREAK]).join.freeze

  # Marcas substituídas depois da conversão pelas entradas dos dois índices
  # (só nesse momento se conhecem todos os títulos e figuras).
  TOC_MARK = '%%ENTRADAS-INDICE%%'
  TOF_MARK = '%%ENTRADAS-FIGURAS%%'

  attr_reader :pkg, :toc_entries, :fig_entries

  class << self
    attr_accessor :context
  end

  def initialize(backend, opts = {})
    super
    ctx = self.class.context || {}
    @pkg = ctx[:pkg]
    @base_dir = ctx[:base_dir]
    init_backend_traits basebackend: 'html', filetype: 'docx',
                        outfilesuffix: '.docx', supports_templates: false
    @raws = []
    @indent = 0                 # recuo herdado, em twips
    @list_depth = 0
    @para_style = 'Normal'
    @pending_break = false
    @footnote_marks = {}
    @labels = ctx[:labels] || { chapter: 'Capítulo', figures: 'Lista de Figuras',
                                figure: 'Figura', sectnumlevels: 3 }
    @toc_entries = []           # títulos de nível 1–2 (índice)
    @fig_entries = []           # legendas de figura (lista de figuras)
    @sec_counters = [0, 0, 0]
    @fig_counter = 0
  end

  # Despacho explícito: mais previsível do que o method_missing da classe base.
  def convert(node, transform = node.node_name, _opts = nil)
    meth = "convert_#{transform}"
    return send(meth, node) if respond_to?(meth)
    @pkg.warn! "sem conversor para '#{transform}'"
    ''
  end

  # ====================================================================== blocos
  # As marcas do índice e da lista de figuras são emitidas aqui — e não no
  # preâmbulo — porque há guias sem qualquer texto de preâmbulo. Só no documento
  # de topo: as células de tabela em AsciiDoc são documentos aninhados e também
  # passam por aqui.
  def convert_document(node)
    front = node.nested? ? '' : TOC_MARK + TOF_MARK
    front + node.blocks.map { |b| b.convert }.join
  end
  alias convert_embedded convert_document

  # Do preâmbulo só se descarta a lista de figuras escrita à mão (que passa a ser
  # um campo do Word); todo o restante conteúdo é convertido.
  def convert_preamble(node)
    node.blocks.reject { |b| lista_de_figuras_manual?(b) }.map { |b| b.convert }.join
  end

  # Reconhece o título "Lista de Figuras" e a lista de remissões que o segue.
  def lista_de_figuras_manual?(block)
    case block.context
    when :floating_title
      block.id.to_s == 'lista-de-figuras' ||
        strip_marks(block.title).downcase.start_with?('lista de figuras')
    when :ulist
      !block.items.empty? && block.items.all? { |i|
        t = i.text.to_s
        t.include?(RAW_A) && t.gsub(/#{RAW_A}\d+#{RAW_B}/, '').strip.empty?
      }
    else
      false
    end
  end

  def convert_section(node)
    lvl = node.level.clamp(1, 5)
    numbered = node.numbered && node.sectname != 'preface' &&
               lvl <= @labels[:sectnumlevels]
    register_toc_entry(node, lvl, numbered)
    heading_para(node.title, style: "Heading#{lvl}", id: node.id, numbered: numbered) +
      node.blocks.map { |b| b.convert }.join
  end

  # Reproduz a numeração da lista multinível para o texto das entradas do
  # índice; só os níveis 1 e 2 entram (o campo é TOC \o "1-2").
  def register_toc_entry(node, lvl, numbered)
    prefix = ''
    if numbered && lvl <= @labels[:sectnumlevels]
      @sec_counters[lvl - 1] += 1
      (lvl..2).each { |i| @sec_counters[i] = 0 }
      nums = @sec_counters[0, lvl].join('.')
      nbsp = "\u00A0"                       # igual ao separador da numeração
      sig = @labels[:chapter].to_s.empty? ? '' : "#{@labels[:chapter]} "
      prefix = lvl == 1 ? "#{sig}#{nums}.#{nbsp}" : "#{nums}.#{nbsp}"
    end
    return unless lvl <= 2 && node.id
    @toc_entries << { level: lvl, text: prefix + strip_marks(node.title),
                      bookmark: @pkg.bookmark(node.id) }
  end

  def convert_floating_title(node)
    heading_para(node.title, style: "Heading#{node.level.clamp(1, 5)}",
                 id: node.id, numbered: false)
  end

  def convert_paragraph(node)
    paragraph(build_runs(node.content), style: @para_style,
              jc: align_from_roles(node), ind: indent_xml)
  end

  def convert_open(node)
    node.blocks.empty? ? convert_paragraph(node) : node.blocks.map { |b| b.convert }.join
  end
  alias convert_example convert_open

  def convert_page_break(_node)
    @pending_break = true
    ''
  end

  def convert_thematic_break(_node)
    paragraph('', borders: '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="4" w:color="EEEEEE"/></w:pBdr>')
  end

  def convert_listing(node) = convert_literal(node)

  def convert_literal(node)
    runs = node.content.to_s.split("\n").map.with_index { |ln, i|
      (i.zero? ? '' : '<w:r><w:br/></w:r>') +
        %(<w:r><w:rPr><w:rStyle w:val="CodigoInline"/></w:rPr><w:t xml:space="preserve">#{O.text(ln)}</w:t></w:r>)
    }.join
    paragraph(runs, jc: 'left', shading: 'F5F5F5', ind: indent_xml,
              borders: '<w:pBdr>' \
                       '<w:top w:val="single" w:sz="6" w:space="4" w:color="CCCCCC"/>' \
                       '<w:left w:val="single" w:sz="6" w:space="4" w:color="CCCCCC"/>' \
                       '<w:bottom w:val="single" w:sz="6" w:space="4" w:color="CCCCCC"/>' \
                       '<w:right w:val="single" w:sz="6" w:space="4" w:color="CCCCCC"/></w:pBdr>')
  end

  def convert_quote(node)
    saved = @indent
    @indent += O.pt2twip(12)
    inner = node.blocks.empty? ? paragraph(build_runs(node.content), ind: indent_xml)
                               : node.blocks.map { |b| b.convert }.join
    @indent = saved
    inner
  end
  alias convert_verse convert_quote
  alias convert_sidebar convert_open

  # ====================================================================== listas
  def convert_olist(node)
    ilvl = @list_depth.clamp(0, 8)
    fmt = case node.style
          when 'loweralpha' then 'lowerLetter'
          when 'upperalpha' then 'upperLetter'
          when 'lowerroman' then 'lowerRoman'
          when 'upperroman' then 'upperRoman'
          else 'decimal'
          end
    num = @pkg.new_list_num(ilvl: ilvl, start: ((node.attr 'start') || 1).to_i, format: fmt)
    list_items(node, num, ilvl)
  end

  def convert_ulist(node)
    list_items(node, S::NUM_BULLET, @list_depth.clamp(0, 8))
  end
  alias convert_colist convert_olist

  def list_items(node, num, ilvl)
    out = +''
    node.items.each do |item|
      out << paragraph(build_runs(item.text), style: 'ListParagraph', num: num, ilvl: ilvl)
      next if item.blocks.empty?
      saved_indent = @indent
      @indent = S::LIST_IND[ilvl][0]
      @list_depth += 1
      out << item.blocks.map { |b| b.convert }.join
      @list_depth -= 1
      @indent = saved_indent
    end
    out
  end

  def convert_dlist(node)
    out = +''
    node.items.each do |terms, desc|
      Array(terms).each do |t|
        out << paragraph(build_runs(B_ON + t.text.to_s + B_OFF), ind: indent_xml,
                         after: O.pt2twip(3))
      end
      next unless desc
      saved = @indent
      @indent += O.pt2twip(15)
      out << paragraph(build_runs(desc.text), ind: indent_xml) if desc.text?
      out << desc.blocks.map { |b| b.convert }.join unless desc.blocks.empty?
      @indent = saved
    end
    out
  end

  # ===================================================================== imagens
  def convert_image(node)
    path = resolve_image(node, node.attr('target'))
    w, h = image_dims(node, path, available_mm)
    return '' unless w
    jc = case node.attr('align').to_s
         when 'center' then 'center'
         when 'right' then 'right'
         else 'left'
         end
    paragraph(%(<w:r>#{@pkg.drawing(path, w, h, alt: node.attr('alt').to_s)}</w:r>),
              style: 'ImagemMonarc', jc: jc, ind: indent_xml) +
      (node.title? ? figure_caption(node) : '')
  end

  # Legenda com campo SEQ: a numeração das figuras passa a ser mantida pelo
  # Word, mas o valor em cache é o número real (o documento lê-se bem mesmo
  # antes de qualquer atualização de campos).
  def figure_caption(node)
    label = @labels[:figure]
    prefix, number, suffix = node.caption.to_s.match(/\A(.*?)(\d+)(.*)\z/)&.captures ||
                             ["#{label} ", '1', '. ']
    inner = run_text(prefix) + field("SEQ #{label} \\* ARABIC", run_text(number)) +
            run_text(suffix) + build_runs(node.title)
    @fig_counter += 1
    id = node.id || "fig-auto-#{@fig_counter}"
    mark = @pkg.bookmark(id)
    @fig_entries << { text: "#{prefix}#{number}#{suffix}#{strip_marks(node.title)}",
                      bookmark: mark }
    paragraph(bookmarked_name(mark, inner), style: 'Caption', ind: indent_xml)
  end

  # Resolve o alvo respeitando o imagesdir do documento, com recurso a
  # images/<ficheiro> quando o caminho relativo não existe.
  def resolve_image(node, target)
    t = target.to_s
    return t if t.start_with?('/')
    uri = node.respond_to?(:image_uri) ? node.image_uri(t) : File.join('images', t)
    path = File.expand_path(File.join(@base_dir, uri))
    return path if File.file?(path)
    alt = File.expand_path(File.join(@base_dir, 'images', File.basename(t)))
    File.file?(alt) ? alt : path
  end

  # Reproduz o dimensionamento do asciidoctor-pdf: pdfwidth em % da largura
  # útil; senão o atributo width (px tratados como pt); senão o tamanho
  # intrínseco, reduzido para caber na página.
  def image_dims(node, path, avail_mm, max_h_mm = 215.0)
    px = O.png_size(path)
    unless px
      @pkg.warn! "dimensões ilegíveis: #{path}"
      return [nil, nil]
    end
    pw, ph = px
    spec = (node.attr('pdfwidth') || node.attr('width')).to_s
    w = case spec
        when /\A([\d.]+)%\z/  then avail_mm * Regexp.last_match(1).to_f / 100.0
        when /\A([\d.]+)mm\z/ then Regexp.last_match(1).to_f
        when /\A([\d.]+)cm\z/ then Regexp.last_match(1).to_f * 10
        when /\A([\d.]+)in\z/ then Regexp.last_match(1).to_f * 25.4
        when /\A([\d.]+)(pt|px)?\z/ then Regexp.last_match(1).to_f * O::MM_PER_PT
        else pw * O::MM_PER_PT
        end
    w = avail_mm if w > avail_mm
    h = w * ph / pw.to_f
    if h > max_h_mm
      w *= max_h_mm / h
      h = max_h_mm
    end
    [w.round(2), h.round(2)]
  end

  # ======================================================== caixas de destaque
  def convert_admonition(node)
    conf = O::ADMON[(node.attr 'name').to_s] || O::ADMON['note']
    icon = File.join(@pkg.assets_dir, conf[:icon])
    total = O.mm2twip(O::CONTENT_W_MM) - @indent
    c1 = O.mm2twip(21.3)
    c2 = total - c1

    saved_style = @para_style
    saved_indent = @indent
    @para_style = 'CaixaMonarc'
    @indent = 0
    content = if node.blocks.empty?
                paragraph(build_runs(node.content), style: 'CaixaMonarc')
              else
                node.blocks.map { |b| b.convert }.join
              end
    @para_style = saved_style
    @indent = saved_indent

    icon_p = paragraph(%(<w:r>#{@pkg.drawing(icon, 7.6, 7.6, alt: (node.attr 'name').to_s.upcase)}</w:r>),
                       style: 'CaixaMonarc', jc: 'center', after: 0)
    cells = cell(c1, icon_p, shading: conf[:bg], valign: 'center',
                 borders: S.borders(color: conf[:border], sz: O.pt2eighth(3),
                                    sides: %w[right], tag: 'w:tcBorders')) +
            cell(c2, content, shading: conf[:bg], valign: 'center')
    table(row(cells), [c1, c2], style: 'TabelaMonarc',
          borders: S.borders(color: conf[:border], sz: O.pt2eighth(0.5)),
          cell_mar: [O.pt2twip(6), O.pt2twip(12)])
  end

  # ==================================================================== tabelas
  def convert_table(node)
    frame = (node.attr 'frame') || 'all'
    grid  = (node.attr 'grid') || 'all'
    borderless = frame == 'none' && grid == 'none'
    pct = ((node.attr 'tablepcwidth') || 100).to_f
    total = ((O.mm2twip(O::CONTENT_W_MM) - @indent) * pct / 100).round
    cols = node.columns
    widths = cols.map { |c| (total * ((c.attr 'colpcwidth') || (100.0 / cols.size)).to_f / 100).round }
    widths[-1] = total - widths[0..-2].sum if widths.size > 1

    caption = node.title? ? paragraph(build_runs(node.title), style: 'Caption',
                                      keep_next: true, ind: indent_xml) : ''
    has_head = !node.rows.head.empty?
    rows = +''
    [[node.rows.head, true], [node.rows.body, false], [node.rows.foot, false]].each do |set, head|
      set.each do |r|
        cells = +''
        col = 0
        r.each do |c|
          span = c.colspan.to_i
          span = 1 if span < 1
          w = (col...(col + span)).sum { |k| widths[k] || 0 }
          cells << cell(w, cell_content(c, head), valign: valign_of(c),
                        span: span > 1 ? span : nil)
          col += span
        end
        rows << row(cells, header: head && has_head)
      end
    end
    caption + table(rows, widths, style: borderless ? 'TabelaSemBordas' : 'TabelaMonarc',
                    borders: borderless ? nil : table_borders(frame, grid),
                    first_row: has_head)
  end

  def table_borders(frame, grid)
    S.borders(color: O::BORDER, sz: O.pt2eighth(0.5),
              sides: frame == 'none' ? [] : %w[top left bottom right],
              inside_h: %w[all rows].include?(grid) ? [O::BORDER, O.pt2eighth(0.5)] : nil,
              inside_v: %w[all cols].include?(grid) ? [O::BORDER, O.pt2eighth(0.5)] : nil)
  end

  def valign_of(cell)
    case (cell.attr 'valign').to_s
    when 'middle' then 'center'
    when 'bottom' then 'bottom'
    else 'top'
    end
  end

  def cell_content(c, head)
    jc = case (c.attr 'halign').to_s
         when 'center' then 'center'
         when 'right' then 'right'
         else 'left'
         end
    if !head && c.style == :asciidoc
      saved_style = @para_style
      saved_indent = @indent
      @para_style = 'TabelaTexto'
      @indent = 0
      out = c.content.to_s
      @para_style = saved_style
      @indent = saved_indent
      return out.strip.empty? ? paragraph('', style: 'TabelaTexto') : out
    end
    texts = head ? [c.text] : Array(c.content)
    body = texts.map { |t|
      paragraph(build_runs(t), style: 'TabelaTexto', jc: jc,
                rpr: head ? '<w:rPr><w:b/><w:bCs/></w:rPr>' : nil)
    }.join
    body.empty? ? paragraph('', style: 'TabelaTexto') : body
  end

  # ============================================================ texto em linha
  def convert_inline_quoted(node)
    case node.type
    when :strong      then B_ON + node.text + B_OFF
    when :emphasis    then I_ON + node.text + I_OFF
    when :monospaced  then M_ON + node.text + M_OFF
    when :superscript then SUP_ON + node.text + SUP_OFF
    when :subscript   then SUB_ON + node.text + SUB_OFF
    when :double      then %(“#{node.text}”)
    when :single      then %(‘#{node.text}’)
    else node.text
    end
  end

  def convert_inline_break(node) = node.text.to_s + BREAK
  def convert_inline_indexterm(_node) = ''
  def convert_inline_callout(node) = "(#{node.text})"
  def convert_inline_button(node) = B_ON + node.text.to_s + B_OFF

  def convert_inline_kbd(node)
    keys = node.attr('keys')
    M_ON + (keys.is_a?(Array) ? keys.join('+') : keys.to_s) + M_OFF
  end

  def convert_inline_menu(node)
    parts = [node.attr('menu')] + Array(node.attr('submenus')) + [node.attr('menuitem')]
    B_ON + parts.compact.reject { |p| p.to_s.empty? }.join(' › ') + B_OFF
  end

  def convert_inline_anchor(node)
    case node.type
    when :link
      raw(%(<w:hyperlink r:id="#{@pkg.link_rid(node.target)}" w:history="1">) +
          build_runs(node.text, style: 'Hyperlink') + '</w:hyperlink>')
    when :xref then raw(xref_xml(node))
    when :ref
      raw(bookmarked(node.id || node.target.to_s.delete('#'), ''))
    when :bibref
      raw(run_text("[#{node.reftext || node.id}]"))
    else
      node.text.to_s
    end
  end

  def convert_inline_image(node)
    if node.type == 'icon'
      @pkg.warn! "ícone sem equivalente gráfico: #{node.target}"
      return node.attr('alt').to_s
    end
    path = resolve_image(node, node.target)
    w, h = image_dims(node, path, available_mm, 20.0)
    return node.attr('alt').to_s unless w
    raw(%(<w:r>#{@pkg.drawing(path, w, h, alt: node.attr('alt').to_s)}</w:r>))
  end

  # Notas de rodapé reais; repetições da mesma nota usam NOTEREF, como o Word.
  def convert_inline_footnote(node)
    # O Asciidoctor devolve id na definição (type :ref) e target nas repetições
    # (type :xref); estas passam a ser remissões NOTEREF para a nota original.
    id = node.id
    if node.type == :xref && (mark = @footnote_marks[node.target.to_s])
      index = (node.attr 'index') || 1
      return raw(field("NOTEREF #{mark} \\h",
                       %(<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:t>#{O.text(index)}</w:t></w:r>)))
    end
    body = %(<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>) +
           '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
           '<w:r><w:t xml:space="preserve"> </w:t></w:r>' + build_runs(node.text.to_s) + '</w:p>'
    fid = @pkg.add_footnote(body)
    ref = %(<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="#{fid}"/></w:r>)
    if id
      mark = @pkg.bookmark("fn_#{id}")
      @footnote_marks[id] = mark
      ref = bookmarked_name(mark, ref)
    end
    raw(ref)
  end

  # ================================================================= auxiliares
  private

  def available_mm = O::CONTENT_W_MM - (@indent / O::TWIP_PER_MM)

  def indent_xml = @indent.positive? ? S.ind(left: @indent) : nil

  def align_from_roles(node)
    roles = node.respond_to?(:roles) ? node.roles : []
    return 'center' if roles.include?('text-center')
    return 'right' if roles.include?('text-right')
    return 'left' if roles.include?('text-left')
    nil
  end

  def raw(xml)
    @raws << xml
    "#{RAW_A}#{@raws.length - 1}#{RAW_B}"
  end

  def run_text(str, extra = {})
    %(<w:r>#{run_rpr({}, extra)}<w:t xml:space="preserve">#{O.text(str)}</w:t></w:r>)
  end

  def run_rpr(state, extra)
    parts = +''
    style = extra[:style] || (state[:mono].to_i.positive? ? 'CodigoInline' : nil)
    parts << %(<w:rStyle w:val="#{style}"/>) if style
    parts << '<w:b/><w:bCs/>' if state[:bold].to_i.positive? || extra[:bold]
    parts << '<w:i/><w:iCs/>' if state[:italic].to_i.positive?
    parts << %(<w:color w:val="#{extra[:color]}"/>) if extra[:color]
    parts << %(<w:sz w:val="#{O.pt2half(extra[:pt])}"/>) if extra[:pt]
    parts << '<w:vertAlign w:val="superscript"/>' if state[:sup].to_i.positive?
    parts << '<w:vertAlign w:val="subscript"/>' if state[:sub].to_i.positive?
    parts.empty? ? '' : "<w:rPr>#{parts}</w:rPr>"
  end

  # Texto marcado → sequência de <w:r>.
  def build_runs(text, extra = {})
    s = text.to_s
    out = +''
    buf = +''
    state = { bold: 0, italic: 0, mono: 0, sup: 0, sub: 0 }
    flush = lambda do
      next if buf.empty?
      out << %(<w:r>#{run_rpr(state, extra)}<w:t xml:space="preserve">#{O.text(buf)}</w:t></w:r>)
      buf.clear
    end
    i = 0
    while i < s.length
      c = s[i]
      if c == RAW_A
        j = s.index(RAW_B, i) || s.length
        flush.call
        out << @raws[s[(i + 1)...j].to_i].to_s
        i = j + 1
      elsif (tog = TOGGLES[c])
        flush.call
        state[tog[0]] += tog[1]
        i += 1
      elsif c == BREAK
        flush.call
        out << '<w:r><w:br/></w:r>'
        i += 1
      else
        buf << c
        i += 1
      end
    end
    flush.call
    out
  end

  def strip_marks(str)
    str.to_s.gsub(/#{RAW_A}\d+#{RAW_B}/, '').delete(MARK_CHARS).strip
  end

  # Remissão: campo REF quando o texto coincide com o título de destino (o Word
  # passa a atualizá-lo sozinho); caso contrário, hiperligação interna.
  def xref_xml(node)
    id = (node.attr 'refid') || node.target.to_s.delete_prefix('#')
    mark = @pkg.bookmark(id)
    inner = build_runs(node.text, style: 'Hyperlink')
    target = node.document.catalog[:refs][id]
    title = target.respond_to?(:title?) && target.title? ? target.title : nil
    figure = target.respond_to?(:context) && target.context == :image
    # O Word grava as remissões como campo REF dentro de uma hiperligação
    # interna: mantém-se clicável, com a cor institucional, e o texto é
    # atualizado sozinho quando o título de destino muda.
    body = if figure || (title && strip_marks(title) == strip_marks(node.text))
             field("REF #{mark} \\h", inner)
           else
             inner
           end
    %(<w:hyperlink w:anchor="#{mark}" w:history="1">#{body}</w:hyperlink>)
  end

  def field(instr, cached = '') = O.field(instr, cached)

  def bookmarked(id, inner) = bookmarked_name(@pkg.bookmark(id), inner)

  def bookmarked_name(mark, inner)
    bid = @pkg.next_bookmark_id
    %(<w:bookmarkStart w:id="#{bid}" w:name="#{mark}"/>#{inner}<w:bookmarkEnd w:id="#{bid}"/>)
  end

  def heading_para(title, style:, id: nil, numbered: true)
    inner = build_runs(title)
    inner = bookmarked(id, inner) if id
    paragraph(inner, style: style, num: numbered ? nil : 0, ilvl: 0)
  end

  # Construtor central de parágrafos (consome também o <<< pendente).
  def paragraph(runs, style: nil, jc: nil, num: nil, ilvl: 0, ind: nil,
                after: nil, before: nil, keep_next: false, borders: nil,
                shading: nil, rpr: nil)
    brk = @pending_break
    @pending_break = false
    pr = S.ppr(style: style || @para_style, keep_next: keep_next, page_break: brk,
               num: num, ilvl: ilvl, shading: shading, borders: borders,
               before: before, after: after, ind: ind, jc: jc, rpr: rpr)
    "<w:p>#{pr}#{runs}</w:p>"
  end

  def cell(width, content, shading: nil, valign: nil, borders: nil, span: nil)
    pr = +%(<w:tcPr><w:tcW w:w="#{width}" w:type="dxa"/>)
    pr << %(<w:gridSpan w:val="#{span}"/>) if span
    pr << borders if borders
    pr << %(<w:shd w:val="clear" w:color="auto" w:fill="#{shading}"/>) if shading
    pr << %(<w:vAlign w:val="#{valign}"/>) if valign
    pr << '</w:tcPr>'
    "<w:tc>#{pr}#{content}</w:tc>"
  end

  def row(cells, header: false)
    pr = +'<w:trPr><w:cantSplit/>'
    pr << '<w:tblHeader/>' if header
    pr << '</w:trPr>'
    "<w:tr>#{pr}#{cells}</w:tr>"
  end

  # Tabela + parágrafo de fecho: garante o espaçamento de bloco (12 pt) e evita
  # que duas tabelas consecutivas se fundam numa só.
  def table(rows, widths, style:, borders: nil, first_row: false, cell_mar: nil)
    brk = @pending_break
    @pending_break = false
    pr = +%(<w:tblPr><w:tblStyle w:val="#{style}"/>)
    pr << %(<w:tblW w:w="#{widths.sum}" w:type="dxa"/>)
    pr << %(<w:tblInd w:w="#{@indent}" w:type="dxa"/>) if @indent.positive?
    pr << borders if borders
    pr << '<w:tblLayout w:type="fixed"/>'
    if cell_mar
      v, h = cell_mar
      pr << %(<w:tblCellMar><w:top w:w="#{v}" w:type="dxa"/><w:left w:w="#{h}" w:type="dxa"/>) <<
            %(<w:bottom w:w="#{v}" w:type="dxa"/><w:right w:w="#{h}" w:type="dxa"/></w:tblCellMar>)
    end
    pr << %(<w:tblLook w:val="04A0" w:firstRow="#{first_row ? 1 : 0}" w:lastRow="0") <<
          ' w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
    pr << '</w:tblPr>'
    grid = "<w:tblGrid>#{widths.map { |w| %(<w:gridCol w:w="#{w}"/>) }.join}</w:tblGrid>"
    lead = brk ? %(<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>) : ''
    trail = %(<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="#{O.pt2twip(O::LINE_PT)}" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>)
    "#{lead}<w:tbl>#{pr}#{grid}#{rows}</w:tbl>#{trail}"
  end
end
