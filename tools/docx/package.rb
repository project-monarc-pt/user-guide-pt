# frozen_string_literal: true
#
# Montagem do pacote OOXML: registos de imagens/ligações/notas, secções da
# página A4, cabeçalho/rodapé, capa e contracapa, e escrita do .docx.

require 'fileutils'
require_relative 'ooxml'
require_relative 'styles'

module Ooxml
  NS = {
    'xmlns:w'   => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'xmlns:r'   => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:wp'  => 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'xmlns:a'   => 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xmlns:pic' => 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'xmlns:v'   => 'urn:schemas-microsoft-com:vml',
    'xmlns:m'   => 'http://schemas.openxmlformats.org/officeDocument/2006/math'
  }.freeze

  REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  CT  = 'application/vnd.openxmlformats-officedocument.wordprocessingml'

  class Package
    O = Ooxml
    S = Ooxml::Styles

    attr_reader :warnings

    def initialize(images_dir:, assets_dir:)
      @images_dir = images_dir
      @assets_dir = assets_dir
      @media = {}                                   # caminho absoluto => nome em word/media
      @media_names = {}                             # nome => caminho (deteta colisões)
      @rels = { document: [], header: [] }          # relações por parte
      @rid = { document: 0, header: 0 }
      @image_rels = {}                              # [parte, nome] => rId
      @links = {}                                   # [parte, url] => rId
      @footnotes = []
      @numbering = []
      @bookmark_names = {}
      @taken_names = []
      @docpr = 0
      @bookmark_id = 0
      @next_num_id = S::NUM_DYNAMIC_BASE
      @warnings = []
      @part = :document
    end

    # ------------------------------------------------------------- registos
    attr_accessor :part
    attr_reader :assets_dir, :images_dir

    def next_docpr = (@docpr += 1)
    def next_bookmark_id = (@bookmark_id += 1)

    def warn!(msg)
      @warnings << msg unless @warnings.include?(msg)
    end

    def add_rel(type, target, mode = nil, part: @part)
      rid = "rId#{@rid[part] += 1}"
      @rels[part] << { id: rid, type: "#{REL}/#{type}", target: target, mode: mode }
      rid
    end

    def media_name(abs)
      @media[abs] ||= begin
        name = File.basename(abs)
        if @media_names.key?(name) && @media_names[name] != abs
          ext = File.extname(name)
          name = "#{File.basename(name, ext)}_#{@media.size + 1}#{ext}"
        end
        @media_names[name] = abs
        name
      end
    end

    def image_rid(path)
      abs = File.expand_path(path)
      unless File.file?(abs)
        warn! "imagem inexistente: #{path}"
        return nil
      end
      name = media_name(abs)
      @image_rels[[@part, name]] ||= add_rel('image', "media/#{name}")
    end

    def link_rid(url)
      @links[[@part, url]] ||= add_rel('hyperlink', url, 'External')
    end

    def media_files = @media

    def add_footnote(paragraphs_xml)
      id = @footnotes.length + 1
      @footnotes << %(<w:footnote w:id="#{id}">#{paragraphs_xml}</w:footnote>)
      id
    end

    # Devolve o numId de uma nova lista (mantém [start=N] via startOverride).
    def new_list_num(ilvl:, start:, format:)
      num_id = @next_num_id
      @next_num_id += 1
      @numbering << { num_id: num_id, abstract_id: num_id, ilvl: ilvl,
                      start: start, formats: { ilvl => format } }
      num_id
    end

    def numbering_defs = @numbering

    def bookmark(id)
      @bookmark_names[id.to_s] ||= O.bookmark_name(id, @taken_names)
    end

    def bookmark?(id) = @bookmark_names.key?(id.to_s)

    # -------------------------------------------------------------- imagens
    # Devolve o XML de um <w:drawing> em linha, dimensionado em mm.
    def drawing(path, width_mm, height_mm, alt: '')
      rid = image_rid(path)
      return '' unless rid
      cx = O.mm2emu(width_mm)
      cy = O.mm2emu(height_mm)
      id = next_docpr
      name = File.basename(path)
      <<~XML.gsub(/\n\s*/, '')
        <w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="#{cx}" cy="#{cy}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="#{id}" name="Imagem #{id}" descr="#{O.attr(alt)}"/>
        <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic><pic:nvPicPr><pic:cNvPr id="#{id}" name="#{O.attr(name)}" descr="#{O.attr(alt)}"/>
        <pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="#{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="#{cx}" cy="#{cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
        </a:graphicData></a:graphic></wp:inline></w:drawing>
      XML
    end

    # Imagem flutuante atrás do texto (fundo teal da capa/contracapa).
    def background_drawing(path, width_mm, height_mm)
      rid = image_rid(path)
      return '' unless rid
      cx = O.mm2emu(width_mm)
      cy = O.mm2emu(height_mm)
      id = next_docpr
      <<~XML.gsub(/\n\s*/, '')
        <w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="#{cx}" cy="#{cy}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:wrapNone/>
        <wp:docPr id="#{id}" name="Fundo #{id}" descr="Fundo institucional CNCS"/>
        <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic><pic:nvPicPr><pic:cNvPr id="#{id}" name="fundo.png"/><pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="#{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="#{cx}" cy="#{cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
        </a:graphicData></a:graphic></wp:anchor></w:drawing>
      XML
    end

    # ------------------------------------------------------------- secções
    def sect_pr(header: nil, footer: nil, margins: nil, restart: nil, last: false)
      m = margins || { top: O::MARGIN_TOP_MM, right: O::MARGIN_SIDE_MM,
                       bottom: O::MARGIN_BOTTOM_MM, left: O::MARGIN_SIDE_MM }
      x = +'<w:sectPr>'
      x << %(<w:headerReference r:id="#{header}" w:type="default"/>) if header
      x << %(<w:footerReference r:id="#{footer}" w:type="default"/>) if footer
      x << '<w:type w:val="nextPage"/>'
      x << %(<w:pgSz w:w="#{O.mm2twip(O::PAGE_W_MM)}" w:h="#{O.mm2twip(O::PAGE_H_MM)}"/>)
      x << %(<w:pgMar w:top="#{O.mm2twip(m[:top])}" w:right="#{O.mm2twip(m[:right])}") <<
           %( w:bottom="#{O.mm2twip(m[:bottom])}" w:left="#{O.mm2twip(m[:left])}") <<
           %( w:header="#{O.mm2twip(O::HEADER_DIST_MM)}" w:footer="#{O.mm2twip(O::FOOTER_DIST_MM)}" w:gutter="0"/>)
      x << %(<w:pgNumType w:start="#{restart}"/>) if restart
      x << '<w:cols w:space="708"/>'
      x << '<w:docGrid w:linePitch="360"/>'
      x << '</w:sectPr>'
      last ? x : x
    end

    # ---------------------------------------------------- cabeçalho / rodapé
    def header_xml(title)
      self.part = :header
      logo = File.join(@images_dir, 'CNCS_positivo.png')
      w_mm = 20.0
      px = O.png_size(logo)
      h_mm = px ? w_mm * px[1] / px[0].to_f : 14.7
      cw = O.mm2twip(O::CONTENT_W_MM)
      c1 = (cw * 0.7).round
      c2 = cw - c1
      cell_pr = ->(w) { %(<w:tcPr><w:tcW w:w="#{w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>) }
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr #{NS.map { |k, v| %(#{k}="#{v}") }.join(' ')}>
        <w:tbl><w:tblPr><w:tblStyle w:val="TabelaSemBordas"/>
        <w:tblW w:w="#{cw}" w:type="dxa"/>
        #{S.borders(color: O::BORDER, sz: 0, sides: [])}
        <w:tblLayout w:type="fixed"/>
        <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>
        <w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>
        </w:tblPr>
        <w:tblGrid><w:gridCol w:w="#{c1}"/><w:gridCol w:w="#{c2}"/></w:tblGrid>
        <w:tr><w:trPr><w:cantSplit/></w:trPr>
        <w:tc>#{cell_pr.call(c1)}<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r>#{S.rpr(color: O::TEAL, pt: O::SMALL_PT)}<w:t xml:space="preserve">#{O.text(title)}</w:t></w:r></w:p></w:tc>
        <w:tc>#{cell_pr.call(c2)}<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/></w:pPr><w:r>#{drawing(logo, w_mm, h_mm, alt: 'CNCS')}</w:r></w:p></w:tc>
        </w:tr></w:tbl>
        <w:p><w:pPr><w:pStyle w:val="Header"/><w:spacing w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>
        </w:hdr>
      XML
    ensure
      self.part = :document
    end

    def footer_xml
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:ftr #{NS.map { |k, v| %(#{k}="#{v}") }.join(' ')}>
        <w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r>#{S.rpr(color: O::TEAL, pt: O::SMALL_PT)}<w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        </w:ftr>
      XML
    end

    # ------------------------------------------------------------- ficheiros
    def settings_xml
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="#{NS['xmlns:w']}">
        <w:zoom w:percent="100"/>
        <w:defaultTabStop w:val="708"/>
        <w:characterSpacingControl w:val="doNotCompress"/>
        <w:updateFields w:val="true"/>
        <w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>
        <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
        <w:themeFontLang w:val="#{O::LANG}"/>
        <w:decimalSymbol w:val=","/>
        <w:listSeparator w:val=";"/>
        </w:settings>
      XML
    end

    def footnotes_xml
      sep = ->(id, type, tag) do
        %(<w:footnote w:type="#{type}" w:id="#{id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/>) +
          %(<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:#{tag}/></w:r></w:p></w:footnote>)
      end
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:footnotes #{NS.map { |k, v| %(#{k}="#{v}") }.join(' ')}>
        #{sep.call(-1, 'separator', 'separator')}
        #{sep.call(0, 'continuationSeparator', 'continuationSeparator')}
        #{@footnotes.join}
        </w:footnotes>
      XML
    end

    # Ordem da sequência CT_CoreProperties (OPC): created, creator, keywords,
    # language, lastModifiedBy, modified, revision, title.
    def core_xml(title:, author:, keywords:, date:)
      stamp = "#{date}T00:00:00Z"
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
          xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
          xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        <dcterms:created xsi:type="dcterms:W3CDTF">#{stamp}</dcterms:created>
        <dc:creator>#{O.text(author)}</dc:creator>
        <cp:keywords>#{O.text(keywords)}</cp:keywords>
        <dc:language>#{O::LANG}</dc:language>
        <cp:lastModifiedBy>#{O.text(author)}</cp:lastModifiedBy>
        <dcterms:modified xsi:type="dcterms:W3CDTF">#{stamp}</dcterms:modified>
        <cp:revision>1</cp:revision>
        <dc:title>#{O.text(title)}</dc:title>
        </cp:coreProperties>
      XML
    end

    # ---------------------------------------------------- capa / contracapa
    COVER_MARGINS = { top: 22.86, right: 22.86, bottom: 0.0, left: 27.94 }.freeze
    BACK_MARGINS  = { top: 55.88, right: 17.018, bottom: 22.86, left: 17.018 }.freeze

    def logo_dims(name, width_mm)
      path = File.join(@images_dir, name)
      px = O.png_size(path)
      h = px ? width_mm * px[1] / px[0].to_f : width_mm * 0.735
      [path, width_mm, h]
    end

    def para(style, runs, extra_ppr: '', jc: nil)
      pr = +"<w:pStyle w:val=\"#{style}\"/>"
      pr << extra_ppr
      pr << %(<w:jc w:val="#{jc}"/>) if jc
      %(<w:p><w:pPr>#{pr}</w:pPr>#{runs}</w:p>)
    end

    def run(text, rpr)
      %(<w:r>#{rpr}<w:t xml:space="preserve">#{O.text(text)}</w:t></w:r>)
    end

    # Capa: fundo teal à sangria (imagem ancorada atrás do texto) e texto
    # totalmente editável, tal como no PDF de referência.
    def cover_xml(title:, subtitle:, date:, sect_pr:)
      bg = File.join(@assets_dir, 'fundo-teal.png')
      logo, lw, lh = logo_dims('CNCS_negativo.png', 82.0)
      x = +''
      x << %(<w:p><w:pPr><w:pStyle w:val="CapaTexto"/><w:spacing w:before="0" w:after="0"/>) <<
           %(<w:jc w:val="center"/></w:pPr>) <<
           %(<w:r>#{background_drawing(bg, O::PAGE_W_MM, O::PAGE_H_MM)}</w:r>) <<
           %(<w:r>#{drawing(logo, lw, lh, alt: 'CNCS')}</w:r></w:p>)
      x << para('CapaTitulo', run(title, S.rpr(bold: true, color: 'FFFFFF', pt: 40)),
                extra_ppr: %(<w:spacing w:before="#{O.mm2twip(73)}" w:after="#{O.pt2twip(6)}"/>))
      x << para('CapaSubtitulo', run(subtitle, S.rpr(color: 'FFFFFF', pt: 17)))
      # A data fecha a secção da capa (transporta as propriedades de secção).
      x << %(<w:p><w:pPr><w:pStyle w:val="CapaTexto"/>) <<
           %(<w:spacing w:before="#{O.mm2twip(76)}" w:after="0"/><w:jc w:val="right"/>) <<
           %(<w:rPr><w:color w:val="FFFFFF"/><w:sz w:val="22"/></w:rPr>#{sect_pr}</w:pPr>) <<
           run(date, S.rpr(color: 'FFFFFF', pt: 11)) << '</w:p>'
      x
    end

    # Contracapa (contracapa.adoc): logótipo, filete branco e blocos centrados.
    def back_cover_xml(lines:)
      bg = File.join(@assets_dir, 'fundo-teal.png')
      logo, lw, lh = logo_dims('CNCS_negativo.png', 55.0)
      x = +''
      x << %(<w:p><w:pPr><w:pStyle w:val="CapaTexto"/><w:spacing w:before="0" w:after="#{O.pt2twip(6)}"/>) <<
           %(<w:jc w:val="center"/></w:pPr>) <<
           %(<w:r>#{background_drawing(bg, O::PAGE_W_MM, O::PAGE_H_MM)}</w:r>) <<
           %(<w:r>#{drawing(logo, lw, lh, alt: 'CNCS')}</w:r></w:p>)
      # filete branco (equivalente ao thematic break do AsciiDoc)
      x << %(<w:p><w:pPr><w:pStyle w:val="CapaTexto"/>) <<
           '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="FFFFFF"/></w:pBdr>' <<
           %(<w:spacing w:before="#{O.pt2twip(6)}" w:after="#{O.pt2twip(14)}" w:line="20" w:lineRule="exact"/>) <<
           '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
      lines.each do |ln|
        rpr = S.rpr(bold: ln[:bold], color: 'FFFFFF', pt: O::BASE_PT)
        runs = ln[:parts].each_with_index.map { |t, i|
          (i.zero? ? '' : '<w:r><w:br/></w:r>') + run(t, rpr)
        }.join
        x << para('CapaTexto', runs,
                  extra_ppr: %(<w:spacing w:before="0" w:after="#{O.pt2twip(ln[:after] || 12)}"/>))
      end
      x
    end

    # ----------------------------------------------------------- ficheiros base
    def content_types
      exts = @media.values.map { |n| File.extname(n).delete('.').downcase }.uniq
      defaults = (%w[rels xml] + exts).uniq.map do |e|
        ct = case e
             when 'rels' then 'application/vnd.openxmlformats-package.relationships+xml'
             when 'xml'  then 'application/xml'
             when 'png'  then 'image/png'
             when 'jpg', 'jpeg' then 'image/jpeg'
             when 'gif'  then 'image/gif'
             when 'svg'  then 'image/svg+xml'
             else 'application/octet-stream'
             end
        %(<Default Extension="#{e}" ContentType="#{ct}"/>)
      end.join
      over = {
        '/word/document.xml'  => "#{CT}.document.main+xml",
        '/word/styles.xml'    => "#{CT}.styles+xml",
        '/word/numbering.xml' => "#{CT}.numbering+xml",
        '/word/settings.xml'  => "#{CT}.settings+xml",
        '/word/footnotes.xml' => "#{CT}.footnotes+xml",
        '/word/header1.xml'   => "#{CT}.header+xml",
        '/word/header2.xml'   => "#{CT}.header+xml",
        '/word/footer1.xml'   => "#{CT}.footer+xml",
        '/word/footer2.xml'   => "#{CT}.footer+xml",
        '/docProps/core.xml'  => 'application/vnd.openxmlformats-package.core-properties+xml',
        '/docProps/app.xml'   => 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
      }.map { |k, v| %(<Override PartName="#{k}" ContentType="#{v}"/>) }.join
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">#{defaults}#{over}</Types>
      XML
    end

    def rels_xml(part)
      items = @rels[part].map do |r|
        %(<Relationship Id="#{r[:id]}" Type="#{r[:type]}" Target="#{O.attr(r[:target])}") +
          (r[:mode] ? %( TargetMode="#{r[:mode]}") : '') + '/>'
      end.join
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">#{items}</Relationships>
      XML
    end

    def root_rels
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdDoc" Type="#{REL}/officeDocument" Target="word/document.xml"/>
        <Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
        <Relationship Id="rIdApp" Type="#{REL}/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
      XML
    end

    def document_xml(body)
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document #{NS.map { |k, v| %(#{k}="#{v}") }.join(' ')}><w:body>#{body}</w:body></w:document>
      XML
    end

    def empty_part(tag)
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:#{tag} #{NS.map { |k, v| %(#{k}="#{v}") }.join(' ')}>
        <w:p><w:pPr><w:pStyle w:val="#{tag == 'hdr' ? 'Header' : 'Footer'}"/><w:spacing w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>
        </w:#{tag}>
      XML
    end

    # Escreve o pacote .docx.
    def write!(out_path, files)
      stage = File.join(File.dirname(out_path), ".docx-stage-#{Process.pid}")
      FileUtils.rm_rf stage
      FileUtils.mkdir_p stage
      files.each do |rel, content|
        dest = File.join(stage, rel)
        FileUtils.mkdir_p File.dirname(dest)
        File.binwrite(dest, content)
      end
      @media.each do |abs, name|
        dest = File.join(stage, 'word', 'media', name)
        FileUtils.mkdir_p File.dirname(dest)
        FileUtils.cp abs, dest
      end
      FileUtils.rm_f out_path
      abs_out = File.expand_path(out_path)
      Dir.chdir(stage) do
        # -D: sem entradas de diretório (o OPC só descreve partes).
        ok = system('zip', '-q', '-X', '-D', '-0', abs_out, '[Content_Types].xml') &&
             system('zip', '-q', '-X', '-D', '-r', '-9', abs_out,
                    '_rels', 'docProps', 'word')
        raise 'falha ao criar o arquivo .docx' unless ok
      end
      FileUtils.rm_rf stage
      abs_out
    end

    # CT_Properties é uma sequência: Template, Company, …, Application,
    # AppVersion, DocSecurity. Não existe elemento <Title> (o título do
    # documento vive em docProps/core.xml) — incluí-lo faz o Word pedir para
    # recuperar o ficheiro.
    def app_xml(company:)
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
          xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        <Template>Normal.dotm</Template>
        <Company>#{O.text(company)}</Company>
        <Application>Asciidoctor - DOCX (CNCS)</Application>
        <AppVersion>16.0000</AppVersion>
        <DocSecurity>0</DocSecurity>
        </Properties>
      XML
    end
  end
end
