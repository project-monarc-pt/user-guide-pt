# frozen_string_literal: true
#
# Converte o Guia do Utilizador MONARC (AsciiDoc) para .docx mantendo o aspecto
# do PDF (tema Monarc-theme.yml) e usando construções nativas do Word.
#
#   ruby tools/docx/adoc2docx.rb [index.adoc] [saída.docx]

require_relative 'ooxml'
require_relative 'styles'
require_relative 'package'
require_relative 'converter'

module Adoc2Docx
  O = Ooxml
  S = Ooxml::Styles

  # Texto da capa/contracapa lido dos próprios ficheiros AsciiDoc.
  def self.cover_texts(base_dir, doc)
    capa = File.join(base_dir, 'capa.adoc')
    src = File.file?(capa) ? File.read(capa) : ''
    title = src[/^\[\.covertitle\]\s*\n(.+)$/, 1] || doc.doctitle
    subtitle = src[/^\[\.coversubtitle\]\s*\n(.+)$/, 1] || ''
    [doc.sub_attributes(title.strip), doc.sub_attributes(subtitle.strip)]
  end

  def self.back_cover_lines(base_dir, doc)
    path = File.join(base_dir, 'contracapa.adoc')
    return [] unless File.file?(path)
    lines = []
    back = Asciidoctor.load_file(path, safe: :unsafe, parse: true,
                                attributes: { 'revnumber' => doc.attr('revnumber', '1.0'),
                                              'docdate' => doc.attr('docdate') })
    back.blocks.each do |b|
      next unless b.context == :paragraph
      src = back.sub_attributes(b.source.to_s)
      next if src.strip.empty? || src.include?('image::')
      parts = src.split(/ \+\n/).map { |p| p.strip.delete_prefix('*').delete_suffix('*') }
      bold = src.strip.start_with?('*') && src.strip.end_with?('*')
      lines << { parts: parts, bold: bold, after: 12 }
    end
    lines
  end

  # ------------------------------------------------- entradas dos dois índices
  # Um campo TOC cujo resultado ocupa vários parágrafos guarda-se com o
  # fldChar "begin" no primeiro parágrafo e o "end" no último — é assim que o
  # Word grava um índice preenchido. As entradas ficam ligadas por
  # hiperligação ao marcador de destino e o número de página é um campo
  # PAGEREF, pelo que tudo se atualiza com F9.
  def self.titulo_frontmatter(texto)
    %(<w:p><w:pPr><w:pStyle w:val="TituloFrontmatter"/></w:pPr>) +
      %(<w:r><w:t xml:space="preserve">#{O.text(texto)}</w:t></w:r></w:p>)
  end

  def self.index_field(instr, entries, pages, style_for)
    empty_para = ->(inner) {
      %(<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>#{inner}</w:p>)
    }
    return empty_para.call(O.field(instr, '')) if entries.empty?

    open_field = '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
                 %(<w:r><w:instrText xml:space="preserve"> #{O.text(instr)} </w:instrText></w:r>) +
                 '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    close_field = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

    entries.each_with_index.map do |e, i|
      runs = +%(<w:r><w:t xml:space="preserve">#{O.text(e[:text])}</w:t></w:r>)
      if (page = pages[e[:bookmark]])
        runs << '<w:r><w:tab/></w:r>'
        runs << O.field("PAGEREF #{e[:bookmark]} \\h",
                        %(<w:r><w:t>#{O.text(page)}</w:t></w:r>))
      end
      body = %(<w:hyperlink w:anchor="#{e[:bookmark]}" w:history="1">#{runs}</w:hyperlink>)
      body = open_field + body if i.zero?
      body += close_field if i == entries.length - 1
      %(<w:p><w:pPr><w:pStyle w:val="#{style_for.call(e)}"/></w:pPr>#{body}</w:p>)
    end.join
  end

  # Mapa marcador => página, produzido por tools/docx/paginate.rb. Um ficheiro
  # de outro guia (ou desatualizado) é ignorado em vez de escrever números
  # errados: só é aceite se cobrir quase todas as entradas deste documento.
  def self.page_map(path, entries, pkg)
    return {} unless path && File.file?(path)
    require 'json'
    mapa = JSON.parse(File.read(path))
    return {} if entries.empty?
    cobertura = entries.count { |e| mapa.key?(e[:bookmark]) }.to_f / entries.length
    if cobertura < 0.8
      pkg.warn! format('%s não corresponde a este documento (cobre %d%% das ' \
                       'entradas); ignorado — correr o alvo docx-pages',
                       File.basename(path), (cobertura * 100).round)
      return {}
    end
    mapa
  rescue StandardError => e
    pkg.warn! "#{File.basename(path)} ilegível (#{e.class}); ignorado"
    {}
  end

  def self.converter_entries(doc, kind)
    conv = doc.converter
    kind == :toc ? conv.toc_entries : conv.fig_entries
  end

  def self.run(index_path, out_path, pages_path = nil)
    base_dir = File.dirname(File.expand_path(index_path))
    assets_dir = File.join(__dir__, 'assets')
    pkg = Ooxml::Package.new(images_dir: File.join(base_dir, 'images'),
                             assets_dir: assets_dir)

    # Relações fixas do document.xml (declaradas antes das imagens/ligações).
    rel = {
      styles: pkg.add_rel('styles', 'styles.xml'),
      numbering: pkg.add_rel('numbering', 'numbering.xml'),
      settings: pkg.add_rel('settings', 'settings.xml'),
      footnotes: pkg.add_rel('footnotes', 'footnotes.xml'),
      header_body: pkg.add_rel('header', 'header1.xml'),
      header_none: pkg.add_rel('header', 'header2.xml'),
      footer_body: pkg.add_rel('footer', 'footer1.xml'),
      footer_none: pkg.add_rel('footer', 'footer2.xml')
    }

    # Os rótulos vêm do próprio AsciiDoc, para o conversor servir qualquer guia
    # com esta estrutura. O sufixo @ nos atributos deixa o documento decidir.
    # parse_header_only: parse:false não leria o cabeçalho e os rótulos sairiam
    # com os valores por omissão do Asciidoctor (em inglês).
    probe = Asciidoctor.load_file(index_path, safe: :unsafe, parse_header_only: true,
                                  attributes: { 'imagesdir' => 'images@',
                                                'includedir' => '_includes@' })
    labels = {
      chapter: probe.attr('chapter-signifier', 'Capítulo'),
      figure: probe.attr('figure-caption', 'Figura'),
      figures: probe.attr('list-of-figures-title', 'Lista de Figuras'),
      sectnumlevels: probe.attr('sectnumlevels', 3).to_i
    }
    DocxConverter.context = { pkg: pkg, base_dir: base_dir, labels: labels }
    doc = Asciidoctor.load_file(index_path, safe: :unsafe, backend: 'docx',
                                converter: DocxConverter, doctype: 'book',
                                attributes: { 'imagesdir' => 'images@',
                                              'includedir' => '_includes@' })
    body = doc.convert
    entradas_toc = converter_entries(doc, :toc)
    entradas_fig = converter_entries(doc, :fig)
    pages = page_map(pages_path, entradas_toc + entradas_fig, pkg)
    body = body.sub(DocxConverter::TOC_MARK,
                    titulo_frontmatter('Índice') +
                    index_field(%(TOC \\o "1-2" \\h \\z \\u), entradas_toc,
                                pages, ->(e) { "TOC#{e[:level]}" }))
    # Sem figuras não há lista de figuras: o bloco desaparece por inteiro.
    body = body.sub(DocxConverter::TOF_MARK,
                    entradas_fig.empty? ? '' :
                    titulo_frontmatter(labels[:figures]) +
                    index_field(%(TOC \\h \\z \\c "#{labels[:figure]}"), entradas_fig,
                                pages, ->(_e) { 'TableofFigures' }))

    title = doc.doctitle || 'Guia do Utilizador'
    author = doc.author || 'Centro Nacional de Cibersegurança (CNCS)'
    docdate = doc.attr('docdate').to_s[0, 10]
    cover_title, cover_subtitle = cover_texts(base_dir, doc)

    sect_cover = pkg.sect_pr(header: rel[:header_none], footer: rel[:footer_none],
                             margins: Ooxml::Package::COVER_MARGINS)
    sect_body  = pkg.sect_pr(header: rel[:header_body], footer: rel[:footer_body],
                             restart: 1)
    sect_back  = pkg.sect_pr(header: rel[:header_none], footer: rel[:footer_none],
                             margins: Ooxml::Package::BACK_MARGINS)

    closing = %(<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>) +
              %(<w:rPr><w:sz w:val="2"/></w:rPr>#{sect_body}</w:pPr></w:p>)

    doc_body = pkg.cover_xml(title: cover_title, subtitle: cover_subtitle,
                             date: docdate, sect_pr: sect_cover) +
               body + closing +
               pkg.back_cover_xml(lines: back_cover_lines(base_dir, doc)) +
               sect_back

    [DocxConverter::TOC_MARK, DocxConverter::TOF_MARK].each do |marca|
      raise "marca #{marca} não substituída (#{body.scan(marca).size}x)" if body.include?(marca)
    end

    # As partes que consomem imagens têm de ser geradas antes dos rels.
    header1 = pkg.header_xml(title)
    footer1 = pkg.footer_xml

    files = {
      '[Content_Types].xml' => nil, # preenchido depois de conhecer os media
      '_rels/.rels' => pkg.root_rels,
      'word/document.xml' => pkg.document_xml(doc_body),
      'word/styles.xml' => S.styles_xml,
      'word/numbering.xml' => S.numbering_xml(pkg.numbering_defs, labels[:chapter],
                                              labels[:sectnumlevels]),
      'word/settings.xml' => pkg.settings_xml,
      'word/footnotes.xml' => pkg.footnotes_xml,
      'word/header1.xml' => header1,
      'word/header2.xml' => pkg.empty_part('hdr'),
      'word/footer1.xml' => footer1,
      'word/footer2.xml' => pkg.empty_part('ftr'),
      'word/_rels/document.xml.rels' => pkg.rels_xml(:document),
      'word/_rels/header1.xml.rels' => pkg.rels_xml(:header),
      'docProps/core.xml' => pkg.core_xml(title: title, author: author,
                                          keywords: doc.attr('keywords').to_s,
                                          date: docdate),
      'docProps/app.xml' => pkg.app_xml(company: author)
    }
    files['[Content_Types].xml'] = pkg.content_types

    out = pkg.write!(out_path, files)
    # Lista ordenada das entradas: alimenta tools/docx/paginate.rb.
    require 'json'
    File.write(out.sub(/\.docx\z/, '') + '.entradas.json',
               JSON.pretty_generate('indice' => converter_entries(doc, :toc),
                                    'figuras' => converter_entries(doc, :fig)))
    [out, pkg.warnings]
  end
end

if $PROGRAM_NAME == __FILE__
  index = ARGV[0] || 'index.adoc'
  out = ARGV[1] || File.join('_build', 'user-guide.docx')
  pages = ARGV[2] || File.join(__dir__, 'pages.json')
  path, warnings = Adoc2Docx.run(index, out, pages)
  warnings.each { |w| $stderr.puts "aviso: #{w}" }
  size = (File.size(path) / 1024.0 / 1024.0).round(2)
  puts "DOCX gerado: #{path} (#{size} MB)"
end
