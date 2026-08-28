# frozen_string_literal: true
#
# Geração de pacotes WordprocessingML (.docx) — camada de baixo nível.
#
# Espelha o aspecto definido em Monarc-theme.yml (paleta CNCS, Arial 10,5 pt,
# títulos teal, caixas de destaque coloridas) usando construções nativas do
# Word: estilos, listas multinível, campos (TOC/SEQ/REF/PAGE), notas de rodapé,
# cabeçalho/rodapé e tabelas com estilo.

require 'fileutils'
require 'securerandom'

module Ooxml
  # ---------------------------------------------------------------- unidades
  TWIP_PER_MM = 1440.0 / 25.4          # 1 pol = 1440 twips = 25,4 mm
  EMU_PER_MM  = 36_000
  MM_PER_PT   = 25.4 / 72.0

  def self.mm2twip(mm) = (mm * TWIP_PER_MM).round
  def self.mm2emu(mm)  = (mm * EMU_PER_MM).round
  def self.pt2half(pt) = (pt * 2).round      # tamanho de letra em meios-pontos
  def self.pt2twip(pt) = (pt * 20).round
  def self.pt2eighth(pt) = (pt * 8).round    # espessura de bordas em 1/8 pt

  # ------------------------------------------------------------------ paleta
  # Monarc-theme.yml → brand.*
  TEAL       = '308AA1'   # azul CNCS (títulos, ligações, cabeçalho/rodapé)
  TEAL_DARK  = '25697A'
  DARK       = '231F20'   # Preto Cyber (corpo de texto)
  SLATE      = '59717C'   # código inline
  WARN       = 'E0A800'
  ALERT      = 'C0392B'
  BORDER     = 'DDDDDD'   # bordas de tabela
  ADMON = {
    'note'      => { bg: 'E7F1F4', border: TEAL,  icon: 'icon-info.png'    },
    'tip'       => { bg: 'E7F1F4', border: TEAL,  icon: 'icon-tip.png'     },
    'warning'   => { bg: 'FDF6E3', border: WARN,  icon: 'icon-warning.png' },
    'caution'   => { bg: 'FBEAE8', border: ALERT, icon: 'icon-alert.png'   },
    'important' => { bg: 'FBEAE8', border: ALERT, icon: 'icon-alert.png'   }
  }.freeze

  # -------------------------------------------------------------- tipografia
  FONT       = 'Arial'
  FONT_MONO  = 'Consolas'
  BASE_PT    = 10.5                       # base.font_size
  LINE_PT    = 12.0                       # base.line_height_length
  SMALL_PT   = 9.0                        # round(base * 0.85)
  LARGE_PT   = 13.0                       # round(base * 1.25)
  CAPTION_PT = 10.0                       # base * 0,95, arredondado
  H_PT       = { 1 => 22.0, 2 => 18.0, 3 => 13.0, 4 => 10.5, 5 => 9.0 }.freeze
  COVER_TITLE_PT    = 27.0                # heading.h1 (floor(10,5 * 2,6))
  COVER_SUBTITLE_PT = 18.0                # heading.h3
  LANG = 'pt-PT'

  # --------------------------------------------------------------- página A4
  PAGE_W_MM = 210.0
  PAGE_H_MM = 297.0
  MARGIN_TOP_MM    = 25.4                 # 1,0 pol
  MARGIN_SIDE_MM   = 17.018               # 0,67 pol
  MARGIN_BOTTOM_MM = 17.018
  HEADER_DIST_MM   = 3.0
  FOOTER_DIST_MM   = 5.5
  CONTENT_W_MM = PAGE_W_MM - 2 * MARGIN_SIDE_MM   # ≈ 175,96 mm

  # ------------------------------------------------------------- utilitários
  # O Asciidoctor entrega texto já com entidades HTML; convertemo-lo para
  # UTF-8 puro e voltamos a escapar apenas o que o XML exige.
  NAMED_ENTITIES = {
    'amp' => '&', 'lt' => '<', 'gt' => '>', 'quot' => '"', 'apos' => "'",
    'nbsp' => " ", 'ndash' => "–", 'mdash' => "—",
    'hellip' => "…", 'lsquo' => "‘", 'rsquo' => "’",
    'ldquo' => "“", 'rdquo' => "”", 'dagger' => "†",
    'Dagger' => "‡", 'ast' => '*', 'copy' => "©", 'reg' => "®",
    'trade' => "™", 'deg' => "°", 'plusmn' => "±",
    'times' => "×", 'larr' => "←", 'rarr' => "→",
    'harr' => "↔", 'lArr' => "⇐", 'rArr' => "⇒",
    'hArr' => "⇔", 'check' => "✓", 'bull' => "•",
    'middot' => "·", 'euro' => "€", 'sect' => "§"
  }.freeze

  def self.decode_entities(str)
    str.gsub(/&(?:#(\d+)|#x([0-9a-fA-F]+)|([a-zA-Z][a-zA-Z0-9]*));/) do
      if $1 then [$1.to_i].pack('U')
      elsif $2 then [$2.to_i(16)].pack('U')
      else NAMED_ENTITIES[$3] || "&#{$3};"
      end
    end
  end

  def self.esc(str)
    str.gsub('&', '&amp;').gsub('<', '&lt;').gsub('>', '&gt;')
  end

  # Texto vindo do Asciidoctor → texto seguro para XML.
  def self.text(str)
    esc(decode_entities(str.to_s)).gsub(/[\x00-\x08\x0B\x0C\x0E-\x1F]/, '')
  end

  def self.attr(str)
    esc(decode_entities(str.to_s)).gsub('"', '&quot;')
  end

  # Nomes de marcador (bookmark) do Word: letras, dígitos e _ ; até 40 chars.
  def self.bookmark_name(id, taken)
    base = 'r_' + id.to_s.gsub(/[^A-Za-z0-9_]/, '_').sub(/\A_+/, '')
    base = base[0, 38]
    name = base
    n = 1
    while taken.include?(name)
      suffix = "_#{n += 1}"
      name = base[0, 38 - suffix.length] + suffix
    end
    taken << name
    name
  end

  # Campo do Word: begin / instrução / resultado em cache / end.
  def self.field(instr, cached = '')
    '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
      %(<w:r><w:instrText xml:space="preserve"> #{text(instr)} </w:instrText></w:r>) +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + cached +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  end

  # Dimensões de um PNG (IHDR).
  def self.png_size(path)
    return nil unless path && File.file?(path)
    data = File.binread(path, 33)
    return nil unless data && data[0, 8] == "\x89PNG\r\n\x1A\n".b
    data[16, 8].unpack('N2')
  end
end
