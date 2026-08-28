# frozen_string_literal: true
#
# Calcula o número de página de cada título e de cada legenda a partir de uma
# renderização do .docx, para que o índice e a lista de figuras já abram
# preenchidos (os campos PAGEREF continuam a atualizar-se com F9).
#
#   ruby tools/docx/paginate.rb _build/user-guide.pdf \
#        _build/user-guide.entradas.json tools/docx/pages.json

require 'json'
require 'shellwords'

pdf, entries_path, out_path = ARGV
abort 'uso: paginate.rb <pdf> <entradas.json> <pages.json>' unless pdf && entries_path && out_path

pages = `pdftotext -layout #{Shellwords.escape(pdf)} - 2>/dev/null`.split("\f")
abort "não foi possível ler #{pdf}" if pages.empty?

# Linhas normalizadas de cada página (as legendas e os títulos ocupam a linha toda).
lines = pages.map { |p| p.split("\n").map { |l| l.strip.squeeze(' ') } }

flat = pages.map { |p| p.gsub(/\s+/, ' ').strip }

entries = JSON.parse(File.read(entries_path))
all = entries['indice'].map { |e| e.merge('kind' => 'indice') } +
      entries['figuras'].map { |e| e.merge('kind' => 'figuras') }

# O corpo começa na página do prefácio; antes disso só há índices, cujas
# entradas repetem os mesmos textos.
body_start = lines.index { |ls| ls.any? { |l| l == 'Histórico de alterações' } } || 0

map = {}
misses = []
cursor = body_start
all.each do |e|
  needle = e['text'].strip.squeeze(' ')
  # Primeiro por linha inteira (títulos e legendas ocupam a linha); em
  # seguida, com o texto da página normalizado, para os títulos que quebram
  # em duas linhas.
  found = (cursor...lines.length).find { |i| lines[i].include?(needle) } ||
          (body_start...lines.length).find { |i| lines[i].include?(needle) } ||
          (cursor...flat.length).find { |i| flat[i].include?(needle) } ||
          (body_start...flat.length).find { |i| flat[i].include?(needle) }
  if found
    map[e['bookmark']] = found + 1
    cursor = found
  else
    misses << e['text']
  end
end

File.write(out_path, JSON.pretty_generate(map))
warn "entradas sem página: #{misses.length}"
misses.first(8).each { |m| warn "  - #{m}" }
puts "#{map.size} páginas resolvidas -> #{out_path}"
