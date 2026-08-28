# CNCS admonition colouring extension for Asciidoctor PDF.
# Gives each admonition type a coloured box (background + border + left rule)
# so information/warning/alert stand out, per the CNCS review notes:
#   NOTE/TIP  -> azul/teal (informação)
#   WARNING   -> amarelo (aviso)
#   CAUTION/IMPORTANT -> vermelho (alerta)
require 'asciidoctor/pdf'

module CNCSAdmonitionColours
  PALETTE = {
    'note'      => { bg: 'E7F1F4', border: '308AA1', rule: '308AA1' },
    'tip'       => { bg: 'E7F1F4', border: '308AA1', rule: '308AA1' },
    'warning'   => { bg: 'FDF6E3', border: 'E0A800', rule: 'E0A800' },
    'caution'   => { bg: 'FBEAE8', border: 'C0392B', rule: 'C0392B' },
    'important'  => { bg: 'FBEAE8', border: 'C0392B', rule: 'C0392B' },
  }.freeze

  def convert_admonition node
    colours = PALETTE[(node.attr 'name')]
    return super unless colours

    theme = @theme
    saved = {
      bg:    theme.admonition_background_color,
      bcol:  theme.admonition_border_color,
      bwid:  theme.admonition_border_width,
      brad:  theme.admonition_border_radius,
      rule:  theme.admonition_column_rule_color,
      rwid:  theme.admonition_column_rule_width,
    }
    begin
      theme.admonition_background_color  = colours[:bg]
      theme.admonition_border_color      = colours[:border]
      theme.admonition_border_width      = (saved[:bwid] && saved[:bwid] > 0 ? saved[:bwid] : 0.5)
      theme.admonition_border_radius     = (saved[:brad] || 4)
      theme.admonition_column_rule_color = colours[:rule]
      theme.admonition_column_rule_width = 3
      super
    ensure
      theme.admonition_background_color  = saved[:bg]
      theme.admonition_border_color      = saved[:bcol]
      theme.admonition_border_width      = saved[:bwid]
      theme.admonition_border_radius     = saved[:brad]
      theme.admonition_column_rule_color = saved[:rule]
      theme.admonition_column_rule_width = saved[:rwid]
    end
  end
end

Asciidoctor::PDF::Converter.prepend CNCSAdmonitionColours
