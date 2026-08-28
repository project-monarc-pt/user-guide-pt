# MONARC User guide

[MONARC](http://monarc.lu) user guide.

In order to generate user guide, install
[Asciidoctor](http://asciidoctor.org/#installation):

```bash
$ sudo apt-get install asciidoctor
$ sudo gem install --pre asciidoctor-pdf
$ sudo gem install coderay pygments.rb

$ make html
$ make pdf
```

## Versão editável em Word (.docx)

```bash
$ make docx           # gera _build/user-guide.docx
$ make docx-pages     # idem, recalculando os números de página dos índices
$ make docx-verificar # verifica a integridade do pacote OOXML
```

`make docx-verificar` deteta o que faz o Word pedir a recuperação do ficheiro
(partes sem content type, relações por resolver, marcadores repetidos, campos
desequilibrados, propriedades fora do esquema) — problemas que não se notam ao
abrir o documento no LibreOffice. Valida também cada parte pelos esquemas
oficiais ECMA-376 Transitional, descarregados e guardados em `_build/xsd` na
primeira utilização. Para usar uma cópia local dos esquemas:

```bash
$ make docx-verificar XSD=/caminho/para/xsd
```

Para converter outro guia com este mesmo formato — copiar `tools/docx/` (sem o
`pages.json`, que é específico deste guia) e acrescentar os alvos com
`cat tools/docx/makefile-fragmento.mk >> Makefile` — seguir
[tools/docx/PLANO-CONVERSAO-DOCX.md](tools/docx/PLANO-CONVERSAO-DOCX.md).

### Os dois índices

| Índice | Campo | Alvo |
| --- | --- | --- |
| Índice | `TOC \o "1-2" \u` | níveis de destaque dos estilos *Título 1* e *Título 2* |
| Lista de Figuras | `TOC \c "Figura"` | sequência `SEQ Figura` das legendas de figura |

Os alvos são deliberadamente diferentes: as legendas de figura e as legendas de
tabela partilham o estilo *Legenda*, por isso um índice de figuras baseado no
estilo apanharia também as tabelas. Ao recolher pela sequência `Figura`, entram
apenas as figuras.

### Acrescentar ou remover figuras

Na fonte AsciiDoc basta acrescentar/remover o bloco `image::` com o seu título e
correr `make docx-pages`.

A editar diretamente no Word:

1. Insira a imagem e aplique-lhe o estilo *Imagem*.
2. **Referências › Inserir legenda**, com a legenda `Figura`, abaixo da imagem.
3. `Ctrl+A` e `F9` para atualizar tudo.

A numeração das figuras é um campo `SEQ`, e as remissões são campos
`REF`/`NOTEREF`: inserir uma figura a meio do documento renumera as seguintes,
atualiza a lista de figuras e corrige as remissões. Ao abrir o ficheiro, o Word
atualiza os campos automaticamente (`w:updateFields`); em alternativa, `Ctrl+A`
seguido de `F9`.
