#!/usr/bin/env bash
# Build da versão de análise do Guia do Método (CNCS).
# Passo 1: gera a contracapa (PDF A4, fundo teal) a partir de contracapa.adoc.
# Passo 2: gera o documento final, embebendo a contracapa via :back-cover-image:.
set -e
cd "$(dirname "$0")"
mkdir -p _build

# 1a) Capa -> images/capa.pdf (para :front-cover-image: image:capa.pdf[])
asciidoctor -a pdf-theme="$PWD/capa-theme.yml" \
  -a pdf-fontsdir="$PWD/fonts;GEM_FONTS_DIR" \
  -a imagesdir="$PWD/images" \
  -r asciidoctor-pdf -b pdf capa.adoc -o images/capa.pdf 2>&1 | grep -viE "deprecated" || true

# 1b) Contracapa -> images/contracapa.pdf (para :back-cover-image: image:contracapa.pdf[])
asciidoctor -a pdf-theme="$PWD/contracapa-theme.yml" \
  -a pdf-fontsdir="$PWD/fonts;GEM_FONTS_DIR" \
  -a imagesdir="$PWD/images" \
  -r asciidoctor-pdf -b pdf contracapa.adoc -o images/contracapa.pdf 2>&1 | grep -viE "deprecated" || true

# 2) Documento final
asciidoctor -a pdf-theme="$PWD/Monarc-theme.yml" \
  -a pdf-fontsdir="$PWD/fonts;GEM_FONTS_DIR" \
  -r asciidoctor-pdf -r ./cncs-admonition-ext.rb \
  -b pdf index.adoc -o _build/user-guide.pdf 2>&1 | grep -viE "deprecated" || true

echo "Capa:       images/capa.pdf"
echo "Contracapa: images/contracapa.pdf"
echo "PDF final:  _build/user-guide.pdf"
