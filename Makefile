BASE_NAME=index
OUTPUT=_build

html: prepare
	cp -R images $(OUTPUT)/html
	asciidoctor *.adoc -D $(OUTPUT)/html

pdf: prepare
	asciidoctor -a pdf-style=Monarc-theme.yml \
		-r asciidoctor-pdf -b pdf *.adoc \
		-o $(OUTPUT)/pdf/user-guide.pdf

prepare:
	mkdir -p $(OUTPUT)

clean:
	rm -Rf $(OUTPUT)

serve:
	"$(shell which xdg-open || which open || which x-www-browser)" \
		http://localhost:8000/$(OUTPUT)/html
	python3 -m http.server 8000

# Versão editável em Word (.docx): estilos, índice, legendas e remissões
# como campos nativos do Word. Ver tools/docx/.
SOFFICE?=libreoffice

docx: prepare
	ruby tools/docx/adoc2docx.rb index.adoc $(OUTPUT)/user-guide.docx

# Verifica a integridade do pacote (o que faz o Word pedir recuperação).
# Com os esquemas oficiais descompactados: make docx-verificar XSD=/caminho/xsd
docx-verificar: docx
	python3 tools/docx/verificar.py $(OUTPUT)/user-guide.docx $(XSD)

# Recalcula os números de página guardados em tools/docx/pages.json (o índice e
# a lista de figuras abrem já preenchidos). Requer LibreOffice e poppler-utils.
docx-pages: docx
	$(SOFFICE) --headless --norestore --convert-to pdf \
		--outdir $(OUTPUT)/paginacao $(OUTPUT)/user-guide.docx
	ruby tools/docx/paginate.rb $(OUTPUT)/paginacao/user-guide.pdf \
		$(OUTPUT)/user-guide.entradas.json tools/docx/pages.json
	$(MAKE) docx
