# Alvos da versão editável em Word (.docx). Acrescentar ao Makefile do guia:
#     cat tools/docx/makefile-fragmento.mk >> Makefile
# Ajustar ENTRADA e SAIDA ao guia; OUTPUT é a pasta de build (por norma _build).

ENTRADA?=index.adoc
SAIDA?=$(OUTPUT)/user-guide.docx
SOFFICE?=libreoffice

docx: prepare
	ruby tools/docx/adoc2docx.rb $(ENTRADA) $(SAIDA)

# Verifica a integridade do pacote (o que faz o Word pedir recuperação).
docx-verificar: docx
	python3 tools/docx/verificar.py $(SAIDA) $(XSD)

# Recalcula os números de página do índice e da lista de figuras.
docx-pages: docx
	$(SOFFICE) --headless --norestore --convert-to pdf \
		--outdir $(OUTPUT)/paginacao $(SAIDA)
	ruby tools/docx/paginate.rb \
		$(OUTPUT)/paginacao/$(notdir $(basename $(SAIDA))).pdf \
		$(basename $(SAIDA)).entradas.json tools/docx/pages.json
	$(MAKE) docx
