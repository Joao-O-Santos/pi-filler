include config.mk

.PHONY: all site sync-static verify clean

all: site

site: $(PUBLIC_DIR)/index.html $(DOC_TARGETS) sync-static

$(PUBLIC_DIR):
	@mkdir -p $(PUBLIC_DIR)

$(PUBLIC_DIR)/index.html: README.md templates/header.html templates/menu.html templates/footer.html | $(PUBLIC_DIR)
	@sed 's|@@TITLE@@|$(SITE_TITLE)|g' templates/header.html > $@
	@cat templates/menu.html >> $@
	@$(PANDOC) $(PANDOC_FLAGS) $< >> $@
	@cat templates/footer.html >> $@

$(PUBLIC_DIR)/%.html: %.md templates/header.html templates/menu.html templates/footer.html | $(PUBLIC_DIR)
	@sed 's|@@TITLE@@|$(SITE_TITLE)|g' templates/header.html > $@
	@cat templates/menu.html >> $@
	@$(PANDOC) $(PANDOC_FLAGS) $< >> $@
	@cat templates/footer.html >> $@

sync-static: | $(PUBLIC_DIR)
	@cp -R $(STATIC_DIR)/. $(PUBLIC_DIR)/

verify:
	@npm run check
	@npm pack --dry-run

clean:
	@rm -rf $(PUBLIC_DIR)
