PANDOC ?= pandoc
PANDOC_FLAGS ?= --from=gfm --to=html --wrap=none
PUBLIC_DIR := public
STATIC_DIR := static
TEMPLATES_DIR := templates
SITE_TITLE := pi-filler
DOC_NAMES := PROJECT ARCHITECTURE STYLE CHANGELOG PLAN
DOC_TARGETS := $(addprefix $(PUBLIC_DIR)/,$(addsuffix .html,$(DOC_NAMES)))
