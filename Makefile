.PHONY: help bootstrap dev sideload lint lint-docs format check-types test test-e2e test-e2e-live validate \
        build build-manifest generate-schemas ci-local clean deps doctor sync bridge \
        up down budget sync-plugin sync-plugin-and-clear

# Synaplan-side companion plugin (synamail-plugin/) — released into the main
# Synaplan repo's plugins/ directory, same flow as Synaform.
SYNAPLAN_DIR ?= /wwwroot/synaplan
PLUGIN_SRC    = synamail-plugin
PLUGIN_DST    = $(SYNAPLAN_DIR)/plugins/synamail

# Default target — print help.
help: ## Show this help
	@echo "Synamail — common commands:"
	@echo ""
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) \
	  | sort \
	  | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-18s\033[0m %s\n", $$1, $$2}'
	@echo ""
	@echo "Run 'make doctor' to check your local toolchain."

## ---------------------------------------------------------------------------
## Bootstrap
## ---------------------------------------------------------------------------

bootstrap: ## One-time setup: enable git hooks, install deps if available
	git config core.hooksPath .githooks
	@echo "git hooks enabled (.githooks)"
	@if [ -f package.json ]; then \
	  npm install; \
	else \
	  echo "no package.json yet — run 'make bootstrap' again after Sprint 2.1 scaffold lands"; \
	fi

deps: ## Install dependencies (alias for `npm ci`)
	@if [ -f package.json ]; then npm ci; else echo "no package.json yet — skipping"; fi

doctor: ## Verify your local toolchain (node, npm, git hooks)
	@echo "== Local toolchain =="
	@node --version || (echo "::error::node not found"; exit 1)
	@node -e 'const [m] = process.versions.node.split(".").map(Number); if (m < 22) { console.error("::warning::Node " + process.versions.node + " is below the minimum (>=22). Active LTS is Node 24; see .nvmrc."); }'
	@npm --version  || (echo "::error::npm not found";  exit 1)
	@git --version
	@echo ""
	@echo "== Git hooks =="
	@configured=$$(git config core.hooksPath || true); \
	if [ "$$configured" = ".githooks" ]; then \
	  echo "OK: hooks path = .githooks"; \
	else \
	  echo "::warning::run 'make bootstrap' to enable hooks (current: $$configured)"; \
	fi
	@echo ""
	@echo "== Repo state =="
	@[ -f package.json ] && echo "package.json: yes" || echo "package.json: no (Sprint 2.1)"
	@[ -f manifest.xml ] && echo "manifest.xml: yes" || echo "manifest.xml: no (Sprint 2.2)"
	@[ -d src ] && echo "src/: yes" || echo "src/: no (Sprint 2.x)"

## ---------------------------------------------------------------------------
## Development
## ---------------------------------------------------------------------------

dev: ## Run the Vite dev server on https://localhost:3000
	npm run dev

sideload: ## Sideload manifest.xml into Outlook on the Web
	npm run sideload

sync: ## (WSL) Copy a bumped manifest to C:\addin-catalog — only when manifest.xml itself changes
	@scripts/win-sync.sh

bridge: ## HTTPS-terminate the local Synaplan frontend on :5174 for the dev sign-in loop
	@scripts/dev-bridge-proxy.sh

up: ## One-shot: bring up Synaplan Docker + Synamail taskpane + HTTPS sign-in bridge (Mac/Linux/WSL)
	@./start-dev.sh

down: ## Stop Synamail dev servers (Synaplan Docker is left running)
	@./start-dev.sh stop

## ---------------------------------------------------------------------------
## Quality (mirrors CI)
## ---------------------------------------------------------------------------

lint: ## Run ESLint + Prettier
	@if [ -f package.json ]; then npm run lint; else echo "skip: no package.json (Sprint 2.1)"; fi

lint-docs: ## Run markdownlint exactly like CI's "Docs lint" job (uses .markdownlint.jsonc)
	npx --yes markdownlint-cli2 '**/*.md' '!node_modules/**' '!dist/**' '!.github/copilot-instructions.md'

format: ## Apply Prettier formatting in place
	@if [ -f package.json ]; then npm run format; else echo "skip: no package.json (Sprint 2.1)"; fi

check-types: ## Run vue-tsc -b (catches errors ESLint misses)
	@if [ -f package.json ]; then npm run check:types; else echo "skip: no package.json (Sprint 2.1)"; fi

test: ## Run Vitest unit + component tests
	@if [ -f package.json ]; then npm run test -- --run; else echo "skip: no package.json (Sprint 2.1)"; fi

test-e2e: ## Run Playwright E2E suite (taskpane + Office shim, mocked Synaplan)
	@if [ -f package.json ]; then npm run test:e2e; else echo "skip: no package.json (Sprint 2.1)"; fi

test-e2e-live: ## Run the E2E flows against a real Synaplan (set SYNAPLAN_BASE_URL + SYNAPLAN_API_KEY)
	@SYNAPLAN_E2E_LIVE=1 npm run test:e2e

validate: ## Validate the dev + production manifests (and unified manifest if present)
	@if [ -f manifest.xml ]; then \
	  npx --yes office-addin-manifest validate manifest.xml; \
	else \
	  echo "skip: no manifest.xml (Sprint 2.2)"; \
	fi
	@if [ -f manifest.prod.xml ]; then \
	  npx --yes office-addin-manifest validate manifest.prod.xml; \
	fi
	@if [ -f manifest.unified.json ]; then \
	  npx --yes office-addin-manifest validate manifest.unified.json; \
	fi

## ---------------------------------------------------------------------------
## Build
## ---------------------------------------------------------------------------

build: ## Vite production build (also gates bundle size)
	@if [ -f package.json ]; then npm run build; else echo "skip: no package.json (Sprint 2.1)"; fi

budget: ## Enforce the dist bundle-size budget (mirrors CI; cross-platform)
	@node -e "const fs=require('fs'),p=require('path');if(!fs.existsSync('dist')){console.log('budget: no dist/ — run make build first');process.exit(0)}let t=0;const w=d=>{for(const e of fs.readdirSync(d,{withFileTypes:true})){const f=p.join(d,e.name);e.isDirectory()?w(f):t+=fs.statSync(f).size}};w('dist');const b=2*1024*1024;console.log('dist size: '+t+' bytes (budget: '+b+')');if(t>b){console.error('::error::Bundle exceeds budget');process.exit(1)}console.log('budget: within 2 MiB')"

build-manifest: ## Generate manifest.unified.json from the PRODUCTION manifest (store submission)
	npx --yes office-addin-manifest-converter convert manifest.prod.xml -o manifest.unified.json

generate-schemas: ## Regenerate Zod schemas from Synaplan OpenAPI (Sprint 3)
	@if [ -f package.json ]; then npm run generate:schemas; else echo "skip: no package.json (Sprint 2.1)"; fi

## ---------------------------------------------------------------------------
## Synaplan plugin (Contact AI Profiling backend)
## ---------------------------------------------------------------------------

sync-plugin: ## Release synamail-plugin/ into $(SYNAPLAN_DIR)/plugins/synamail
	rm -rf $(PLUGIN_DST)
	cp -r $(PLUGIN_SRC) $(PLUGIN_DST)
	@echo "Synced to $(PLUGIN_DST)"
	@# plugins/synamail is TRACKED in the synaplan repo (like plugins/synaform).
	@# Surface the drift right away so the sync never gets swept silently into
	@# an unrelated synaplan commit — release it as its own dedicated commit.
	@if git -C $(SYNAPLAN_DIR) rev-parse --is-inside-work-tree >/dev/null 2>&1; then \
	  status=$$(git -C $(SYNAPLAN_DIR) status --short -- plugins/synamail); \
	  if [ -n "$$status" ]; then \
	    echo ""; \
	    echo "Uncommitted plugin changes in $(SYNAPLAN_DIR):"; \
	    echo "$$status"; \
	    echo "→ commit them there as a dedicated commit, e.g.:"; \
	    echo "  git -C $(SYNAPLAN_DIR) add plugins/synamail && git -C $(SYNAPLAN_DIR) commit -m 'feat(plugins): update synamail plugin to <version>'"; \
	  else \
	    echo "synaplan repo: plugins/synamail is clean (in sync with HEAD)"; \
	  fi; \
	fi

sync-plugin-and-clear: sync-plugin ## Sync the plugin and clear the Synaplan Symfony cache
	docker compose -f $(SYNAPLAN_DIR)/docker-compose.yml exec -T backend php bin/console cache:clear
	@echo "Cache cleared"

clean: ## Remove build output + node_modules cache
	rm -rf dist coverage playwright-report test-results .vite

## ---------------------------------------------------------------------------
## Pre-commit gate (run before every commit, also runs in CI)
## ---------------------------------------------------------------------------

ci-local: ## Run the full pre-commit gate locally
	@echo "→ make lint-docs"
	@$(MAKE) lint-docs
	@echo "→ make lint"
	@$(MAKE) lint
	@echo "→ make check-types"
	@$(MAKE) check-types
	@echo "→ make test"
	@$(MAKE) test
	@echo "→ make validate"
	@$(MAKE) validate
	@echo "→ make build"
	@$(MAKE) build
	@echo "→ make budget"
	@$(MAKE) budget
	@echo ""
	@echo "ci-local: all gates passed"
