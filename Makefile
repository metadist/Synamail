.PHONY: help bootstrap dev sideload lint format check-types test test-e2e validate \
        build build-manifest generate-schemas ci-local clean deps doctor sync bridge

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

## ---------------------------------------------------------------------------
## Quality (mirrors CI)
## ---------------------------------------------------------------------------

lint: ## Run ESLint + Prettier
	@if [ -f package.json ]; then npm run lint; else echo "skip: no package.json (Sprint 2.1)"; fi

format: ## Apply Prettier formatting in place
	@if [ -f package.json ]; then npm run format; else echo "skip: no package.json (Sprint 2.1)"; fi

check-types: ## Run vue-tsc -b (catches errors ESLint misses)
	@if [ -f package.json ]; then npm run check:types; else echo "skip: no package.json (Sprint 2.1)"; fi

test: ## Run Vitest unit + component tests
	@if [ -f package.json ]; then npm run test -- --run; else echo "skip: no package.json (Sprint 2.1)"; fi

test-e2e: ## Run Playwright E2E suite (requires sideloadable manifest)
	@if [ -f package.json ]; then npm run test:e2e; else echo "skip: no package.json (Sprint 2.1)"; fi

validate: ## Validate manifest.xml (and unified manifest if present)
	@if [ -f manifest.xml ]; then \
	  npx --yes office-addin-manifest validate manifest.xml; \
	else \
	  echo "skip: no manifest.xml (Sprint 2.2)"; \
	fi
	@if [ -f manifest.unified.json ]; then \
	  npx --yes office-addin-manifest validate manifest.unified.json; \
	fi

## ---------------------------------------------------------------------------
## Build
## ---------------------------------------------------------------------------

build: ## Vite production build (also gates bundle size)
	@if [ -f package.json ]; then npm run build; else echo "skip: no package.json (Sprint 2.1)"; fi

build-manifest: ## Generate manifest.unified.json from manifest.xml (Sprint 4)
	npx --yes office-addin-manifest-converter manifest.xml -o manifest.unified.json

generate-schemas: ## Regenerate Zod schemas from Synaplan OpenAPI (Sprint 3)
	@if [ -f package.json ]; then npm run generate:schemas; else echo "skip: no package.json (Sprint 2.1)"; fi

clean: ## Remove build output + node_modules cache
	rm -rf dist coverage playwright-report test-results .vite

## ---------------------------------------------------------------------------
## Pre-commit gate (run before every commit, also runs in CI)
## ---------------------------------------------------------------------------

ci-local: ## Run the full pre-commit gate locally
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
	@echo ""
	@echo "ci-local: all gates passed"
