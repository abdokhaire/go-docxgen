TEST_DIR ?= ./...
COVERAGE_FILE = coverage.out
COVERAGE_HTML = coverage.html

# Default version for release (override with: make release VERSION=v0.1.3)
VERSION ?= v0.0.0

.PHONY: all build test test-coverage test-clean benchmark lint deps deps-update deps-check deps-verify release release-patch release-minor release-major

all: deps-verify build test

# Build
build:
	@echo "Building..."
	@go build ./...

# Testing
test:
	@echo "Running tests..."
	@go test $(TEST_DIR)

test-coverage:
	@echo "Running tests with coverage..."
	@go test -coverprofile=$(COVERAGE_FILE) $(TEST_DIR)
	@echo "Generating HTML coverage report..."
	@go tool cover -html=$(COVERAGE_FILE) -o $(COVERAGE_HTML)
	@echo "Coverage file generated at: $(CURDIR)/$(COVERAGE_HTML)"
	@if grep -qi microsoft /proc/version; then \
		powershell.exe Start-Process \"chrome\" \"$(shell wslpath -w $(CURDIR)/$(COVERAGE_HTML) | sed 's/\\/\\\\/g')\" || \
		powershell.exe Start-Process \"msedge\" \"$(shell wslpath -w $(CURDIR)/$(COVERAGE_HTML) | sed 's/\\/\\\\/g')\" || \
		powershell.exe Start-Process \"firefox\" \"$(shell wslpath -w $(CURDIR)/$(COVERAGE_HTML) | sed 's/\\/\\\\/g')\"; \
	elif command -v xdg-open > /dev/null; then \
		xdg-open $(COVERAGE_HTML) > /dev/null 2>&1 & \
	elif command -v open > /dev/null; then \
		open $(COVERAGE_HTML); \
	else \
		echo "Could not detect how to open the coverage report."; \
	fi

test-clean:
	@rm -f $(COVERAGE_FILE) $(COVERAGE_HTML)
	@echo "Cleaned up coverage files."

benchmark:
	@go test -bench=Benchmark* -benchtime=1x

# Linting
lint:
	@echo "Running linter..."
	@golangci-lint run ./... || echo "Install golangci-lint: go install github.com/golangci/golangci-lint/cmd/golangci-lint@latest"

# Dependency Management
deps:
	@echo "Downloading dependencies..."
	@go mod download

deps-tidy:
	@echo "Tidying dependencies..."
	@go mod tidy

deps-update:
	@echo "Updating all dependencies..."
	@go get -u ./...
	@go mod tidy
	@echo "Dependencies updated. Run 'make test' to verify compatibility."

deps-check:
	@echo "Checking for dependency updates..."
	@go list -u -m all

deps-verify:
	@echo "Verifying dependencies..."
	@go mod verify
	@echo "All dependencies verified."

# Vendor dependencies (prevents unexpected updates)
vendor:
	@echo "Vendoring dependencies..."
	@go mod vendor
	@echo "Dependencies vendored. Add '-mod=vendor' to go commands to use."

# Release Management
release:
	@if [ "$(VERSION)" = "v0.0.0" ]; then \
		echo "Error: Please specify VERSION, e.g., make release VERSION=v0.1.3"; \
		exit 1; \
	fi
	@echo "Preparing release $(VERSION)..."
	@go mod tidy
	@go build ./...
	@go test ./...
	@git add .
	@git commit -m "Release $(VERSION)" || echo "No changes to commit"
	@git tag $(VERSION)
	@echo "Created tag $(VERSION)"
	@echo "Run 'make push VERSION=$(VERSION)' to push to remote"

push:
	@if [ "$(VERSION)" = "v0.0.0" ]; then \
		echo "Error: Please specify VERSION, e.g., make push VERSION=v0.1.3"; \
		exit 1; \
	fi
	@echo "Pushing $(VERSION) to remote..."
	@git push origin master
	@git push origin $(VERSION)
	@echo "Release $(VERSION) pushed successfully!"
	@echo "View at: https://pkg.go.dev/github.com/abdokhaire/go-docxgen@$(VERSION)"

# Get current version from latest tag
version:
	@git describe --tags --abbrev=0 2>/dev/null || echo "No tags found"

# Show next version suggestions
next-version:
	@echo "Current version: $$(git describe --tags --abbrev=0 2>/dev/null || echo 'none')"
	@echo "Suggested next versions:"
	@CURRENT=$$(git describe --tags --abbrev=0 2>/dev/null | sed 's/v//'); \
	if [ -n "$$CURRENT" ]; then \
		MAJOR=$$(echo $$CURRENT | cut -d. -f1); \
		MINOR=$$(echo $$CURRENT | cut -d. -f2); \
		PATCH=$$(echo $$CURRENT | cut -d. -f3); \
		echo "  Patch: v$$MAJOR.$$MINOR.$$((PATCH + 1))"; \
		echo "  Minor: v$$MAJOR.$$((MINOR + 1)).0"; \
		echo "  Major: v$$((MAJOR + 1)).0.0"; \
	else \
		echo "  Initial: v0.1.0"; \
	fi

# CI check - run before committing
ci: deps-verify build test lint
	@echo "All CI checks passed!"

# Help
help:
	@echo "Available targets:"
	@echo "  build          - Build the project"
	@echo "  test           - Run tests"
	@echo "  test-coverage  - Run tests with coverage report"
	@echo "  lint           - Run linter"
	@echo "  deps           - Download dependencies"
	@echo "  deps-tidy      - Tidy go.mod and go.sum"
	@echo "  deps-update    - Update all dependencies (CAREFUL!)"
	@echo "  deps-check     - Check for available updates"
	@echo "  deps-verify    - Verify dependency checksums"
	@echo "  vendor         - Vendor dependencies locally"
	@echo "  release        - Create a release (VERSION=vX.Y.Z required)"
	@echo "  push           - Push release to remote (VERSION=vX.Y.Z required)"
	@echo "  version        - Show current version"
	@echo "  next-version   - Suggest next version numbers"
	@echo "  ci             - Run all CI checks"
	@echo ""
	@echo "Example release workflow:"
	@echo "  make release VERSION=v0.1.3"
	@echo "  make push VERSION=v0.1.3"
