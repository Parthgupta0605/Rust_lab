TARGET = 2023CS11005_2023CS10507_2023CS10235
BUILD_DIR = target/release
all: build

# Build the project
build:
	cargo build --release

# Run the project
run: build
	@$(BUILD_DIR)/$(TARGET) 999 18278
	
ext1:
	@$(BUILD_DIR)/$(TARGET) -vim 100 100
# Run tests

test:
	cargo test

# Clean the build artifacts

docs:
	cargo doc
	@pdflatex report.tex

clean:
	cargo clean

# Format the code
fmt:
	cargo fmt

# Check for linting issues
lint:
	cargo clippy -- -D warnings

coverage:
	cargo tarpaulin --ignore-tests --exclude-files "src/extended.rs"

# Help message
help:
	@echo "Available targets:"
	@echo "  build   - Build the project"
	@echo "  run     - Run the project (use ARGS='...' to pass arguments)"
	@echo "  test    - Run tests"
	@echo "  clean   - Clean the build artifacts"
	@echo "  fmt     - Format the code"
	@echo "  lint    - Check for linting issues"
	@echo "  help    - Show this help message"
