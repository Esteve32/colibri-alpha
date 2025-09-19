# Colibri Alpha Development Instructions

Colibri Alpha is a repository for alpha versions of browser and software elements for testing and iterative design.

**ALWAYS** follow these instructions first and only fallback to additional search and context gathering if the information in these instructions is incomplete or found to be in error.

## Environment Setup

The development environment includes:
- Node.js v20.19.5 with npm 10.8.2
- Python 3.12.3 with pip 24.0  
- Go 1.24.7
- Rust/Cargo 1.89.0
- Standard build tools (make, gcc, etc.)

## Current Repository State

As of the latest update, this repository contains:
- `README.md` - Basic project description
- `.github/` - GitHub configuration directory
- This repository is currently minimal and serves as a foundation for alpha software development

## Working Effectively

### Initial Repository Setup
- Clone: `git clone https://github.com/Esteve32/colibri-alpha.git`
- Navigate: `cd colibri-alpha`
- Check status: `git --no-pager status`

### Common Development Workflows

#### For JavaScript/Node.js Projects (when added)
- Initialize: `npm init -y` (takes ~0.5 seconds)
- Install dependencies: `npm install` (timing varies by dependencies - typically 10-60 seconds)
- Build: `npm run build` (if build script exists - timing varies, expect 30 seconds to 5 minutes)
- Test: `npm test` (timing varies by test suite - typically 5-30 seconds for unit tests)
- Dev server: `npm run dev` or `npm start` (if scripts exist)
- Lint: `npm run lint` (if configured - typically 5-15 seconds)
- Format: `npm run format` (if configured - typically 2-10 seconds)

**CRITICAL BUILD TIMINGS:**
- `npm install` for large projects: **NEVER CANCEL** - can take 5-15 minutes. Set timeout to 20+ minutes.
- `npm run build` for complex projects: **NEVER CANCEL** - can take 2-10 minutes. Set timeout to 15+ minutes.
- `npm test` for full test suites: **NEVER CANCEL** - can take 1-10 minutes. Set timeout to 15+ minutes.

#### For Python Projects (when added)
- Setup virtual environment: `python3 -m venv venv` (takes ~5-10 seconds)
- Activate: `source venv/bin/activate`
- Install dependencies: `pip install -r requirements.txt` (timing varies - typically 30 seconds to 5 minutes)
- Run tests: `python -m pytest` or `python -m unittest` (timing varies by test suite)
- Format: `black .` (if using Black formatter)
- Lint: `flake8 .` or `pylint .` (if configured)

**CRITICAL BUILD TIMINGS:**
- `pip install` for data science projects: **NEVER CANCEL** - can take 10-30 minutes. Set timeout to 45+ minutes.
- Large test suites: **NEVER CANCEL** - can take 5-20 minutes. Set timeout to 30+ minutes.

#### For Go Projects (when added)
- Initialize module: `go mod init [module-name]`
- Download dependencies: `go mod download` (typically 10-60 seconds)
- Build: `go build` (typically 5-30 seconds)
- Test: `go test ./...` (timing varies by test suite)
- Format: `go fmt ./...`
- Vet: `go vet ./...`

#### For Rust Projects (when added)
- Initialize: `cargo init` or `cargo new [project-name]`
- Build: `cargo build` (first build can take 5-15 minutes for large dependency trees)
- Build release: `cargo build --release` (typically 2x slower than debug builds)
- Test: `cargo test` (timing varies by test suite)
- Format: `cargo fmt`
- Lint: `cargo clippy`

**CRITICAL BUILD TIMINGS:**
- First `cargo build`: **NEVER CANCEL** - can take 15-45 minutes. Set timeout to 60+ minutes.
- `cargo build --release`: **NEVER CANCEL** - can take 20-60 minutes. Set timeout to 75+ minutes.

### Browser/Web Development (when applicable)
- Start dev server with live reload
- Open browser to `http://localhost:3000` (or configured port)
- Verify hot reloading works
- Test cross-browser compatibility if required

## Validation Requirements

### ALWAYS Run These Validation Steps
1. **Build Validation**: Always run the complete build process after making changes
2. **Test Validation**: Run the full test suite, not just affected tests
3. **Lint Validation**: Run all configured linters and formatters
4. **Functional Validation**: Test the actual application functionality, not just compilation

### Manual Testing Scenarios
When code is added to this repository, **ALWAYS** perform these validations:

#### For Web Applications
- Load the application in a browser
- Test primary user workflows (login, navigation, core features)
- Verify responsive design on different screen sizes
- Check browser console for errors
- Test form submissions and data persistence

#### For CLI Applications  
- Run `--help` command to verify usage information
- Test with sample inputs/files
- Verify output files are created correctly
- Test error handling with invalid inputs

#### For Libraries/APIs
- Import/require the library successfully  
- Call main API methods with valid inputs
- Test error handling with invalid inputs
- Verify return values match expected formats

### Performance Validation
- Measure and document actual timing for build/test commands
- Note any commands that take longer than expected
- Update timeout recommendations based on actual performance

## Git Workflow

### Before Making Changes
- `git --no-pager status` - Check current state
- `git pull origin main` - Get latest changes (if working on main)
- Create feature branch: `git checkout -b feature/[description]`

### After Making Changes
- `git --no-pager diff` - Review changes
- Stage changes: `git add .` (or specific files)
- Commit: `git commit -m "descriptive message"`
- Push: `git push origin [branch-name]`

### Code Quality Checks
ALWAYS run these before committing:
- All configured linters
- All configured formatters  
- Full test suite
- Build process (if applicable)

## Debugging Guidelines

### Common Issues and Solutions
- **Port conflicts**: Check for running processes on common ports (3000, 8080, 5000)
- **Permission errors**: Use `sudo` sparingly, prefer fixing file permissions
- **Dependency conflicts**: Clear caches (`npm cache clean --force`, `pip cache purge`)
- **Build failures**: Check for missing system dependencies

### Debugging Commands
- Check running processes: `ps aux | grep [process-name]`
- Check port usage: `netstat -tulpn | grep [port]`
- Check disk space: `df -h`
- Check memory usage: `free -h`

## Repository Structure Expectations

As the codebase grows, expect these typical structures:

### JavaScript/Node.js Projects
```
/
├── package.json          # Dependencies and scripts
├── package-lock.json     # Locked dependency versions
├── src/                  # Source code
├── dist/ or build/       # Built/compiled output
├── test/ or tests/       # Test files
├── public/               # Static assets (for web apps)
├── .eslintrc.*          # ESLint configuration
├── .prettierrc.*        # Prettier configuration
├── webpack.config.js    # Build configuration
└── node_modules/        # Dependencies (git-ignored)
```

### Python Projects
```
/
├── requirements.txt     # Dependencies
├── setup.py            # Package configuration
├── src/ or [package]/  # Source code
├── tests/              # Test files
├── venv/               # Virtual environment (git-ignored)
├── .pylintrc           # Pylint configuration
└── pytest.ini         # Pytest configuration
```

### Go Projects
```
/
├── go.mod              # Module definition
├── go.sum              # Dependency checksums
├── main.go             # Main application (or cmd/)
├── internal/           # Private packages
├── pkg/                # Public packages
└── vendor/             # Vendored dependencies (optional)
```

### Rust Projects
```
/
├── Cargo.toml          # Project configuration
├── Cargo.lock          # Locked dependencies
├── src/                # Source code
│   └── main.rs         # Main entry point
├── tests/              # Integration tests
└── target/             # Build output (git-ignored)
```

## Troubleshooting

### When Commands Fail
1. **Check prerequisites**: Ensure all required tools are installed
2. **Check permissions**: Verify file/directory permissions
3. **Check paths**: Ensure you're in the correct directory
4. **Check network**: Some failures may be due to network connectivity
5. **Check logs**: Look for detailed error messages in command output

### Performance Issues
- Monitor system resources during builds
- Consider using faster alternatives (e.g., `yarn` instead of `npm`)
- Use parallel processing where available
- Clear caches when encountering unexplained slowdowns

### Environment Issues
- Verify tool versions match project requirements
- Check environment variables are set correctly
- Ensure sufficient disk space for builds and dependencies
- Consider using containerization for complex environments

## Critical Reminders

- **NEVER CANCEL long-running builds or tests** - They may take 45+ minutes and that's normal
- **ALWAYS** set appropriate timeouts (60+ minutes for builds, 30+ minutes for tests)
- **ALWAYS** validate functionality manually, not just compilation
- **ALWAYS** run the complete test suite before committing
- **ALWAYS** check that your changes don't break existing functionality
- **ALWAYS** document any new requirements or setup steps you discover

## Time Expectations

Document actual timings as you encounter them:
- Simple npm install: 10-60 seconds
- Complex npm install: 5-15 minutes (**NEVER CANCEL**)
- JavaScript builds: 30 seconds to 5 minutes 
- Complex JavaScript builds: 5-15 minutes (**NEVER CANCEL**)
- Python pip install: 30 seconds to 5 minutes
- Data science pip install: 10-30 minutes (**NEVER CANCEL**)
- First Rust build: 15-45 minutes (**NEVER CANCEL**)
- Go builds: 5-30 seconds
- Test suites: 5 seconds to 20+ minutes (**NEVER CANCEL**)

Always add 50% buffer to these times when setting command timeouts.