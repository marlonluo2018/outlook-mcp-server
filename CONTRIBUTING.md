# 🤝 Contributing to Outlook MCP Server

Thank you for your interest in contributing to Outlook MCP Server! We welcome contributions from everyone.

## 🎯 Ways to Contribute

### 1. 🐛 Report Bugs
Found a bug? Please check if it's already reported in [Issues](https://github.com/marlonluo2018/outlook-mcp-server/issues). If not, create a new issue using the bug report template.

### 2. 💡 Suggest Features
Have an idea for improvement? Check the [Issues](https://github.com/marlonluo2018/outlook-mcp-server/issues) first, then create a feature request using the template.

### 3. 🔧 Code Contributions
Want to fix a bug or add a feature? Follow these steps:

#### Development Setup
```bash
# Fork and clone the repository
git clone https://github.com/YOUR_USERNAME/outlook-mcp-server.git
cd outlook-mcp-server

# Install in development mode
pip install -e ".[dev]"

# Run tests to ensure everything works
pytest tests/ -v
```

#### Code Style
We use several tools to maintain code quality:
- **Black** for code formatting
- **Flake8** for linting
- **MyPy** for type checking

```bash
# Format code
black outlook_mcp_server/

# Check linting
flake8 outlook_mcp_server/

# Check types
mypy outlook_mcp_server/
```

#### Testing
- Write tests for new functionality
- Ensure all tests pass before submitting
- Include both unit and integration tests when appropriate

### 4. 📚 Documentation
Help improve our documentation:
- Fix typos or unclear explanations
- Add examples and tutorials
- Improve API documentation
- Translate documentation

### 5. 🎨 Design & UX
- Suggest UI/UX improvements
- Create mockups or prototypes
- Improve user workflows

## 📋 Pull Request Process

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/amazing-feature`
3. **Make your changes**
4. **Run tests**: `pytest tests/ -v`
5. **Check code quality**: Run Black, Flake8, and MyPy
6. **Commit your changes**: Use descriptive commit messages
7. **Push to your branch**: `git push origin feature/amazing-feature`
8. **Open a Pull Request**

### Pull Request Guidelines
- Use the PR template provided
- Link to related issues
- Include tests for new functionality
- Update documentation as needed
- Keep changes focused and atomic

## 🏷️ Issue Labels

We use labels to categorize issues:
- `bug` - Something isn't working
- `enhancement` - New feature or improvement
- `documentation` - Documentation improvements
- `good first issue` - Good for newcomers
- `help wanted` - Extra attention needed
- `question` - Further information is requested

## 🎯 Good First Issues

New to the project? Look for issues labeled `good first issue`. These are specially curated for newcomers.

## 📞 Communication

- **GitHub Issues**: For bug reports and feature requests
- **GitHub Discussions**: For questions, ideas, and community discussions
- **Pull Requests**: For code contributions

## 🔒 Security Issues

Please report security issues privately to the maintainers. Do not create public issues for security vulnerabilities.

## 📜 Code of Conduct

We expect all contributors to adhere to our Code of Conduct. Please be respectful and inclusive in all interactions.

## 🙏 Thank You!

Your contributions are what make this project great. Thank you for helping improve Outlook MCP Server for everyone! 🎉