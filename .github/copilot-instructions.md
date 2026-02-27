<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# Accounting Mailbox Reader - Project Guidelines

## Project Overview
This is the foundation phase of a Python CLI tool for reading, analyzing, and triaging emails from Planted's accounting mailbox. The system integrates with Microsoft Graph API and progressively adds AI-powered classification, risk assessment, and Business Central integration.

## Architecture
- **Phase 1 (Current)**: Email reading & attachment extraction
- **Phase 2**: AI classification with accounting-specific categories
- **Phase 3**: Business Central integration for autonomous processing
- **Phase 4**: Risk assessment and escalation engine
- **Phase 5**: Jira workflow automation

## Key Files
- `main.py` - CLI entry point (Click framework)
- `src/email_reader.py` - Main email reading logic
- `src/graph_client.py` - Microsoft Graph API wrapper
- `src/attachment_analyzer.py` - PDF, Excel, CSV extraction
- `src/output_formatter.py` - JSON, table, detailed text output
- `src/config.py` - Configuration management
- `config/settings.yaml` - Application settings
- `.env` - Azure credentials (DO NOT commit)

## Development Standards
- Use type hints for all functions
- Docstrings for classes and public methods
- Dataclasses for data models (@dataclass decorator)
- Single responsibility principle for components
- Comprehensive logging (logging module)

## Testing
- Unit tests in `tests/` directory
- Integration tests with mocked Graph API
- Manual testing against real mailbox before production

## Security
- All API tokens in environment variables
- `.env` file in `.gitignore`
- Read-only access to mailbox (no mutations)
- Full audit logging for all operations

## Commands (with Virtual Environment)
```
.venv\Scripts\python main.py config-show      # Show configuration
.venv\Scripts\python main.py preview           # Preview recent emails
.venv\Scripts\python main.py read --format json # Read all emails as JSON
.venv\Scripts\python main.py read --format detailed # Detailed output
```

Windows convenience: Use `run.bat` helper script
- `run.bat config` - Show configuration
- `run.bat preview` - Preview emails
- `run.bat read [options]` - Read emails

## Next Steps After Foundation
1. **Phase 1**: Add AccountingClassifier for email categorization
2. **Phase 2**: Implement RiskEngine for risk assessment
3. **Phase 3**: Create BusinessCentralClient for invoice lookup
4. **Phase 4**: Build autonomous processing workflows
5. **Phase 5**: Integrate Jira for ticket management

## For Contributors
- Keep changes focused and testable
- Update README.md for new features
- Add type hints and docstrings
- Run `pylint` or `flake8` before committing
- Test with both dry-run and real scenarios
