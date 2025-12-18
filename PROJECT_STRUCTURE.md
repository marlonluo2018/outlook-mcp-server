# Project Structure Documentation

This document provides a comprehensive overview of the Outlook MCP Server project structure and explains the purpose of each file and directory.

## Root Directory

### Configuration Files
- **`pyproject.toml`** - Project configuration file defining dependencies, build settings, and project metadata
- **`requirements.txt`** - List of Python package dependencies required for the project
- **`.gitignore`** - Git configuration file specifying files and directories to exclude from version control

### Entry Points
- **`cli_interface.py`** - Command-line interface entry point for the Outlook MCP Server
- **`mcp-config-direct.json`** - Direct execution configuration for MCP server
- **`mcp-config-python.json`** - Python-based execution configuration for MCP server
- **`mcp-config-uvx.json`** - UVX-based execution configuration for MCP server

### Documentation
- **`README.md`** - Main project documentation and setup instructions
- **`PROJECT_STRUCTURE.md`** - This file - comprehensive project structure documentation

## Directory Structure

### `docs/` - Technical Documentation
Contains detailed technical documentation for various aspects of the project:
- **`performance-optimization.md`** - Performance optimization strategies and implementation details
- **`server-side-search.md`** - Documentation on server-side search implementation
- **`win32com-api-basics.md`** - Basic Win32COM API usage and concepts
- **`win32com_api_implementation.md`** - Detailed Win32COM API implementation guide

### `dev_scripts/` - Development Utilities
Contains scripts and tools used during development:
- **`organize_files.ps1`** - PowerShell script for organizing project files (used during project cleanup)
- **`profile_performance.py`** - Performance profiling script for analyzing code execution

### `outlook_mcp_server/` - Main Package
The core package containing all MCP server functionality.

#### `backend/` - Core Backend Functionality
- **`__init__.py`** - Package initialization file
- **`batch_operations.py`** - Batch email operations (forward, delete, move in bulk)
- **`email_composition.py`** - Email composition and sending functionality
- **`email_data_extractor.py`** - Extracts data from email objects
- **`email_metadata.py`** - Handles email metadata operations
- **`email_search.py`** - Main email search functionality (consolidated optimized version)
- **`email_utils.py`** - Utility functions for email operations
- **`outlook_session.py`** - Outlook session management
- **`shared.py`** - Shared constants and configurations
- **`utils.py`** - General utility functions
- **`validators.py`** - Input validation functions

##### `backend/email_search/` - Modular Search Components
- **`__init__.py`** - Package initialization
- **`body_search.py`** - Email body content search functionality
- **`recipient_search.py`** - Email recipient search functionality
- **`search_utils.py`** - Common search utilities and helpers
- **`sender_search.py`** - Email sender search functionality
- **`subject_search.py`** - Email subject search functionality

##### `backend/outlook_session/` - Session Management
- **`__init__.py`** - Package initialization
- **`decorators.py`** - Session management decorators
- **`email_operations.py`** - Core email operations within sessions
- **`exceptions.py`** - Custom exception definitions
- **`folder_operations.py`** - Outlook folder operations
- **`session_manager.py`** - Main session management logic
- **`utils.py`** - Session-related utilities

#### `tools/` - MCP Tool Implementations
- **`__init__.py`** - Package initialization
- **`batch_operations.py`** - MCP tools for batch email operations
- **`email_operations.py`** - MCP tools for individual email operations
- **`folder_tools.py`** - MCP tools for folder management

- **`registration.py`** - Tool registration and discovery
- **`search_tools.py`** - MCP tools for email searching
- **`viewing_tools.py`** - MCP tools for email viewing and retrieval

### `tests/` - Test Suite
Comprehensive test suite organized by test type:

#### `tests/unit/` - Unit Tests
Tests for individual components and functions:
- **`test_corrected_criteria.py`** - Tests for search criteria correction
- **`test_direct_search.py`** - Tests for direct search functionality
- **`test_early_termination_fix.py`** - Tests for early termination fixes
- **`test_final_search.py`** - Tests for final search implementation
- **`test_search.py`** - General search functionality tests
- **`test_search_fixed.py`** - Tests for fixed search functionality
- **`test_search_formats.py`** - Tests for different search format handling
- **`test_search_terms.py`** - Tests for search term processing

#### `tests/integration/` - Integration Tests
Tests for integrated system functionality:
- **`test_dynamic_limit.py`** - Tests for dynamic result limiting
- **`test_dynamic_limit_fixed.py`** - Tests for fixed dynamic limiting
- **`test_optimized_list.py`** - Tests for optimized email listing
- **`test_performance_comparison.py`** - Performance comparison tests
- **`test_search_days.py`** - Tests for date-based search functionality

#### `tests/scripts/` - Development and Analysis Scripts
Utility scripts for debugging and analysis:
- **`analyze_distribution.py`** - Analyzes email distribution patterns
- **`analyze_email_distribution.py`** - Detailed email distribution analysis
- **`check_1000th_email.py`** - Checks the 1000th email in collection
- **`check_approval_emails.py`** - Checks for approval-related emails
- **`check_emails.py`** - General email checking utility
- **`check_total_emails.py`** - Counts total emails
- **`debug_date_filtering.py`** - Debugs date filtering functionality
- **`debug_dates.py`** - General date debugging
- **`debug_dates_full.py`** - Comprehensive date debugging
- **`debug_search_criteria.py`** - Debugs search criteria processing
- **`find_7day_cutoff.py`** - Finds 7-day email cutoff point

## Key Architectural Decisions

### Backend Organization
The backend is organized into modular components:
- **Search functionality** is split into specialized modules (subject, body, sender, recipient)
- **Session management** is encapsulated in its own package with proper error handling
- **Email operations** are separated from search operations for better maintainability

### Testing Strategy
- **Unit tests** focus on individual function behavior
- **Integration tests** verify system-level functionality
- **Development scripts** provide debugging and analysis capabilities

### Configuration Management
Multiple configuration files support different deployment scenarios:
- Direct Python execution
- UVX-based execution
- Direct server execution

This structure ensures the project is maintainable, testable, and follows Python best practices for package organization.