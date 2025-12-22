# Project Structure (Corrected – 2025)

This document provides a concise and accurate overview of the current Outlook MCP Server project structure based on the actual repository contents.

## Root Directory

- **cli_interface.py** – Text‑based CLI for interacting with the MCP server.
- **README.md** – Main setup and usage instructions.
- **pyproject.toml** – Project configuration and dependencies.
- **requirements.txt** – Python dependency list.
- **mcp-config-*.json** – MCP launch configuration files.
- **check_email_dates.py** – Standalone debug/check script.

## docs/ — Technical Documentation

Contains advanced documentation for specific areas of the backend:

- **email-loading-mutual-exclusion.md** – Notes on safe concurrent email loading.
- **email-ordering-optimization.md** – Explains ordering fixes for Outlook items.
- **performance-optimization.md** – (Primary performance doc, maintained separately)
- **server-side-search.md** – Details of server‑side filtering using Outlook’s Restrict API.
- **win32com-api-basics.md** – Introductory overview of Outlook COM automation.
- **win32com_api_implementation.md** – Deep dive into low‑level COM operations.

## outlook_mcp_server/ — Main Application Package

### Top-Level Package Files
- **__init__.py** – Exposes MCP tool entry points (search, view, list, move, reply, etc.).
- **__main__.py** – Allows running the package as a module.

---

## backend/ — Core Backend Logic

This is the heart of the system, implementing email extraction, listing, cache handling, and session management.

### Key Backend Modules
- **batch_operations.py** – Backend batch forward/multi‑action logic.
- **email_composition.py** – Sending, composing, and replying to emails.
- **email_data_extractor.py** – Extracts fields from Outlook MailItem objects.
- **email_metadata.py** – Metadata handling utilities.
- **email_utils.py** – Common Outlook email helpers.
- **shared.py** – Disk+memory email cache system.
- **utils.py** – General helpers (filters, DASL building, retry decorator, etc.).
- **validators.py** – Pydantic validation models for tool inputs.

---

## backend/email_search/ — Modular Search Pipeline

Implements specialized search logic and optimized email listing techniques.

- **body_search.py** – Body text searching.
- **recipient_search.py** – Search by recipient.
- **sender_search.py** – Search by sender.
- **subject_search.py** – Subject‑based searching.
- **email_listing.py** – Optimized listing using Restrict, sorting, and iteration.
- **search_common.py** – Shared extraction and DASL helpers.
- **parallel_extractor.py** – ThreadPool‑based parallel extraction logic.
- **server_search.py** – Server‑side filtering helpers.
- **unified_search.py** – Unified entry point for consolidated search logic.

---

## backend/outlook_session/ — Outlook COM Session Layer

Encapsulates connection logic, folder lookups, and COM exception handling.

- **session_manager.py** – Main session orchestration.
- **folder_operations.py** – Folder navigation helpers.
- **email_operations.py** – Common email manipulation operations.
- **decorators.py** – Session‑safe decorators.
- **exceptions.py** – Custom session‑related errors.
- **utils.py** – COM utilities and helpers.

---

## tools/ — MCP Tool Implementations

These files expose backend functionality as MCP‑compatible tools.

- **email_operations.py** – View, delete, reply, get-by-number tools.
- **batch_operations.py** – Batch‑related tools.
- **folder_tools.py** – Folder create/remove/list.
- **search_tools.py** – All search tools.
- **viewing_tools.py** – Email viewing / cache viewing tools.
- **registration.py** – Registers all MCP tools.

---

## tests/ — Unit Tests

Current repository includes only **unit tests**:

### tests/unit/
- **test_direct_search.py**
- **test_early_termination_fix.py**
- **test_final_search.py**
- **test_search.py**
- **check_cache.py**

### tests/integration/**
(Currently empty)

### tests/scripts/**
(Currently empty)

---

## Summary

The project is structured around three main pillars:

1. **Backend** – Real Outlook COM handling, search, extraction, and cache logic  
2. **Tools** – MCP‑exposed API for interacting with the backend  
3. **Documentation** – Technical references on performance and COM behavior  

This corrected project structure reflects the current working state of the repository as of December 2025.