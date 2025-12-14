#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Test script to verify all improvements work correctly"""
import sys
import os
import logging

# Fix Windows console encoding
if sys.platform == 'win32':
    os.system('chcp 65001 > nul')
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def test_utils():
    """Test utils.py module"""
    print("\n=== Testing utils.py ===")
    try:
        from outlook_mcp_server.backend.utils import (
            OutlookFolderType,
            OutlookItemClass,
            safe_encode_text,
            validate_email_address,
            build_dasl_filter,
            get_pagination_info
        )
        
        # Test enums
        assert OutlookFolderType.INBOX == 6
        assert OutlookItemClass.MAIL_ITEM == 43
        print("[OK] Enums work correctly")
        
        # Test encoding
        result = safe_encode_text("test string", "test_field")
        assert result == "test string"
        result = safe_encode_text(b"test bytes", "test_field")
        assert isinstance(result, str)
        print("[OK] safe_encode_text works")
        
        # Test email validation
        assert validate_email_address("test@example.com") == True
        assert validate_email_address("invalid") == False
        print("[OK] validate_email_address works")
        
        # Test pagination
        info = get_pagination_info(47, 5)
        assert info['total_pages'] == 10
        assert info['total_items'] == 47
        print("[OK] get_pagination_info works")
        
        # Test DASL filter
        from datetime import datetime, timezone
        filter_str = build_dasl_filter(
            ['test', 'email'],
            datetime.now(timezone.utc),
            'subject',
            match_all=True
        )
        assert '@SQL=' in filter_str
        assert 'test' in filter_str
        print("[OK] build_dasl_filter works")
        
        print("[OK][OK][OK] utils.py - ALL TESTS PASSED")
        return True
    except Exception as e:
        print(f"[X] utils.py test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_validators():
    """Test validators.py module"""
    print("\n=== Testing validators.py ===")
    try:
        from outlook_mcp_server.backend.validators import (
            EmailSearchParams,
            EmailListParams,
            PaginationParams
        )
        
        # Test EmailSearchParams
        params = EmailSearchParams(
            search_term="test",
            days=7,
            folder_name="Inbox",
            match_all=True
        )
        assert params.search_term == "test"
        assert params.days == 7
        print("[OK] EmailSearchParams works")
        
        # Test validation error
        try:
            EmailSearchParams(search_term="", days=7)  # Empty string
            print("[X] Should have raised validation error")
            return False
        except ValueError:
            print("[OK] Validation catches empty search term")
        
        # Test days range validation
        try:
            EmailSearchParams(search_term="test", days=100)  # Over limit
            print("[X] Should have raised validation error for days")
            return False
        except ValueError:
            print("[OK] Validation catches invalid days range")
        
        # Test EmailListParams
        params = EmailListParams(days=7, folder_name="Inbox")
        assert params.days == 7
        print("[OK] EmailListParams works")
        
        # Test PaginationParams
        params = PaginationParams(page=1, per_page=5)
        assert params.page == 1
        assert params.per_page == 5
        print("[OK] PaginationParams works")
        
        print("[OK][OK][OK] validators.py - ALL TESTS PASSED")
        return True
    except Exception as e:
        print(f"[X] validators.py test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_outlook_session():
    """Test outlook_session.py improvements"""
    print("\n=== Testing outlook_session.py ===")
    try:
        from outlook_mcp_server.backend.outlook_session import OutlookSessionManager
        from outlook_mcp_server.backend.utils import OutlookFolderType
        
        # Test that it imports correctly
        print("[OK] OutlookSessionManager imports correctly")
        
        # Test that enums are used (check if the method exists)
        assert hasattr(OutlookSessionManager, '_connect')
        assert hasattr(OutlookSessionManager, '_disconnect')
        assert hasattr(OutlookSessionManager, '_cleanup_partial_connection')
        print("[OK] New methods exist")
        
        print("[OK][OK][OK] outlook_session.py - STRUCTURE TESTS PASSED")
        print("Note: Actual Outlook connection requires Outlook to be installed")
        return True
    except Exception as e:
        print(f"[X] outlook_session.py test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_email_retrieval():
    """Test email_retrieval.py improvements"""
    print("\n=== Testing email_retrieval.py ===")
    try:
        from outlook_mcp_server.backend.email_retrieval import (
            search_email_by_subject,
            search_email_by_from,
            search_email_by_to,
            search_email_by_body,
            _unified_search
        )
        
        # Test that unified search function exists
        print("[OK] _unified_search function exists")
        
        # Test that all search functions exist
        print("[OK] All search functions import correctly")
        
        print("[OK][OK][OK] email_retrieval.py - STRUCTURE TESTS PASSED")
        print("Note: Actual search requires Outlook connection")
        return True
    except Exception as e:
        print(f"[X] email_retrieval.py test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_email_composition():
    """Test email_composition.py improvements"""
    print("\n=== Testing email_composition.py ===")
    try:
        from outlook_mcp_server.backend.email_composition import (
            reply_to_email_by_number,
            compose_email
        )
        
        print("[OK] Functions import correctly")
        
        print("[OK][OK][OK] email_composition.py - STRUCTURE TESTS PASSED")
        print("Note: Actual email operations require Outlook connection")
        return True
    except Exception as e:
        print(f"[X] email_composition.py test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_batch_operations():
    """Test batch_operations.py improvements"""
    print("\n=== Testing batch_operations.py ===")
    try:
        from outlook_mcp_server.backend.batch_operations import batch_forward_emails
        print("[OK] batch_forward_emails imports correctly")
        
        print("[OK][OK][OK] batch_operations.py - STRUCTURE TESTS PASSED")
        return True
    except Exception as e:
        print(f"[X] batch_operations.py test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Run all tests"""
    print("="*60)
    print("TESTING ALL IMPROVEMENTS")
    print("="*60)
    
    results = []
    
    # Run tests
    results.append(("utils.py", test_utils()))
    results.append(("validators.py", test_validators()))
    results.append(("outlook_session.py", test_outlook_session()))
    results.append(("email_retrieval.py", test_email_retrieval()))
    results.append(("email_composition.py", test_email_composition()))
    results.append(("batch_operations.py", test_batch_operations()))
    
    # Summary
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for module, result in results:
        status = "PASS" if result else "FAIL"
        symbol = "[OK]" if result else "[X]"
        print(f"{symbol} {module}: {status}")
    
    print("\n" + "="*60)
    print(f"RESULTS: {passed}/{total} tests passed")
    print("="*60)
    
    if passed == total:
        print("\n ALL TESTS PASSED! ")
        return 0
    else:
        print(f"\n[!] {total - passed} test(s) failed")
        return 1

if __name__ == "__main__":
    sys.exit(main())
