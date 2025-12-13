#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Detailed functionality tests"""
import sys

def test_encoding_edge_cases():
    """Test encoding with various edge cases"""
    print("\n=== Testing Encoding Edge Cases ===")
    from outlook_mcp_server.backend.utils import safe_encode_text
    
    test_cases = [
        ("normal string", "normal string"),
        (b"bytes string", "bytes string"),
        (None, ""),
        (123, "123"),
        (b"\xe4\xb8\xad\xe6\x96\x87", "中文"),  # UTF-8 Chinese
    ]
    
    for input_val, expected_type in test_cases:
        result = safe_encode_text(input_val, "test")
        assert isinstance(result, str), f"Failed for {input_val}"
        print(f"  [OK] {type(input_val).__name__} -> str")
    
    print("[OK] All encoding edge cases handled")
    return True

def test_dasl_filter_variations():
    """Test DASL filter with different configurations"""
    print("\n=== Testing DASL Filter Variations ===")
    from outlook_mcp_server.backend.utils import build_dasl_filter
    from datetime import datetime, timezone
    
    now = datetime.now(timezone.utc)
    
    # Test single term
    filter1 = build_dasl_filter(["test"], now, "subject", match_all=True)
    assert "test" in filter1
    assert "@SQL=" in filter1
    print("  [OK] Single term filter")
    
    # Test multiple terms with AND
    filter2 = build_dasl_filter(["test", "email"], now, "subject", match_all=True)
    assert "test" in filter2
    assert "email" in filter2
    assert "AND" in filter2
    print("  [OK] Multiple terms with AND")
    
    # Test multiple terms with OR
    filter3 = build_dasl_filter(["test", "email"], now, "subject", match_all=False)
    assert "test" in filter3
    assert "email" in filter3
    assert "OR" in filter3
    print("  [OK] Multiple terms with OR")
    
    # Test different fields
    for field in ["subject", "sender", "recipient", "body"]:
        filter_field = build_dasl_filter(["test"], now, field, match_all=True)
        assert "@SQL=" in filter_field
        print(f"  [OK] Field: {field}")
    
    print("[OK] All DASL filter variations work")
    return True

def test_pagination_edge_cases():
    """Test pagination with edge cases"""
    print("\n=== Testing Pagination Edge Cases ===")
    from outlook_mcp_server.backend.utils import get_pagination_info
    
    # Test normal case
    info = get_pagination_info(47, 5)
    assert info['total_pages'] == 10
    assert info['total_items'] == 47
    print("  [OK] Normal pagination: 47 items, 5 per page = 10 pages")
    
    # Test exact division
    info = get_pagination_info(50, 5)
    assert info['total_pages'] == 10
    print("  [OK] Exact division: 50 items, 5 per page = 10 pages")
    
    # Test empty
    info = get_pagination_info(0, 5)
    assert info['total_pages'] == 0
    print("  [OK] Empty: 0 items = 0 pages")
    
    # Test single item
    info = get_pagination_info(1, 5)
    assert info['total_pages'] == 1
    print("  [OK] Single item: 1 item = 1 page")
    
    print("[OK] All pagination edge cases handled")
    return True

def test_email_validation():
    """Test email validation with various formats"""
    print("\n=== Testing Email Validation ===")
    from outlook_mcp_server.backend.utils import validate_email_address
    
    valid_emails = [
        "test@example.com",
        "user.name@domain.co.uk",
        "user+tag@example.com",
        "user_name@example.com",
        "123@example.com",
    ]
    
    invalid_emails = [
        "invalid",
        "@example.com",
        "user@",
        "user name@example.com",
        "user@domain",
        "",
    ]
    
    for email in valid_emails:
        assert validate_email_address(email) == True, f"Should be valid: {email}"
        print(f"  [OK] Valid: {email}")
    
    for email in invalid_emails:
        assert validate_email_address(email) == False, f"Should be invalid: {email}"
        print(f"  [OK] Invalid: {email}")
    
    print("[OK] Email validation works correctly")
    return True

def test_pydantic_validation():
    """Test Pydantic validation edge cases"""
    print("\n=== Testing Pydantic Validation ===")
    from outlook_mcp_server.backend.validators import (
        EmailSearchParams,
        PaginationParams
    )
    
    # Test auto-correction
    params = EmailSearchParams(search_term="  test  ", days=7)
    assert params.search_term == "test", "Should strip whitespace"
    print("  [OK] Whitespace trimming works")
    
    # Test min/max validation
    try:
        PaginationParams(page=0)  # Below minimum
        assert False, "Should have raised error"
    except ValueError:
        print("  [OK] Min page validation works")
    
    try:
        PaginationParams(page=1, per_page=100)  # Above maximum
        assert False, "Should have raised error"
    except ValueError:
        print("  [OK] Max per_page validation works")
    
    # Test type coercion
    params = PaginationParams(page="1", per_page="5")  # Strings to ints
    assert params.page == 1
    assert params.per_page == 5
    print("  [OK] Type coercion works")
    
    print("[OK] Pydantic validation comprehensive")
    return True

def test_enums():
    """Test enum values match expected Outlook constants"""
    print("\n=== Testing Enum Values ===")
    from outlook_mcp_server.backend.utils import OutlookFolderType, OutlookItemClass
    
    # Test folder types
    assert OutlookFolderType.DELETED_ITEMS == 3
    assert OutlookFolderType.OUTBOX == 4
    assert OutlookFolderType.SENT_MAIL == 5
    assert OutlookFolderType.INBOX == 6
    assert OutlookFolderType.CALENDAR == 9
    assert OutlookFolderType.CONTACTS == 10
    assert OutlookFolderType.TASKS == 13
    assert OutlookFolderType.DRAFTS == 16
    print("  [OK] All folder type enums correct")
    
    # Test item classes
    assert OutlookItemClass.MAIL_ITEM == 43
    assert OutlookItemClass.APPOINTMENT_ITEM == 26
    assert OutlookItemClass.CONTACT_ITEM == 40
    assert OutlookItemClass.TASK_ITEM == 48
    print("  [OK] All item class enums correct")
    
    print("[OK] All enum values match Outlook constants")
    return True

def main():
    """Run detailed tests"""
    print("="*60)
    print("DETAILED FUNCTIONALITY TESTS")
    print("="*60)
    
    tests = [
        ("Encoding Edge Cases", test_encoding_edge_cases),
        ("DASL Filter Variations", test_dasl_filter_variations),
        ("Pagination Edge Cases", test_pagination_edge_cases),
        ("Email Validation", test_email_validation),
        ("Pydantic Validation", test_pydantic_validation),
        ("Enum Values", test_enums),
    ]
    
    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"\n[X] {name} FAILED: {e}")
            import traceback
            traceback.print_exc()
            results.append((name, False))
    
    # Summary
    print("\n" + "="*60)
    print("DETAILED TEST SUMMARY")
    print("="*60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "PASS" if result else "FAIL"
        symbol = "[OK]" if result else "[X]"
        print(f"{symbol} {name}: {status}")
    
    print("\n" + "="*60)
    print(f"RESULTS: {passed}/{total} detailed tests passed")
    print("="*60)
    
    return 0 if passed == total else 1

if __name__ == "__main__":
    sys.exit(main())