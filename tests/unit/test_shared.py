import pytest
from datetime import datetime, timezone, timedelta
from unittest.mock import patch, MagicMock
from outlook_mcp_server.backend.shared import (
    email_cache,
    email_cache_order,
    _email_time_cache,
    add_email_to_cache,
    get_email_from_cache,
    clear_cache,
    get_cache_size,
    get_cache_stats,
    cleanup_cache,
    get_emails_by_date_range,
    get_emails_by_sender,
    get_emails_by_subject,
    get_emails_by_date_range_cached,
    get_emails_by_sender_cached,
    get_emails_by_subject_cached
)


class TestCacheManagement:
    """Test suite for cache management functions."""

    def setup_method(self):
        """Setup method to clear cache before each test."""
        clear_cache()

    def teardown_method(self):
        """Teardown method to clear cache after each test."""
        clear_cache()

    def test_add_email_to_cache(self):
        """Test adding an email to cache."""
        email_data = {
            "subject": "Test Subject",
            "sender": "test@example.com",
            "received_time": datetime.now(timezone.utc).isoformat(),
            "body": "Test body"
        }
        email_id = "test_id_1"
        
        add_email_to_cache(email_id, email_data)
        
        assert email_id in email_cache
        assert email_cache[email_id] == email_data
        assert email_id in email_cache_order

    def test_add_email_to_cache_eviction(self):
        """Test cache eviction when max size is exceeded."""
        from outlook_mcp_server.backend.config import CacheConfig
        
        for i in range(CacheConfig.MAX_EMAILS + 10):
            email_data = {
                "subject": f"Test Subject {i}",
                "sender": f"test{i}@example.com",
                "received_time": datetime.now(timezone.utc).isoformat(),
                "body": f"Test body {i}"
            }
            add_email_to_cache(f"test_id_{i}", email_data)
        
        assert len(email_cache) <= CacheConfig.MAX_EMAILS
        assert len(email_cache_order) <= CacheConfig.MAX_EMAILS

    def test_add_email_to_cache_time_cleanup(self):
        """Test that time cache is cleaned up when email is evicted."""
        from outlook_mcp_server.backend.config import CacheConfig
        
        received_time = datetime.now(timezone.utc).isoformat()
        
        for i in range(CacheConfig.MAX_EMAILS + 10):
            email_data = {
                "subject": f"Test Subject {i}",
                "sender": f"test{i}@example.com",
                "received_time": received_time,
                "body": f"Test body {i}"
            }
            add_email_to_cache(f"test_id_{i}", email_data)
        
        assert len(_email_time_cache) <= CacheConfig.MAX_EMAILS

    def test_get_email_from_cache_exists(self):
        """Test retrieving an existing email from cache."""
        email_data = {
            "subject": "Test Subject",
            "sender": "test@example.com",
            "received_time": datetime.now(timezone.utc).isoformat(),
            "body": "Test body"
        }
        email_id = "test_id_1"
        
        add_email_to_cache(email_id, email_data)
        result = get_email_from_cache(email_id)
        
        assert result == email_data

    def test_get_email_from_cache_not_exists(self):
        """Test retrieving a non-existent email from cache."""
        result = get_email_from_cache("non_existent_id")
        
        assert result is None

    def test_clear_cache(self):
        """Test clearing the cache."""
        email_data = {
            "subject": "Test Subject",
            "sender": "test@example.com",
            "received_time": datetime.now(timezone.utc).isoformat(),
            "body": "Test body"
        }
        
        add_email_to_cache("test_id_1", email_data)
        clear_cache()
        
        assert len(email_cache) == 0
        assert len(email_cache_order) == 0
        assert len(_email_time_cache) == 0

    def test_get_cache_size(self):
        """Test getting cache size."""
        assert get_cache_size() == 0
        
        email_data = {
            "subject": "Test Subject",
            "sender": "test@example.com",
            "received_time": datetime.now(timezone.utc).isoformat(),
            "body": "Test body"
        }
        
        add_email_to_cache("test_id_1", email_data)
        assert get_cache_size() == 1

    def test_get_cache_stats(self):
        """Test getting cache statistics."""
        email_data = {
            "subject": "Test Subject",
            "sender": "test@example.com",
            "received_time": datetime.now(timezone.utc).isoformat(),
            "body": "Test body"
        }
        
        add_email_to_cache("test_id_1", email_data)
        stats = get_cache_stats()
        
        assert stats["total_emails"] == 1
        assert "oldest_email" in stats
        assert "newest_email" in stats

    def test_cleanup_cache(self):
        """Test cache cleanup functionality."""
        from outlook_mcp_server.backend.config import CacheConfig
        
        old_time = (datetime.now(timezone.utc) - timedelta(days=CacheConfig.CACHE_EXPIRY_HOURS + 1)).isoformat()
        new_time = datetime.now(timezone.utc).isoformat()
        
        add_email_to_cache("old_id", {"subject": "Old", "received_time": old_time})
        add_email_to_cache("new_id", {"subject": "New", "received_time": new_time})
        
        cleanup_cache()
        
        assert "old_id" not in email_cache
        assert "new_id" in email_cache


class TestCacheQueries:
    """Test suite for cache query functions."""

    def setup_method(self):
        """Setup method to populate cache with test data."""
        clear_cache()
        now = datetime.now(timezone.utc)
        
        test_emails = [
            {
                "id": "email_1",
                "subject": "Meeting Tomorrow",
                "sender": "alice@example.com",
                "received_time": (now - timedelta(hours=1)).isoformat(),
                "body": "Meeting at 10am"
            },
            {
                "id": "email_2",
                "subject": "Project Update",
                "sender": "bob@example.com",
                "received_time": (now - timedelta(hours=2)).isoformat(),
                "body": "Project status update"
            },
            {
                "id": "email_3",
                "subject": "Meeting Tomorrow",
                "sender": "charlie@example.com",
                "received_time": (now - timedelta(hours=3)).isoformat(),
                "body": "Another meeting"
            }
        ]
        
        for email in test_emails:
            add_email_to_cache(email["id"], email)

    def teardown_method(self):
        """Teardown method to clear cache after each test."""
        clear_cache()

    def test_get_emails_by_date_range(self):
        """Test getting emails by date range."""
        now = datetime.now(timezone.utc)
        start_time = (now - timedelta(hours=2)).isoformat()
        end_time = now.isoformat()
        
        results = get_emails_by_date_range(start_time, end_time)
        
        assert len(results) == 2

    def test_get_emails_by_sender(self):
        """Test getting emails by sender."""
        results = get_emails_by_sender("alice@example.com")
        
        assert len(results) == 1
        assert results[0]["sender"] == "alice@example.com"

    def test_get_emails_by_subject(self):
        """Test getting emails by subject."""
        results = get_emails_by_subject("Meeting")
        
        assert len(results) == 2

    def test_get_emails_by_date_range_cached(self):
        """Test getting emails by date range with caching."""
        now = datetime.now(timezone.utc)
        start_time = (now - timedelta(hours=2)).isoformat()
        end_time = now.isoformat()
        
        results1 = get_emails_by_date_range_cached(start_time, end_time)
        results2 = get_emails_by_date_range_cached(start_time, end_time)
        
        assert len(results1) == 2
        assert len(results2) == 2

    def test_get_emails_by_sender_cached(self):
        """Test getting emails by sender with caching."""
        results1 = get_emails_by_sender_cached("alice@example.com")
        results2 = get_emails_by_sender_cached("alice@example.com")
        
        assert len(results1) == 1
        assert len(results2) == 1

    def test_get_emails_by_subject_cached(self):
        """Test getting emails by subject with caching."""
        results1 = get_emails_by_subject_cached("Meeting")
        results2 = get_emails_by_subject_cached("Meeting")
        
        assert len(results1) == 2
        assert len(results2) == 2


class TestCacheEdgeCases:
    """Test suite for cache edge cases."""

    def setup_method(self):
        """Setup method to clear cache before each test."""
        clear_cache()

    def teardown_method(self):
        """Teardown method to clear cache after each test."""
        clear_cache()

    def test_add_email_with_missing_fields(self):
        """Test adding email with missing fields."""
        email_data = {
            "subject": "Test Subject"
        }
        
        add_email_to_cache("test_id", email_data)
        
        assert "test_id" in email_cache

    def test_get_emails_by_empty_sender(self):
        """Test getting emails with empty sender filter."""
        results = get_emails_by_sender("")
        
        assert results == []

    def test_get_emails_by_empty_subject(self):
        """Test getting emails with empty subject filter."""
        results = get_emails_by_subject("")
        
        assert results == []

    def test_get_emails_by_invalid_date_range(self):
        """Test getting emails with invalid date range."""
        now = datetime.now(timezone.utc)
        start_time = now.isoformat()
        end_time = (now - timedelta(hours=1)).isoformat()
        
        results = get_emails_by_date_range(start_time, end_time)
        
        assert results == []

    def test_cache_stats_empty(self):
        """Test getting cache stats when cache is empty."""
        stats = get_cache_stats()
        
        assert stats["total_emails"] == 0
