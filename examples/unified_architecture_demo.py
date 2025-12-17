"""Example usage of the unified email retrieval architecture."""

from outlook_mcp_server.backend.email_retrieval_unified import (
    get_email_by_number_unified,
    EmailRetrievalMode,
    format_email_with_media
)
from outlook_mcp_server.backend.email_tools_unified import (
    get_email_by_number_tool_unified,
    get_email_tool_legacy_wrapper,
    get_email_with_media_tool_legacy_wrapper
)


def demonstrate_unified_email_retrieval():
    """Demonstrate different modes of the unified email retrieval system."""
    
    print("=== Unified Email Retrieval Architecture Demo ===\n")
    
    # Example 1: Basic mode (fast, backward compatible)
    print("1. Basic Mode (Fast, Backward Compatible):")
    print("   Use case: Quick email preview, minimal resource usage")
    print("   Code: get_email_by_number_unified(1, mode='basic')")
    print("   Features:")
    print("   - Basic email metadata (subject, sender, date)")
    print("   - Simple body content")
    print("   - Basic attachment info (names, sizes)")
    print("   - Fast performance")
    print()
    
    # Example 2: Enhanced mode (full media support)
    print("2. Enhanced Mode (Full Media Support):")
    print("   Use case: Complete email analysis with media content")
    print("   Code: get_email_by_number_unified(1, mode='enhanced', include_attachments=True, embed_images=True)")
    print("   Features:")
    print("   - All basic features")
    print("   - Base64-encoded attachment content")
    print("   - Inline image embedding in HTML")
    print("   - Enhanced metadata (importance, sensitivity, categories)")
    print("   - Conversation threading info")
    print("   - Text file previews")
    print()
    
    # Example 3: Lazy mode (performance optimized)
    print("3. Lazy Mode (Performance Optimized):")
    print("   Use case: Balanced performance with enhanced features when needed")
    print("   Code: get_email_by_number_unified(1, mode='lazy')")
    print("   Features:")
    print("   - Uses cached data when available")
    print("   - Falls back to enhanced mode for missing data")
    print("   - Optimal for large email sets")
    print()


def demonstrate_mcp_tool_usage():
    """Demonstrate MCP tool usage with unified architecture."""
    
    print("=== MCP Tool Usage Examples ===\n")
    
    # Example 1: Basic email retrieval
    print("1. Basic Email Retrieval (MCP Tool):")
    print("   Tool: get_email_by_number_tool_unified")
    print("   Parameters:")
    print("   - email_number: 1")
    print("   - mode: 'basic'")
    print("   Response: Formatted text with basic email info")
    print()
    
    # Example 2: Enhanced email with media
    print("2. Enhanced Email with Media (MCP Tool):")
    print("   Tool: get_email_by_number_tool_unified")
    print("   Parameters:")
    print("   - email_number: 1")
    print("   - mode: 'enhanced'")
    print("   - include_attachments: True")
    print("   - embed_images: True")
    print("   Response: Formatted text with media content")
    print()
    
    # Example 3: Backward compatibility
    print("3. Backward Compatibility (Legacy Tools):")
    print("   Legacy tools still work via wrappers:")
    print("   - get_email_by_number_tool(1) → calls unified with mode='basic'")
    print("   - get_email_with_media_tool(1) → calls unified with mode='enhanced'")
    print("   This ensures existing code continues to work")
    print()


def demonstrate_configuration_options():
    """Demonstrate configuration options for different use cases."""
    
    print("=== Configuration Options ===\n")
    
    # Configuration for different scenarios
    configurations = [
        {
            "name": "Quick Preview",
            "description": "Fast email preview for listing",
            "config": {
                "mode": EmailRetrievalMode.BASIC,
                "include_attachments": False,
                "embed_images": False
            },
            "use_case": "Email list views, quick scanning"
        },
        {
            "name": "Detailed Analysis",
            "description": "Complete email analysis with all media",
            "config": {
                "mode": EmailRetrievalMode.ENHANCED,
                "include_attachments": True,
                "embed_images": True
            },
            "use_case": "Forensic analysis, content extraction"
        },
        {
            "name": "Performance Optimized",
            "description": "Balanced performance with smart caching",
            "config": {
                "mode": EmailRetrievalMode.LAZY,
                "include_attachments": True,
                "embed_images": False
            },
            "use_case": "Large email datasets, batch processing"
        },
        {
            "name": "Media Focused",
            "description": "Email with rich media content",
            "config": {
                "mode": EmailRetrievalMode.ENHANCED,
                "include_attachments": True,
                "embed_images": True
            },
            "use_case": "Marketing emails, newsletters, rich content"
        }
    ]
    
    for i, config in enumerate(configurations, 1):
        print(f"{i}. {config['name']}:")
        print(f"   Description: {config['description']}")
        print(f"   Configuration: {config['config']}")
        print(f"   Use Case: {config['use_case']}")
        print()


def demonstrate_error_handling():
    """Demonstrate error handling and edge cases."""
    
    print("=== Error Handling and Edge Cases ===\n")
    
    print("1. Invalid Email Number:")
    print("   Code: get_email_by_number_unified(-1, mode='basic')")
    print("   Result: Returns None with warning log")
    print()
    
    print("2. Out of Range Email Number:")
    print("   Code: get_email_by_number_unified(999, mode='basic')")
    print("   Result: Returns None with warning log")
    print()
    
    print("3. Invalid Mode:")
    print("   Code: get_email_by_number_unified(1, mode='invalid')")
    print("   Result: Returns None with error message")
    print()
    
    print("4. Session Error Fallback:")
    print("   Code: get_email_by_number_unified(1, mode='enhanced') with Outlook session error")
    print("   Result: Falls back to basic mode gracefully")
    print()


def demonstrate_performance_comparison():
    """Compare performance characteristics of different modes."""
    
    print("=== Performance Comparison ===\n")
    
    print("Mode Performance Characteristics:")
    print()
    
    modes = [
        {
            "name": "Basic",
            "speed": "Fastest",
            "memory": "Low",
            "network": "Minimal",
            "features": "Basic",
            "best_for": "Quick previews, large lists"
        },
        {
            "name": "Lazy",
            "speed": "Fast",
            "memory": "Medium",
            "network": "Minimal (cached)",
            "features": "Smart",
            "best_for": "General usage, mixed workloads"
        },
        {
            "name": "Enhanced",
            "speed": "Slower",
            "memory": "High",
            "network": "Full",
            "features": "Complete",
            "best_for": "Detailed analysis, media extraction"
        }
    ]
    
    for mode in modes:
        print(f"{mode['name']} Mode:")
        print(f"   Speed: {mode['speed']}")
        print(f"   Memory Usage: {mode['memory']}")
        print(f"   Network Calls: {mode['network']}")
        print(f"   Features: {mode['features']}")
        print(f"   Best For: {mode['best_for']}")
        print()


def main():
    """Main demonstration function."""
    
    print("Unified Email Retrieval Architecture")
    print("=" * 50)
    print()
    
    # Run all demonstrations
    demonstrate_unified_email_retrieval()
    print("\n" + "=" * 50 + "\n")
    
    demonstrate_mcp_tool_usage()
    print("\n" + "=" * 50 + "\n")
    
    demonstrate_configuration_options()
    print("\n" + "=" * 50 + "\n")
    
    demonstrate_error_handling()
    print("\n" + "=" * 50 + "\n")
    
    demonstrate_performance_comparison()
    
    print("\n" + "=" * 50)
    print("Demo completed! The unified architecture provides:")
    print("- Single, consistent API")
    print("- Configurable functionality levels")
    print("- Backward compatibility")
    print("- Performance optimization")
    print("- Better maintainability")


if __name__ == "__main__":
    main()