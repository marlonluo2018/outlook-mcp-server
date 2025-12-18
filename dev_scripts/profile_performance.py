#!/usr/bin/env python3
"""
Performance profiling script to identify bottlenecks in email listing.
"""

import win32com.client
import pythoncom
import time
from datetime import datetime, timedelta, timezone
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def profile_email_processing():
    """Profile different aspects of email processing."""
    pythoncom.CoInitialize()
    
    try:
        # Initialize Outlook
        start_time = time.time()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        init_time = time.time() - start_time
        logger.info(f"Outlook initialization: {init_time:.3f}s")
        
        # Get Inbox folder
        start_time = time.time()
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        folder_time = time.time() - start_time
        logger.info(f"Folder access: {folder_time:.3f}s")
        
        # Get total item count
        start_time = time.time()
        total_items = inbox.Items.Count
        count_time = time.time() - start_time
        logger.info(f"Item count: {count_time:.3f}s - Total: {total_items}")
        
        # Test different batch sizes
        batch_sizes = [10, 25, 50, 100]
        test_items = min(200, total_items)
        
        for batch_size in batch_sizes:
            logger.info(f"\n--- Testing batch size: {batch_size} ---")
            
            # Test item access
            start_time = time.time()
            items_accessed = 0
            filtered_items = []
            
            for i in range(0, test_items, batch_size):
                batch_start = time.time()
                batch_items = []
                
                for j in range(i, min(i + batch_size, test_items)):
                    try:
                        item_index = total_items - j
                        if item_index <= 0:
                            continue
                            
                        item = inbox.Items.Item(item_index)
                        if not item:
                            continue
                            
                        # Basic filtering (similar to production code)
                        if not hasattr(item, 'ReceivedTime') or not hasattr(item, 'Class'):
                            continue
                        
                        if hasattr(item, 'Class') and item.Class != 43:  # 43 = olMail
                            continue
                        
                        if not item.ReceivedTime:
                            continue
                        
                        batch_items.append(item)
                        items_accessed += 1
                        
                    except Exception as e:
                        continue
                
                batch_time = time.time() - batch_start
                logger.info(f"  Batch {i//batch_size + 1}: {len(batch_items)} items, {batch_time:.3f}s")
                
                # Simulate date filtering
                date_limit = datetime.now(timezone.utc) - timedelta(days=7)
                filtered_batch = []
                
                for item in batch_items:
                    try:
                        item_time = item.ReceivedTime
                        if item_time.tzinfo is None:
                            item_time = item_time.replace(tzinfo=timezone.utc)
                        
                        if item_time >= date_limit:
                            filtered_batch.append(item)
                    except:
                        continue
                
                filtered_items.extend(filtered_batch)
                
                # Stop if we have enough for testing
                if len(filtered_items) >= 50:
                    break
            
            total_time = time.time() - start_time
            logger.info(f"Total time: {total_time:.3f}s")
            logger.info(f"Items accessed: {items_accessed}")
            logger.info(f"Filtered items: {len(filtered_items)}")
            logger.info(f"Average time per item: {total_time/items_accessed*1000:.1f}ms")
        
        # Test email extraction
        logger.info(f"\n--- Testing email extraction ---")
        if filtered_items:
            start_time = time.time()
            extracted_emails = []
            
            for item in filtered_items[:20]:  # Test with 20 items
                email_info = {
                    "subject": getattr(item, 'Subject', 'No Subject'),
                    "sender": getattr(item, 'SenderName', 'Unknown'),
                    "received_time": getattr(item, 'ReceivedTime', None),
                    "entry_id": getattr(item, 'EntryID', ''),
                }
                
                if email_info["received_time"] is None:
                    email_info["received_time"] = "Unknown"
                else:
                    email_info["received_time"] = str(email_info["received_time"])
                
                extracted_emails.append(email_info)
            
            extraction_time = time.time() - start_time
            logger.info(f"Extraction time for {len(extracted_emails)} emails: {extraction_time:.3f}s")
            logger.info(f"Average time per email: {extraction_time/len(extracted_emails)*1000:.1f}ms")
        
        # Test COM object cleanup
        logger.info(f"\n--- Testing COM object cleanup ---")
        start_time = time.time()
        
        # Clear references
        del filtered_items
        del extracted_emails
        
        cleanup_time = time.time() - start_time
        logger.info(f"Cleanup time: {cleanup_time:.3f}s")
        
    except Exception as e:
        logger.error(f"Error during profiling: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        pythoncom.CoUninitialize()

def suggest_optimizations():
    """Suggest specific optimizations based on common bottlenecks."""
    print("\n" + "="*60)
    print("PERFORMANCE OPTIMIZATION SUGGESTIONS:")
    print("="*60)
    
    suggestions = [
        "1. Reduce COM object creation overhead:",
        "   - Reuse Outlook session objects when possible",
        "   - Minimize folder.Items access calls",
        "",
        "2. Optimize batch processing:",
        "   - Use optimal batch size (25-50 items)",
        "   - Process items in parallel if possible",
        "",
        "3. Minimize attribute access:",
        "   - Cache frequently accessed attributes",
        "   - Use hasattr() checks efficiently",
        "",
        "4. Early termination for date filtering:",
        "   - Stop processing when emails are older than date limit",
        "   - Process newest emails first",
        "",
        "5. Memory management:",
        "   - Clear COM object references promptly",
        "   - Use generators for large datasets"
    ]
    
    for suggestion in suggestions:
        print(suggestion)

if __name__ == "__main__":
    print("Profiling email processing performance...")
    profile_email_processing()
    suggest_optimizations()