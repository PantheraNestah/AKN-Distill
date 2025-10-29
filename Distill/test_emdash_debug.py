"""
Quick test script to check what the em-dash recipe is doing.
Run this with: python test_emdash_debug.py
"""
import os
import sys
import django

# Setup Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Distill.settings')
django.setup()

from myapp.models import Document
from myapp.processor.recipes_word import lists_dot_to_emdash
from myapp.processor.engines import WordComEngine
from pathlib import Path

# Get the most recent document
doc = Document.objects.filter(status='COMPLETED').order_by('-id').first()

if not doc or not doc.processed_file:
    print("No processed document found!")
    sys.exit(1)

print(f"Testing document: {doc.id}")
print(f"File: {doc.processed_file_path}")
print("\n" + "="*60)

# Run the recipe directly with debug output
engine = WordComEngine()
doc_obj = engine.open_document(Path(doc.processed_file_path))

try:
    result = lists_dot_to_emdash.lists_dot_to_emdash_py(doc_obj)
    
    print(f"\nResult: {result['ok']}")
    print(f"Changed: {result.get('changed', 0)} replacements")
    print(f"\nWarnings/Debug output ({len(result.get('warnings', []))} messages):")
    print("="*60)
    
    for i, warning in enumerate(result.get('warnings', [])[:100], 1):
        print(f"{i}. {warning}")
    
    if len(result.get('warnings', [])) > 100:
        print(f"\n... and {len(result.get('warnings', [])) - 100} more messages")
        
finally:
    doc_obj.Close(SaveChanges=False)
    engine.close()

print("\n" + "="*60)
print("Test complete! Check the output above for DEBUG/FALLBACK messages")
