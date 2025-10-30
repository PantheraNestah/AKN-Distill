"""
Test if Python COM .Duplicate ranges actually persist changes
"""
import os, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Distill.settings')
django.setup()

from myapp.models import Document
import win32com.client
import pythoncom
from pathlib import Path

pythoncom.CoInitialize()

try:
    # Get a recent document
    doc_model = Document.objects.filter(processed_file__isnull=False).order_by('-id').first()
    
    if not doc_model:
        print("No processed document found!")
        exit()
    
    print(f"Testing with document: {doc_model.id}")
    print(f"File: {doc_model.processed_file_path}")
    
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    
    doc = word_app.Documents.Open(doc_model.processed_file_path)
    
    # Get first paragraph with text
    for para in doc.Paragraphs:
        if len(para.Range.Text.strip()) > 10:
            print(f"\nOriginal paragraph text (first 80 chars):")
            print(f"  '{para.Range.Text[:80]}'")
            
            # Test 1: Try to replace using Duplicate (VBA style)
            print(f"\nTest 1: Using .Duplicate (VBA style)")
            head = para.Range.Duplicate
            head.End = head.Start + 60
            
            find = head.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Text = "([A-Za-z]{1})"  # Match first letter
            find.Replacement.Text = "X"  # Replace with X
            find.MatchWildcards = True
            find.Forward = True
            find.Wrap = 0  # wdFindStop
            find.Format = False
            
            result = find.Execute(Replace=1)  # wdReplaceOne
            
            print(f"  Execute returned: {result}")
            print(f"  Head text after: '{head.Text[:80]}'")
            print(f"  Para text after: '{para.Range.Text[:80]}'")
            
            if "X" in para.Range.Text[:80]:
                print(f"  ✅ Change PERSISTED to paragraph!")
            else:
                print(f"  ❌ Change DID NOT persist to paragraph")
            
            break
    
    doc.Close(SaveChanges=False)
    word_app.Quit()
    
except Exception as e:
    print(f"Error: {str(e)}")
    import traceback
    traceback.print_exc()

finally:
    pythoncom.CoUninitialize()
