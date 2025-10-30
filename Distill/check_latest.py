import os, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Distill.settings')
django.setup()

from myapp.models import Document, RuleTaskState

# Get the most recent document
doc = Document.objects.order_by('-id').first()

print(f"\n{'='*70}")
print(f"LATEST DOCUMENT: ID {doc.id}")
print(f"{'='*70}")
print(f"Status: {doc.status}")
print(f"File: {doc.original_file.name if doc.original_file else 'None'}")
print(f"Processed: {doc.processed_file.name if doc.processed_file else 'None'}")
print(f"Upload time: {doc.uploaded_at.strftime('%H:%M:%S')}")

# Check its tasks
tasks = RuleTaskState.objects.filter(document_id=doc.id).order_by('created_at')

print(f"\nTasks ({tasks.count()}):")
for i, task in enumerate(tasks, 1):
    icon = "✅" if task.status == 'completed' else "❌" if task.status == 'failed' else "⏳"
    print(f"  {i}. {icon} {task.rule_name}: {task.status}")
    if task.error_message:
        print(f"      ERROR: {task.error_message[:80]}")

# Check if em-dash was applied
if doc.processed_file:
    import win32com.client
    import pythoncom
    import re
    
    pythoncom.CoInitialize()
    
    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        
        doc_obj = word_app.Documents.Open(doc.processed_file_path)
        
        # Check page 3 for em-dash patterns
        rng = doc_obj.GoTo(What=1, Which=1, Count=3)
        try:
            rNext = rng.GoTo(What=1, Which=2)
            rng.End = rNext.Start - 1
        except:
            rng.End = doc_obj.Content.End
        
        text = rng.Text[:500]
        
        em_count = len(re.findall(r'\d+—', text))
        dot_count = len(re.findall(r'\d+\.', text))
        
        print(f"\nEm-dash check (page 3):")
        print(f"  Em-dashes (e.g., '1—'): {em_count}")
        print(f"  Dots (e.g., '1.'): {dot_count}")
        
        if em_count > 0:
            print(f"  ✓ Em-dash WAS APPLIED")
        else:
            print(f"  ✗ Em-dash NOT APPLIED")
        
        # Show sample
        lines = text.split('\r')[:8]
        print(f"\nSample from page 3:")
        for line in lines:
            if line.strip():
                print(f"    {line.strip()[:70]}")
        
        doc_obj.Close(SaveChanges=False)
        word_app.Quit()
        
    finally:
        pythoncom.CoUninitialize()

print(f"\n{'='*70}")
