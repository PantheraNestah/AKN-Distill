import os, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Distill.settings')
django.setup()

from myapp.models import Document
import win32com.client
import pythoncom
import re

pythoncom.CoInitialize()

try:
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    
    # Check documents 22, 21, 20
    docs_to_check = Document.objects.filter(id__in=[22, 21, 20]).order_by('-id')
    
    print("="*70)
    print("CHECKING EM-DASH IN LAST 3 DOCUMENTS")
    print("="*70)
    
    for doc_model in docs_to_check:
        print(f"\n📄 Document ID {doc_model.id}:")
        print(f"   File: {doc_model.original_file.name if doc_model.original_file else 'None'}")
        print(f"   Processed: {doc_model.processed_file.name if doc_model.processed_file else 'None'}")
        
        if not doc_model.processed_file_path or not os.path.exists(doc_model.processed_file_path):
            print("   ⚠️  Processed file doesn't exist!")
            continue
        
        try:
            doc = word_app.Documents.Open(doc_model.processed_file_path)
            
            total_pages = doc.ComputeStatistics(2)  # wdStatisticPages
            pages_to_check = min(4, total_pages)
            
            em_dash_count = 0
            dot_number_count = 0
            samples = []
            
            for pg in range(1, pages_to_check + 1):
                rng = doc.GoTo(What=1, Which=1, Count=pg)
                try:
                    rNext = rng.GoTo(What=1, Which=2)
                    rng.End = rNext.Start - 1
                except:
                    rng.End = doc.Content.End
                
                text = rng.Text[:800]
                
                # Count patterns
                em_matches = re.findall(r'\d+—', text)
                dot_matches = re.findall(r'\d+\.(?=\s*[A-Z])', text)
                
                em_dash_count += len(em_matches)
                dot_number_count += len(dot_matches)
                
                # Get sample of first few lines on page 3 (where Section usually is)
                if pg == 3:
                    lines = text.split('\r')[:10]
                    samples = [l.strip() for l in lines if l.strip() and len(l.strip()) > 3][:5]
            
            doc.Close(SaveChanges=False)
            
            print(f"   Pages checked: {pages_to_check}")
            print(f"   ✅ Em-dash patterns (e.g., '1—'): {em_dash_count}")
            print(f"   ❌ Dot patterns (e.g., '1.'): {dot_number_count}")
            
            if samples:
                print(f"   Sample from page 3:")
                for s in samples[:3]:
                    print(f"      {s[:80]}")
            
            if em_dash_count > 0:
                print(f"   ✓ Em-dash rule WAS APPLIED")
            elif dot_number_count > 0:
                print(f"   ✗ Em-dash rule NOT APPLIED (still has dots)")
            else:
                print(f"   ? No numbered items found")
                
        except Exception as e:
            print(f"   ❌ Error: {str(e)}")
    
    word_app.Quit()
    
except Exception as e:
    print(f"Error: {str(e)}")
    import traceback
    traceback.print_exc()

finally:
    pythoncom.CoUninitialize()

print("\n" + "="*70)
