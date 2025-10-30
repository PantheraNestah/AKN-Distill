import os, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Distill.settings')
django.setup()

from myapp.models import RuleTaskState, Document

print("Recent em-dash tasks:")
recent = RuleTaskState.objects.filter(rule_name='lists_dot_to_emdash').order_by('-id')[:5]
for t in recent:
    doc = Document.objects.get(id=t.document_id)
    print(f"  Doc {t.document_id}: {t.status} at {t.created_at.strftime('%H:%M:%S')}")
    if t.error_message:
        print(f"    ERROR: {t.error_message[:100]}")
    print(f"    File: {doc.original_file.name if doc.original_file else 'None'}")
