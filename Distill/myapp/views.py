from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.http import HttpResponse, FileResponse
from django.utils import timezone
from .models import Document
from .services.document_processor import DocumentProcessingService
import logging
import os
import yaml
from pathlib import Path

logger = logging.getLogger(__name__)

def document_list(request):
    documents = Document.objects.all().order_by('-uploaded_at')
    return render(request, 'myapp/document_list.html', {
        'documents': documents
    })

def document_detail(request, pk):
    document = get_object_or_404(Document, pk=pk)
    return render(request, 'myapp/document_detail.html', {
        'document': document
    })

from django.http import JsonResponse
import threading

def process_document_view(request):
    if request.method == 'POST':
        if not request.FILES.get('document'):
            return JsonResponse({'error': 'Please select a document to upload'}, status=400)

        uploaded_file = request.FILES['document']
        
        # Validate file type
        if not uploaded_file.name.endswith('.docx'):
            return JsonResponse({'error': 'Please upload a Word document (.docx)'}, status=400)

        # Get selected rules
        selected_rules = request.POST.getlist('rules')
        
        # Validate that at least one rule is selected
        if not selected_rules:
            return JsonResponse({
                'error': 'Please select at least one processing rule'
            }, status=400)
        
        try:
            # Create document record
            document = Document.objects.create(
                original_file=uploaded_file,
                status='UPLOADED'
            )

            # Start processing in background thread
            thread = threading.Thread(
                target=process_document_background,
                args=(document.id, selected_rules)
            )
            thread.daemon = True
            thread.start()

            return JsonResponse({
                'document_id': document.id,
                'message': 'Document processing started'
            })
        except Exception as e:
            logger.exception("Error starting document processing")
            return JsonResponse({
                'error': str(e)
            }, status=500)
    
    return render(request, 'myapp/process_new.html')

def document_status(request, document_id):
    """Return the current status of a document being processed."""
    try:
        document = get_object_or_404(Document, id=document_id)
        data = {
            'status': document.status,
            'error_message': document.error_message or '',
            'is_complete': document.status in ['COMPLETED', 'FAILED'],
            'processed_at': document.processed_at.isoformat() if document.processed_at else None,
        }
        
        if document.status == 'COMPLETED':
            data['processed_file_url'] = document.processed_file.url if document.processed_file else None
            
        return JsonResponse(data)
    except Exception as e:
        logger.exception("Error checking document status")
        return JsonResponse({
            'error': str(e)
        }, status=500)

def process_document_background(document_id, selected_rules):
    """Background task to process document."""
    document = Document.objects.get(id=document_id)
    
    try:
        # Update status to processing
        document.status = 'PROCESSING'
        document.save()

        # Generate custom rules
        custom_rules = generate_custom_rules(selected_rules, 'custom')
        logger.info(f"Applying selected rules: {selected_rules}")

        # Initialize and run processor
        processor = DocumentProcessingService()
        success = processor.process_document(document, custom_rules)

        if success:
            document.status = 'COMPLETED'
            document.processed_at = timezone.now()
            document.save()
        else:
            raise Exception("Processing failed")

    except Exception as e:
        document.status = 'FAILED'
        document.error_message = str(e)
        document.save()
        logger.exception("Document processing failed")

def generate_custom_rules(selected_rules, preset='custom'):
    """Generate a custom rules configuration based on selected rules."""
    
    # Define preset configurations
    PRESETS = {
        'legal_default': ['fix_aos_all_parts', 'follow_number_with_none_level2', 
                         'follow_number_with_none_level3', 'tighten_level3_spacing'],
        'minimal': ['fix_aos_all_parts'],
        'custom': []  # Use only selected rules
    }

    # If no rules selected, use minimal preset as default
    active_rules = selected_rules if selected_rules else PRESETS['minimal']

    # Validate selected rules
    if not active_rules:
        raise ValueError("At least one processing rule must be selected")

    # Build rules configuration
    rules_config = {
        'name': 'custom',
        'description': 'Custom processing rules',
        'safety': {
            'allow_text_changes': True,
            'allow_style_changes': True
        },
        'steps': []
    }

    # Map rule names to their configurations
    RULE_CONFIGS = {
        'fix_aos_all_parts': {
            'name': 'Apply AOS Formatting',
            'select': {'document': True},
            'actions': [{
                'word_recipe': {
                    'name': 'fix_aos_all_parts',
                    'enabled': True,
                    'params': {'block_width_cm': 12.0}
                }
            }]
        },
        'follow_number_with_none_level2': {
            'name': 'Fix Level 2 Numbering',
            'select': {'document': True},
            'actions': [{
                'word_recipe': {
                    'name': 'follow_number_with_none_level2',
                    'enabled': True,
                    'params': {'page_start': 1, 'page_end': 999}
                }
            }]
        },
        'follow_number_with_none_level3': {
            'name': 'Fix Level 3 Numbering',
            'select': {'document': True},
            'actions': [{
                'word_recipe': {
                    'name': 'follow_number_with_none_level3',
                    'enabled': True
                }
            }]
        },
        'tighten_level3_spacing': {
            'name': 'Tighten Level 3 Spacing',
            'select': {'document': True},
            'actions': [{
                'word_recipe': {
                    'name': 'tighten_level3_spacing',
                    'enabled': True
                }
            }]
        }
    }

    # Add selected rules to configuration
    for rule in active_rules:
        if rule in RULE_CONFIGS:
            rules_config['steps'].append(RULE_CONFIGS[rule])

    return rules_config