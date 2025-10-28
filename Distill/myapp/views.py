from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.http import HttpResponse, FileResponse
from .models import Document
from .services.document_processor import DocumentProcessingService
import logging
import os

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

def upload_document(request):
    if request.method == 'POST' and request.FILES.get('document'):
        uploaded_file = request.FILES['document']
        
        # Validate file type
        if not uploaded_file.name.endswith(('.doc', '.docx')):
            messages.error(request, 'Please upload a Word document (.doc or .docx)')
            return redirect('document_list')

        # Create document record
        document = Document.objects.create(
            original_file=uploaded_file,
            status='UPLOADED'
        )

        try:
            # Start processing
            processor = DocumentProcessingService()
            success = processor.process_document(document)

            if success:
                messages.success(request, 'Document uploaded and processed successfully.')
            else:
                messages.error(request, 'Document processing failed.')
        except Exception as e:
            document.status = 'failed'
            document.error_message = str(e)
            document.save()
            messages.error(request, f'Error processing document: {str(e)}')
        
        return redirect('document_detail', pk=document.id)
    
    return redirect('document_list')

def process_document(request, document_id):
    document = get_object_or_404(Document, id=document_id)
    
    if document.status != 'UPLOADED':
        messages.error(request, 'This document is not in an uploaded state.')
        return redirect('document_detail', pk=document.id)

    try:
        processor = DocumentProcessingService()
        success = processor.process_document(document)

        if success:
            messages.success(request, 'Document processed successfully.')
        else:
            messages.error(request, 'Document processing failed.')
    except Exception as e:
        document.status = 'failed'
        document.error_message = str(e)
        document.save()
        messages.error(request, f'Error processing document: {str(e)}')

    return redirect('document_detail', pk=document.id)