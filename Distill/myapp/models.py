from django.db import models
from django.urls import reverse
from django.conf import settings
import os
import re

class Document(models.Model):
    PROCESSING_STATUS = [
        ('UPLOADED', 'Uploaded'),
        ('PROCESSING', 'Processing'),
        ('COMPLETED', 'Completed'),
        ('FAILED', 'Failed'),
    ]

    original_file = models.FileField(upload_to='documents/')
    processed_file = models.FileField(upload_to='processed/', null=True, blank=True)
    description = models.TextField(blank=True, null=True)  # NEW FIELD
    status = models.CharField(max_length=20, choices=PROCESSING_STATUS, default='UPLOADED')
    uploaded_at = models.DateTimeField(auto_now_add=True)
    processed_at = models.DateTimeField(null=True, blank=True)
    error_message = models.TextField(blank=True, null=True)
    progress = models.IntegerField(default=0)

    def __str__(self):
        return f"Document {self.id} - {self.status}"

    def get_absolute_url(self):
        return reverse('process_status', args=[str(self.id)])
    
    @property
    def file_path(self):
        """Returns the full path of the original file"""
        return os.path.join(settings.MEDIA_ROOT, str(self.original_file))
    
    @property
    def processed_file_path(self):
        """Returns the full path of the processed file"""
        return os.path.join(settings.MEDIA_ROOT, str(self.processed_file)) if self.processed_file else None
    
    def get_clean_filename(self):
        """Returns cleaned filename without path, random prefix, and extension"""
        # Get just the filename from the full path
        filename = os.path.basename(str(self.original_file))
        
        # Remove the random prefix pattern (e.g., "1a_ryHc3Ao._")
        # Pattern: documents/[random]_[random]._
        filename = re.sub(r'^[^_]+_[^_]+\._', '', filename)
        
        # Remove file extension
        filename = os.path.splitext(filename)[0]
        
        # Replace underscores with spaces for better readability
        filename = filename.replace('_', ' ')
        
        return filename