from django.db import models
from django.urls import reverse
from django.conf import settings
import os

class RuleTaskState(models.Model):
    TASK_STATUS = [
        ('pending', 'Pending'),
        ('processing', 'Processing'),
        ('completed', 'Completed'),
        ('failed', 'Failed'),
    ]

    document = models.ForeignKey('Document', on_delete=models.CASCADE, related_name='rule_tasks')
    rule_name = models.CharField(max_length=100)
    status = models.CharField(max_length=20, choices=TASK_STATUS, default='pending')
    state = models.IntegerField(default=0)
    processing_time = models.DurationField(null=True, blank=True)
    error_message = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['created_at']

    def __str__(self):
        return f"{self.rule_name} - {self.status}"

class Document(models.Model):
    PROCESSING_STATUS = [
        ('UPLOADED', 'Uploaded'),
        ('PROCESSING', 'Processing'),
        ('COMPLETED', 'Completed'),
        ('FAILED', 'Failed'),
    ]

    original_file = models.FileField(upload_to='documents/')
    processed_file = models.FileField(upload_to='processed/', null=True, blank=True)
    status = models.CharField(max_length=20, choices=PROCESSING_STATUS, default='UPLOADED')
    uploaded_at = models.DateTimeField(auto_now_add=True)
    processed_at = models.DateTimeField(null=True, blank=True)
    error_message = models.TextField(blank=True, null=True)
    progress = models.IntegerField(default=0)

    def __str__(self):
        return f"Document {self.id} - {self.status}"

    def get_absolute_url(self):
        # Don't redirect to process_status if processing
        if self.status == 'PROCESSING':
            return None
        return reverse('process_status', args=[str(self.id)])
    
    @property
    def file_path(self):
        """Returns the full path of the original file"""
        return os.path.join(settings.MEDIA_ROOT, str(self.original_file))
    
    @property
    def processed_file_path(self):
        """Returns the full path of the processed file"""
        return os.path.join(settings.MEDIA_ROOT, str(self.processed_file)) if self.processed_file else None