"""
Service for handling task progress updates and notifications
"""
from channels.layers import get_channel_layer
from asgiref.sync import async_to_sync
from django.utils import timezone
from ..models import Document, RuleTaskState
import logging

logger = logging.getLogger(__name__)

class TaskProgressService:
    @staticmethod
    def update_task_progress(document_id: int, rule_name: str, progress: int, message: str, status: str = None):
        """
        Update task progress and send notification via WebSocket
        
        Args:
            document_id: ID of the document being processed
            rule_name: Name of the current rule being applied
            progress: Progress percentage (0-100)
            message: Status message
            status: Task status (if changed)
        """
        try:
            # Update task state
            task_state = RuleTaskState.objects.get(
                document_id=document_id,
                rule_name=rule_name,
                status__in=['pending', 'processing']  # Only update active tasks
            )
            
            if status:
                task_state.status = status
                
            task_state.save()
            
            # Calculate overall document progress
            document = Document.objects.get(id=document_id)
            total_tasks = document.rule_tasks.count()
            completed_tasks = document.rule_tasks.filter(status='completed').count()
            current_task_progress = progress if task_state.status == 'processing' else 0
            
            # Calculate overall progress:
            # Each completed task contributes (100/total_tasks)%
            # Current task contributes (progress/100 * 100/total_tasks)%
            base_progress = (completed_tasks * 100) / total_tasks
            current_progress = (current_task_progress / 100) * (100 / total_tasks)
            overall_progress = base_progress + current_progress
            
            # Update document progress
            document.progress = int(overall_progress)
            document.save()

            try:
                # Send WebSocket notification and log to console
                channel_layer = get_channel_layer()
                if channel_layer is not None:
                    task_progress_data = {
                        'type': 'processing_update',
                        'progress': overall_progress,
                        'status': status or task_state.status,
                        'message': message,
                        'rule_name': rule_name,
                        'timestamp': timezone.now().isoformat(),
                        'task_details': {
                            'current_task': rule_name,
                            'task_progress': progress,
                            'completed_tasks': completed_tasks,
                            'total_tasks': total_tasks,
                            'current_task_number': completed_tasks + 1
                        }
                    }
                    
                    async_to_sync(channel_layer.group_send)(
                        f'doc_{document_id}',
                        task_progress_data
                    )
                
                # Log progress to console
                console_message = f"""
{'='*80}
Document {document_id} - {rule_name}
Status: {status or task_state.status}
Progress: [{('='*int(overall_progress/2)).ljust(50)}] {overall_progress:.1f}%
Current Task: {completed_tasks + 1}/{total_tasks}
Message: {message}
{'='*80}
"""
                print(console_message)
            except Exception as e:
                logger.warning(f"Failed to send WebSocket notification: {str(e)}")
            logger.info(f"Task progress updated - Document: {document_id}, Rule: {rule_name}, "
                       f"Progress: {progress}%, Message: {message}")
                       
        except Exception as e:
            logger.error(f"Failed to update task progress: {str(e)}")

    @staticmethod
    def task_complete(document_id: int, rule_name: str, success: bool, error_message: str = None):
        """
        Mark a task as complete or failed and notify via WebSocket
        
        Args:
            document_id: ID of the document being processed
            rule_name: Name of the rule that completed
            success: Whether the task completed successfully
            error_message: Error message if task failed
        """
        try:
            task_state = RuleTaskState.objects.get(
                document_id=document_id,
                rule_name=rule_name
            )
            
            status = 'completed' if success else 'failed'
            task_state.status = status
            if error_message:
                task_state.error_message = error_message
            task_state.save()
            
            # Update document status if all tasks are complete
            document = Document.objects.get(id=document_id)
            all_tasks = document.rule_tasks.all()
            all_complete = all(task.status in ['completed', 'failed'] for task in all_tasks)
            any_failed = any(task.status == 'failed' for task in all_tasks)
            
            if all_complete:
                document.status = 'FAILED' if any_failed else 'COMPLETED'
                document.save()
                
                try:
                    # Send completion notification
                    channel_layer = get_channel_layer()
                    if channel_layer is not None:
                        async_to_sync(channel_layer.group_send)(
                            f'doc_{document_id}',
                            {
                                'type': 'processing_complete',
                                'status': document.status,
                                'preview_url': document.processed_file.url if document.processed_file else None,
                                'error_message': error_message if any_failed else None
                            }
                        )
                    
                    # Print completion message to terminal
                    console_message = f"""
{'='*80}
Document {document_id} Processing Complete
Status: {document.status}
{'Error: ' + error_message if error_message else 'Successfully processed'}
{'='*80}
"""
                    print(console_message)
                except Exception as e:
                    logger.warning(f"Failed to send completion notification: {str(e)}")
            
            logger.info(f"Task {rule_name} {'completed' if success else 'failed'} for document {document_id}")
            
        except Exception as e:
            logger.error(f"Failed to mark task complete: {str(e)}")