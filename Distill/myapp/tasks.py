from background_task import background
from django.utils import timezone
from .models import Document, RuleTaskState
from .services.document_processor import DocumentProcessingService
from .services.task_progress import TaskProgressService
import logging

logger = logging.getLogger(__name__)
progress_service = TaskProgressService()

@background(schedule=timezone.now())
def process_rule(document_id, rule_name, state):
    """
    Process a single rule for a document.
    Updates task state and schedules next rule if successful.
    """
    import pythoncom
    pythoncom.CoInitialize()  # Initialize COM for this thread
    
    logger.info(f"Processing rule {rule_name} (state {state}) for document {document_id}")
    
    try:
        # Get task state
        task_state = RuleTaskState.objects.get(
            document_id=document_id,
            rule_name=rule_name,
            state=state
        )
        
        start_time = timezone.now()
        task_state.status = 'processing'
        task_state.save()
        
        # Notify task start
        TaskProgressService.update_task_progress(
            document_id=document_id,
            rule_name=rule_name,
            progress=0,
            message=f"Starting rule: {rule_name}",
            status='processing'
        )

        # Define rules that need page parameters
        PAGE_PARAM_RULES = [
            'lists_dot_to_emdash',
            'remove_spaces_around_em_dash'
        ]

        # Create custom rule config for single rule
        rule_action = {
            'word_recipe': {
                'name': rule_name,
                'enabled': True,
            }
        }

        # Add page parameters only for rules that need them
        if rule_name in PAGE_PARAM_RULES:
            # Default to first 4 pages for rules that have page restrictions
            rule_action['word_recipe']['params'] = {
                'page_start': 1,
                'page_end': 4
            }

        rule_config = {
            'name': 'custom',
            'description': f'Processing rule: {rule_name}',
            'safety': {
                'allow_text_changes': True,
                'allow_style_changes': True
            },
            'steps': [{
                'name': rule_name,
                'select': {'document': True},
                'actions': [rule_action]
            }]
        }

        # Process the rule
        processor = DocumentProcessingService()
        result = processor.process_document(task_state.document, rule_config)

        # Handle both dictionary and boolean results
        success = result.get('ok') if isinstance(result, dict) else result

        if success:
            task_state.status = 'completed'
            task_state.state += 1
            
            # Get completion message
            completion_message = (result.get('description') if isinstance(result, dict) 
                               else f"Rule {rule_name} completed successfully")
            
            # Update progress and notify completion
            TaskProgressService.task_complete(
                document_id=document_id,
                rule_name=rule_name,
                success=True
            )
            
            logger.info(completion_message)

            # Schedule next rule if available
            next_task = RuleTaskState.objects.filter(
                document_id=document_id,
                status='pending'
            ).first()

            if next_task:
                logger.info(f"Scheduling next rule: {next_task.rule_name}")
                process_rule(document_id, next_task.rule_name, next_task.state)
            else:
                # All rules completed
                document = Document.objects.get(id=document_id)
                document.status = 'COMPLETED'
                document.processed_at = timezone.now()
                document.save()
                logger.info(f"All rules completed for document {document_id}")
        else:
            task_state.status = 'failed'
            error_msg = result.get('error', 'Unknown error') if isinstance(result, dict) else 'Rule processing failed'
            task_state.error_message = error_msg
            
            # Update progress and notify failure
            TaskProgressService.task_complete(
                document_id=document_id,
                rule_name=rule_name,
                success=False,
                error_message=error_msg
            )
            
            logger.error(f"Rule {rule_name} failed: {error_msg}")

    except Exception as e:
        logger.exception(f"Error processing rule {rule_name}")
        task_state.status = 'failed'
        task_state.error_message = str(e)

    finally:
        task_state.processing_time = timezone.now() - start_time
        task_state.save()
        pythoncom.CoUninitialize()  # Clean up COM for this thread