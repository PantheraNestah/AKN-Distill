"""
Document processing service for handling Word document formatting.
"""
from pathlib import Path
import logging
import os
from django.conf import settings
import win32com.client
import pythoncom
import yaml
from ..models import Document
from ..processor import engines
from ..processor.rules import Rules
from ..processor import ops

logger = logging.getLogger(__name__)

class DocumentProcessingService:
    def __init__(self):
        self.word_app = None
        self.engine = None
        
    def process_document(self, document: Document, custom_rules=None, progress_callback=None) -> bool:
        """
        Process a document using Word COM automation and custom recipes.
        Returns True if processing was successful, False otherwise.
        
        Args:
            document: The Document model instance to process
            custom_rules: Optional dictionary containing custom rules configuration.
                        If not provided, will load from default rules.yaml
            progress_callback: Optional callback function(progress: int, message: str)
                             to report processing progress
        """
        try:
            # COM should already be initialized by the background task
            document.status = 'PROCESSING'
            document.save()
            
            # Load processing rules
            if custom_rules:
                rules = Rules.from_dict(custom_rules)
            else:
                rules_path = Path(__file__).parent.parent / 'processor' / 'config' / 'rules.yaml'
                with open(rules_path, 'r') as f:
                    rules = Rules.from_dict(yaml.safe_load(f))
            
            # Clean up any existing Word instances
            try:
                if self.word_app:
                    self.word_app.Quit()
            except:
                pass
            
            if progress_callback:
                progress_callback(5, "Initializing Word...")

            # Initialize engine with proper COM setup
            self.engine = None
            for retry in range(3):  # Retry up to 3 times
                try:
                    self.engine = engines.WordComEngine()
                    # Get Word application from engine
                    self.word_app = self.engine.app
                    self.word_app.Visible = False  # Ensure Word stays hidden
                    self.word_app.DisplayAlerts = False  # Suppress Word dialogs
                    break
                except Exception as e:
                    logger.warning(f"Word initialization attempt {retry + 1} failed: {str(e)}")
                    if retry == 2:  # Last attempt failed
                        raise
                    try:
                        pythoncom.CoUninitialize()
                        pythoncom.CoInitialize()
                    except:
                        pass
            
            if progress_callback:
                progress_callback(10, "Word initialized successfully")
            
            # Determine which file to open:
            # If this is not the first rule and processed_file exists, use it
            # Otherwise use the original file
            if document.processed_file and os.path.exists(document.processed_file.path):
                file_to_process = document.processed_file.path
                logger.info(f"Processing existing processed file: {file_to_process}")
            else:
                file_to_process = document.original_file.path
                logger.info(f"Processing original file: {file_to_process}")
            
            # Open document
            if progress_callback:
                progress_callback(15, "Opening document...")

            doc = self.word_app.Documents.Open(file_to_process)
            
            try:
                # Create output directories if they don't exist
                if progress_callback:
                    progress_callback(20, "Preparing output directories...")

                pdf_dir = Path(settings.MEDIA_ROOT) / 'processed'
                pdf_dir.mkdir(parents=True, exist_ok=True)
                
                if progress_callback:
                    progress_callback(30, "Applying formatting rules...")

                # Process document using rules
                total_steps = len(rules.steps)
                for idx, step in enumerate(rules.steps, 1):
                    current_progress = 30 + int((idx / total_steps) * 40)  # Progress from 30% to 70%
                    progress_message = f"Applying rule: {step.name}"
                    
                    if progress_callback:
                        progress_callback(current_progress, progress_message)
                    
                    # Apply the step and track progress
                    result = ops.apply_steps(
                        self.engine,
                        doc,
                        [step],  # Apply one step at a time
                        rules.safety,
                        logger
                    )
                    
                    # Report detailed progress through callback if available
                    if isinstance(result, dict) and 'count_updated' in result:
                        details = f"Updated {result['count_updated']} items"
                        if progress_callback:
                            progress_callback(current_progress, f"{progress_message} - {details}")
                
                # Save the processed Word document
                if progress_callback:
                    progress_callback(75, "Saving processed Word document...")

                # Determine output filename
                docx_name = f"{Path(document.original_file.name).stem}_processed.docx"
                docx_path = pdf_dir / docx_name
                
                # Save changes to the current document first
                doc.Save()  # Save in-place changes
                
                # Save as new processed file
                doc.SaveAs2(str(docx_path), FileFormat=16)  # 16 = .docx format
                
                # Update document record to point to processed file
                # This ensures the next rule will work on the updated document
                document.processed_file = f"processed/{docx_name}"
                document.save()
                
                if progress_callback:
                    progress_callback(80, "Reopening document for PDF conversion...")
                
                # Close and reopen to ensure all changes are persisted
                doc.Close(SaveChanges=False)
                doc = self.word_app.Documents.Open(str(docx_path))
                
                # Export to PDF
                if progress_callback:
                    progress_callback(85, "Converting to PDF...")

                pdf_name = f"{Path(document.original_file.name).stem}_processed.pdf"
                pdf_path = pdf_dir / pdf_name
                doc.SaveAs2(str(pdf_path), FileFormat=17)  # 17 = PDF
                
                # Update document record with both files
                if progress_callback:
                    progress_callback(95, "Finalizing...")

                document.pdf_file = f"processed/{pdf_name}"  # PDF
                document.status = 'COMPLETED'
                document.save()

                if progress_callback:
                    progress_callback(100, "Processing complete!")
                
                return True
                
            finally:
                doc.Close(SaveChanges=False)  # No need to save again, already saved above
                
        except Exception as e:
            logger.error(f"Document processing failed: {str(e)}")
            document.status = 'FAILED'
            document.error_message = str(e)
            document.save()
            return False
            
        finally:
            # Cleanup in reverse order of creation
            if self.word_app:
                try:
                    for doc in self.word_app.Documents:
                        try:
                            doc.Close(SaveChanges=False)
                        except:
                            pass
                    self.word_app.Quit()
                except:
                    pass
                
                # Force cleanup of any remaining Word processes if needed
                try:
                    import win32com.client
                    win32com.client.Dispatch("WScript.Shell").Run("taskkill /f /im WINWORD.EXE", 0, True)
                except:
                    pass
                    
            # COM cleanup is handled by the background task
    def _initialize_word_engine(self):
        """Initialize Word COM engine"""
        # Initialize Word
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = False
        return engines.WordComEngine()