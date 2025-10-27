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
            # Initialize COM for this thread with multi-threaded apartment
            pythoncom.CoInitialize()
            
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
            
            # Open document
            if progress_callback:
                progress_callback(15, "Opening document...")

            doc = self.word_app.Documents.Open(document.original_file.path)
            
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
                    if progress_callback:
                        progress = 30 + int((idx / total_steps) * 40)  # Progress from 30% to 70%
                        progress_callback(progress, f"Applying rule: {step.name}")
                    
                    ops.apply_steps(
                        self.engine,
                        doc,
                        [step],  # Apply one step at a time
                        rules.safety,
                        logger
                    )
                
                # Export to PDF
                if progress_callback:
                    progress_callback(80, "Converting to PDF...")

                pdf_name = f"{Path(document.original_file.name).stem}.pdf"
                pdf_path = pdf_dir / pdf_name
                doc.SaveAs2(str(pdf_path), FileFormat=17)  # 17 = PDF
                
                # Update document record
                if progress_callback:
                    progress_callback(90, "Finalizing...")

                document.processed_file = f"processed/{pdf_name}"
                document.status = 'COMPLETED'
                document.save()

                if progress_callback:
                    progress_callback(100, "Processing complete!")
                
                return True
                
            finally:
                doc.Close(SaveChanges=False)
                
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
                    
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    def _initialize_word_engine(self):
        """Initialize Word COM engine"""
        # Initialize Word
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = False
        return engines.WordComEngine()