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
from ..processor.pipeline import run
from ..processor import engines
from ..processor.rules import Rules

logger = logging.getLogger(__name__)

class DocumentProcessingService:
    def __init__(self):
        self.word_app = None
        self.engine = None
        
    def process_document(self, document: Document) -> bool:
        """
        Process a document using Word COM automation and custom recipes.
        Returns True if processing was successful, False otherwise.
        """
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            document.status = 'PROCESSING'
            document.save()
            
            # Load processing rules
            rules_path = Path(__file__).parent.parent / 'processor' / 'config' / 'rules.yaml'
            with open(rules_path, 'r') as f:
                rules = Rules.from_dict(yaml.safe_load(f))
            
            # Initialize engine
            self.engine = engines.WordComEngine()
            
            # Get Word application from engine
            self.word_app = self.engine.app
            
            # Open document
            doc = self.word_app.Documents.Open(document.original_file.path)
            
            try:
                # Create output directories if they don't exist
                pdf_dir = Path(settings.MEDIA_ROOT) / 'processed'
                pdf_dir.mkdir(parents=True, exist_ok=True)
                
                # Process document using rules
                run(
                    self.engine,
                    doc,
                    rules,
                    base_out_dir=str(pdf_dir),
                    write_audit=True,
                    dry_run=False,
                    verbose=True
                )
                
                # Export to PDF
                pdf_name = f"{Path(document.original_file.name).stem}.pdf"
                pdf_path = pdf_dir / pdf_name
                doc.SaveAs2(str(pdf_path), FileFormat=17)  # 17 = PDF
                
                # Update document record
                document.processed_file = f"processed/{pdf_name}"
                document.status = 'COMPLETED'
                document.save()
                
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
            if self.word_app:
                try:
                    self.word_app.Quit()
                except:
                    pass
                finally:
                    pythoncom.CoUninitialize()
                
    def _initialize_word_engine(self):
        """Initialize Word COM engine"""
        # Initialize Word
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = False
        return engines.WordComEngine()