"""
Document processing service for handling Word document formatting.
"""
from pathlib import Path
import logging
import os
from django.conf import settings

# Defer importing Windows-only COM libraries until needed. On non-Windows
# systems these imports will not be present; importing them at module load
# time causes Django to fail to start (ModuleNotFoundError). Use HAVE_WIN32
# to gate COM-specific code paths.
try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
    HAVE_WIN32 = True
except Exception:
    win32com = None
    pythoncom = None
    HAVE_WIN32 = False
import yaml
from ..models import Document
from ..processor import engines
from ..processor.rules import Rules
from ..processor import ops
import subprocess

def _convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    """Convert a .docx file to PDF using LibreOffice/soffice headless.

    This is a best-effort converter; it requires `soffice` (LibreOffice) to be
    available on PATH. The function will raise RuntimeError if conversion fails.
    """
    # Prefer soffice (LibreOffice); some systems have `libreoffice` binary name
    cmd_candidates = [
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(docx_path)],
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(docx_path)],
    ]

    for cmd in cmd_candidates:
        try:
            res = subprocess.run(cmd, check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except FileNotFoundError:
            continue

        if res.returncode == 0:
            # LibreOffice will create a file named <stem>.pdf in outdir
            if pdf_path.exists():
                return
            # Try common output name
            alt = pdf_path.parent / (docx_path.stem + ".pdf")
            if alt.exists():
                alt.replace(pdf_path)
                return
        # else try next candidate

    raise RuntimeError("Failed to convert DOCX to PDF: soffice/libreoffice not available or conversion failed")

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
                progress_callback(5, "Initializing docx engine...")

            # Use the cross-platform DocxEngine
            self.engine = engines.pick_engine("docx")
            # No special app object for python-docx
            self.word_app = None

            if progress_callback:
                progress_callback(10, "Docx engine initialized successfully")
            
            # Determine which file to open:
            # If this is not the first rule and processed_file exists, use it
            # Otherwise use the original file
            if document.processed_file and os.path.exists(document.processed_file.path):
                file_to_process = document.processed_file.path
                logger.info(f"Processing existing processed file: {file_to_process}")
            else:
                file_to_process = document.original_file.path
                logger.info(f"Processing original file: {file_to_process}")
            
            # Open document using engine
            if progress_callback:
                progress_callback(15, "Opening document...")

            doc = self.engine.open_document(Path(file_to_process))
            
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
                
                # Save processed DOCX using engine
                if progress_callback:
                    progress_callback(75, "Saving processed DOCX...")

                try:
                    self.engine.save_as_new_docx(doc, docx_path)
                except Exception:
                    # As a fallback, if engine provides a direct save method on doc
                    try:
                        doc.save(str(docx_path))
                    except Exception as e:
                        raise

                # Update document record to point to processed file
                document.processed_file = f"processed/{docx_name}"
                document.save()

                if progress_callback:
                    progress_callback(80, "Converting to PDF...")

                pdf_name = f"{Path(document.original_file.name).stem}_processed.pdf"
                pdf_path = pdf_dir / pdf_name

                # Convert DOCX to PDF using LibreOffice headless
                try:
                    _convert_docx_to_pdf(docx_path, pdf_path)
                except Exception as e:
                    logger.warning(f"PDF conversion failed: {e}")
                    raise

                # Update document record with both files
                if progress_callback:
                    progress_callback(95, "Finalizing...")

                document.pdf_file = f"processed/{pdf_name}"  # PDF
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
                    if HAVE_WIN32 and win32com:
                        win32com.client.Dispatch("WScript.Shell").Run("taskkill /f /im WINWORD.EXE", 0, True)
                except Exception:
                    pass
                    
            # COM cleanup is handled by the background task
    def _initialize_word_engine(self):
        """Initialize Word COM engine"""
        # Initialize Word (deferred import)
        if not HAVE_WIN32:
            raise RuntimeError("Word COM not available on this platform")
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = False
        return engines.WordComEngine()
    