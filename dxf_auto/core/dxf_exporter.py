"""
DXF Exporter Module

Handles the export of sheet metal flat patterns to DXF format.
Includes unfold generation, layer management, and file export.
"""

from typing import List, Optional, Callable, Any, Dict
from dataclasses import dataclass, field
from pathlib import Path
from datetime import datetime
import logging
import os
import threading
import time
import tempfile
import shutil

from .kompas_api import (
    KompasAPI,
    KompasDocument,
    DocumentType,
    Part3D,
)
from models.sheet_part import SheetPartInfo, SheetPart
from models.export_settings import ExportSettings, LineTypeSettings

logger = logging.getLogger(__name__)


@dataclass
class ExportResult:
    """Result of a single DXF export operation."""
    
    part_info: SheetPartInfo
    success: bool = False
    output_path: str = ""
    error_message: str = ""
    warnings: List[str] = field(default_factory=list)
    
    # Timing
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    
    @property
    def duration_seconds(self) -> float:
        if self.start_time and self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return 0.0


@dataclass
class ExportProgress:
    """Progress information for export operation."""
    
    current: int = 0
    total: int = 0
    current_part: str = ""
    current_status: str = ""
    message: str = ""
    
    @property
    def percentage(self) -> float:
        if self.total == 0:
            return 0.0
        return (self.current / self.total) * 100


@dataclass
class ExportSummary:
    """Summary of a batch export operation."""
    
    results: List[ExportResult] = field(default_factory=list)
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    
    @property
    def total_count(self) -> int:
        return len(self.results)
    
    @property
    def success_count(self) -> int:
        return sum(1 for r in self.results if r.success)
    
    @property
    def failure_count(self) -> int:
        return sum(1 for r in self.results if not r.success)
    
    @property
    def success_rate(self) -> float:
        if self.total_count == 0:
            return 0.0
        return (self.success_count / self.total_count) * 100
    
    @property
    def duration_seconds(self) -> float:
        if self.start_time and self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return 0.0
    
    def get_failed_results(self) -> List[ExportResult]:
        return [r for r in self.results if not r.success]
    
    def get_successful_results(self) -> List[ExportResult]:
        return [r for r in self.results if r.success]


class DXFExporter:
    """
    Exports sheet metal flat patterns to DXF format.
    
    Process for each part:
    1. Activate/open the part document
    2. Get sheet metal body
    3. Create flat pattern (unfold/straighten)
    4. Create 2D fragment from pattern
    5. Configure layers according to settings
    6. Export to DXF
    7. Restore original state
    
    Usage:
        exporter = DXFExporter(kompas_api, settings)
        summary = exporter.export_parts(sheet_parts)
    """
    
    # KOMPAS converter library for DXF (standard location)
    DXF_CONVERTER_LIBRARY = ""  # Empty string uses default
    
    def __init__(self, api: KompasAPI, settings: ExportSettings):
        """
        Initialize exporter.
        
        Args:
            api: KompasAPI instance
            settings: Export settings
        """
        self._api = api
        self._settings = settings
        self._progress_callback: Optional[Callable[[ExportProgress], None]] = None
        self._cancel_requested = False
    
    def set_progress_callback(self, callback: Callable[[ExportProgress], None]):
        """Set callback for progress updates."""
        self._progress_callback = callback
    
    def request_cancel(self):
        """Request cancellation of export operation."""
        self._cancel_requested = True
    
    def _report_progress(self, progress: ExportProgress):
        """Report progress to callback if set."""
        if self._progress_callback:
            self._progress_callback(progress)
    
    def export_parts(self, parts: List[SheetPartInfo]) -> ExportSummary:
        """
        Export multiple sheet metal parts to DXF.
        
        Args:
            parts: List of SheetPartInfo to export
            
        Returns:
            ExportSummary with results
        """
        self._cancel_requested = False
        summary = ExportSummary(start_time=datetime.now())
        
        # Filter to only selected parts
        export_parts = [p for p in parts if p.export_selected]
        
        progress = ExportProgress(
            total=len(export_parts),
            message="Начало экспорта..."
        )
        self._report_progress(progress)
        
        # Ensure output directory exists
        self._ensure_output_directory()
        
        # Export each part
        for index, part_info in enumerate(export_parts):
            if self._cancel_requested:
                logger.info("Export cancelled by user")
                break
            
            progress.current = index + 1
            progress.current_part = part_info.display_name
            progress.current_status = "Экспорт..."
            self._report_progress(progress)
            
            result = self._export_single_part(part_info, index)
            summary.results.append(result)
            
            # Update part info with result
            part_info.export_path = result.output_path
            part_info.export_status = "OK" if result.success else result.error_message
        
        summary.end_time = datetime.now()
        
        progress.message = f"Экспорт завершен: {summary.success_count}/{summary.total_count}"
        progress.current = progress.total
        self._report_progress(progress)
        
        return summary
    
    def _export_single_part(self, part_info: SheetPartInfo, index: int) -> ExportResult:
        """
        Export a single part to DXF.
        
        Args:
            part_info: Part information
            index: Export index (for filename)
            
        Returns:
            ExportResult
        """
        result = ExportResult(
            part_info=part_info,
            start_time=datetime.now()
        )
        
        doc = None
        
        try:
            # Generate output path
            output_path = self._generate_output_path(part_info, index)
            result.output_path = str(output_path)
            
            # Check if file exists
            if output_path.exists() and not self._settings.overwrite_existing:
                result.error_message = "Файл уже существует"
                result.end_time = datetime.now()
                return result
            
            # Ensure output directory exists
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Open part document (visible for proper export)
            doc = self._open_part_document(part_info)
            if doc is None:
                result.error_message = "Не удалось открыть документ детали"
                result.end_time = datetime.now()
                return result
            
            try:
                # Activate the document
                doc.activate()
                
                # Get 3D document
                doc_3d = doc.get_3d_document()
                if doc_3d is None:
                    result.error_message = "Не удалось получить 3D документ"
                    return result
                
                # Get top part
                top_part = doc_3d.top_part
                if top_part is None:
                    result.error_message = "Не удалось получить компонент"
                    return result
                
                # Get sheet metal container
                container = top_part.get_sheet_metal_container()
                if container is None or not container.has_sheet_metal:
                    result.error_message = "Деталь не является листовой"
                    return result
                
                # Get sheet metal body
                bodies = container.sheet_metal_bodies
                if not bodies:
                    result.error_message = "Листовое тело не найдено"
                    return result
                
                body = bodies[0]
                original_straightened = body.is_straightened
                
                logger.debug(f"Sheet metal body found, original straightened state: {original_straightened}")
                
                try:
                    # CRITICAL: Activate the 3D document before any modification
                    doc.activate()
                    
                    # Unfold (straighten) the body if needed
                    if self._settings.straighten_before_export and not original_straightened:
                        logger.info("Straightening sheet metal body...")
                        
                        # Use the new straighten() method with verification
                        if not body.straighten():
                            logger.warning("straighten() method returned False, trying direct property set")
                            body.is_straightened = True
                        
                        # Verify the change actually happened
                        if not body.is_straightened:
                            logger.error("Sheet metal body failed to straighten - state is still False")
                            result.error_message = "Не удалось развернуть листовое тело"
                            return result
                        
                        logger.info("Sheet metal body straightened successfully")
                        
                        # Rebuild to apply changes - activate first
                        doc.activate()
                        rebuild_result = doc.rebuild()
                        if not rebuild_result:
                            logger.warning("Rebuild returned False, but continuing with export")
                        else:
                            logger.debug("Rebuild completed successfully")
                    
                    # Export to DXF
                    success = self._export_to_dxf(doc, output_path)
                    
                    if success:
                        result.success = True
                        logger.info(f"Successfully exported: {output_path}")
                    else:
                        result.error_message = "Ошибка конвертации в DXF"
                    
                finally:
                    # Restore original state if we changed it
                    if body.is_straightened != original_straightened:
                        logger.debug(f"Restoring original straightened state: {original_straightened}")
                        if original_straightened:
                            body.straighten()
                        else:
                            body.fold()
                    
            except Exception as e:
                logger.exception(f"Error during export of {part_info.display_name}")
                result.error_message = str(e)
                
        except Exception as e:
            logger.exception(f"Error exporting part {part_info.display_name}")
            result.error_message = str(e)
            
        finally:
            # Close document after export
            if doc is not None:
                try:
                    doc.close(save=False)
                except Exception as e:
                    logger.debug(f"Failed to close document: {e}")
        
        result.end_time = datetime.now()
        return result
    
    def _open_part_document(self, part_info: SheetPartInfo) -> Optional[KompasDocument]:
        """
        Open the part document.
        
        Args:
            part_info: Part information with file path
            
        Returns:
            KompasDocument or None
        """
        if not part_info.file_path:
            return None
        
        # Check if file exists
        if not Path(part_info.file_path).exists():
            logger.error(f"Part file not found: {part_info.file_path}")
            return None
        
        # Try to open document - VISIBLE for proper export
        return self._api.documents.open(
            path=part_info.file_path,
            visible=True,  # Must be visible for proper flat pattern export
            read_only=False  # Need write access to modify straighten state
        )
    
    def _export_to_dxf(self, doc: KompasDocument, output_path: Path) -> bool:
        """
        Export document to DXF format.
        
        For 3D documents with sheet metal bodies, uses IConverter for direct
        conversion or creates an associative view in a 2D fragment.
        
        Args:
            doc: Document to export (should be 3D with straightened sheet metal)
            output_path: Output file path
            
        Returns:
            True if successful
        """
        try:
            # Method 1: For 2D documents, use SaveAs directly
            if doc.is_2d:
                logger.debug("Document is 2D, saving directly as DXF")
                return doc.save_as(str(output_path))
            
            # Method 2: For 3D documents, try IConverter first (more reliable)
            logger.debug("Document is 3D, trying IConverter approach")
            if self._export_via_converter(doc, output_path):
                return True
            
            # Method 3: Try Drawing with Associative View
            logger.debug("IConverter failed, trying Drawing with AssociativeView")
            if self._export_via_drawing_view(doc, output_path):
                return True
            
            # Method 4: Fallback to 2D fragment with command 40373
            logger.debug("AssociativeView failed, falling back to 2D fragment method")
            return self._export_via_2d_fragment(doc, output_path)
            
        except Exception as e:
            logger.error(f"DXF export error: {e}")
            return False
    
    def _export_via_converter(self, doc: KompasDocument, output_path: Path) -> bool:
        """
        Export 3D flat pattern to DXF using IConverter interface.
        
        This method:
        1. Saves the straightened 3D model to a temporary file
        2. Uses IConverter to convert the temp file to DXF
        3. Cleans up the temporary file
        
        This is more reliable than the interactive command approach because
        IConverter works non-interactively and doesn't require user input.
        
        Args:
            doc: 3D document with straightened flat pattern
            output_path: Output DXF file path
            
        Returns:
            True if successful
        """
        temp_file = None
        try:
            source_path = doc.path_name
            logger.debug(f"Exporting via IConverter from: {source_path}")
            
            # Step 1: Get the converter interface
            converter = self._api.get_converter()
            if converter is None:
                logger.warning("IConverter not available")
                return False
            
            logger.debug("Got IConverter interface")
            
            # Step 2: Create a temporary file path for the straightened model
            # We need to save the current state (with flat pattern visible)
            temp_dir = tempfile.gettempdir()
            temp_name = f"dxf_export_temp_{os.getpid()}_{int(time.time())}.m3d"
            temp_file = Path(temp_dir) / temp_name
            
            logger.debug(f"Saving temporary model to: {temp_file}")
            
            # Step 3: Save the document with flat pattern to temp file
            # This preserves the straightened state in the temp file
            if not doc.save_as(str(temp_file)):
                logger.warning("Failed to save temporary model file")
                return False
            
            logger.debug("Saved temporary model with flat pattern")
            
            # Step 4: Use IConverter to convert temp .m3d to .dxf
            # Parameters: inputFile, outputFile, commandCode, showParams
            try:
                logger.debug(f"Converting {temp_file} -> {output_path}")
                result = converter.Convert(
                    str(temp_file),    # Input file (3D model with flat pattern)
                    str(output_path),  # Output file (DXF)
                    0,                 # Command code (0 = default/auto)
                    False              # Don't show parameters dialog
                )
                logger.debug(f"IConverter.Convert returned: {result}")
                
                if result:
                    # Verify the DXF was created
                    if output_path.exists():
                        file_size = output_path.stat().st_size
                        if file_size > 500:  # Minimum size for valid DXF
                            logger.info(f"IConverter export successful: {output_path} ({file_size} bytes)")
                            return True
                        else:
                            logger.warning(f"IConverter created small file ({file_size} bytes)")
                            return False
                    else:
                        logger.warning("IConverter returned True but file doesn't exist")
                        return False
                else:
                    logger.warning("IConverter.Convert returned False")
                    return False
                    
            except Exception as conv_error:
                logger.warning(f"IConverter.Convert failed: {conv_error}")
                return False
            
        except Exception as e:
            logger.error(f"IConverter export error: {e}")
            return False
            
        finally:
            # Step 5: Clean up temporary file
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                    logger.debug("Cleaned up temporary model file")
                except Exception as cleanup_error:
                    logger.debug(f"Failed to clean up temp file: {cleanup_error}")
        
        return False  # Should not reach here
    
    def _export_via_drawing_view(self, doc: KompasDocument, output_path: Path) -> bool:
        """
        Export 3D flat pattern to DXF using Drawing with Associative View.
        
        This method:
        1. Saves the 3D model with flat pattern to a temp file
        2. Creates a new Drawing document (ksDocumentDrawing, type=1)
        3. Adds an associative view from the temp 3D file
        4. Saves the drawing as DXF
        5. Cleans up temporary files
        
        This approach uses IViewsCollection.AddAssociationView() which
        is only available on Drawing documents (not fragments).
        
        Args:
            doc: 3D document with straightened flat pattern
            output_path: Output DXF file path
            
        Returns:
            True if successful
        """
        temp_file = None
        drawing = None
        
        try:
            source_path = doc.path_name
            logger.debug(f"Exporting via Drawing AssociativeView from: {source_path}")
            
            # Step 1: Save the 3D model with flat pattern to temp file
            temp_dir = tempfile.gettempdir()
            temp_name = f"dxf_export_temp_{os.getpid()}_{int(time.time())}.m3d"
            temp_file = Path(temp_dir) / temp_name
            
            logger.debug(f"Saving temporary model to: {temp_file}")
            if not doc.save_as(str(temp_file)):
                logger.warning("Failed to save temporary model file")
                return False
            
            # Step 2: Create a new Drawing document (type=1)
            drawing = self._api.documents.add(DocumentType.DRAWING, visible=False)
            if drawing is None:
                logger.warning("Failed to create Drawing document")
                return False
            
            logger.debug("Created Drawing document for associative view")
            
            # Step 3: Get the 2D document interface and add associative view
            try:
                # Get raw COM object for the drawing
                drawing_2d = drawing._doc  # Access the underlying COM object
                
                # Try to get ViewsAndLayersManager
                vlm = drawing_2d.ViewsAndLayersManager
                if vlm is None:
                    logger.warning("ViewsAndLayersManager not available")
                    return False
                
                # Get Views collection
                views = vlm.Views
                if views is None:
                    logger.warning("Views collection not available")
                    return False
                
                # Add associative view pointing to the temp file
                logger.debug("Adding associative view from temp model...")
                assoc_view = views.AddAssociationView()
                if assoc_view is None:
                    logger.warning("Failed to create AssociationView")
                    return False
                
                # Configure the view
                assoc_view.SourceFileName = str(temp_file)
                assoc_view.ProjectionName = "Top"  # Top view for flat pattern
                assoc_view.X = 0.0
                assoc_view.Y = 0.0
                assoc_view.Scale = 1.0
                assoc_view.Update()
                
                logger.debug("Associative view created and configured")
                
            except AttributeError as attr_err:
                logger.warning(f"Drawing API not available: {attr_err}")
                return False
            except Exception as view_err:
                logger.warning(f"Failed to create associative view: {view_err}")
                return False
            
            # Step 4: Save drawing as DXF
            time.sleep(0.3)  # Give KOMPAS time to render the view
            
            if drawing.save_as(str(output_path)):
                # Verify the DXF was created
                if output_path.exists():
                    file_size = output_path.stat().st_size
                    if file_size > 500:
                        logger.info(f"Drawing export successful: {output_path} ({file_size} bytes)")
                        return True
                    else:
                        logger.warning(f"Drawing created small DXF ({file_size} bytes)")
                        return False
            
            logger.warning("Failed to save drawing as DXF")
            return False
            
        except Exception as e:
            logger.error(f"Drawing view export error: {e}")
            return False
            
        finally:
            # Clean up
            if drawing is not None:
                try:
                    drawing.close(save=False)
                except:
                    pass
            
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                    logger.debug("Cleaned up temporary model file")
                except:
                    pass
    
    def _export_via_2d_fragment(self, doc: KompasDocument, output_path: Path) -> bool:
        """
        Export 3D flat pattern via 2D fragment using interactive command.
        
        This method uses the KOMPAS command ksCMCreateSheetFromModel (40373)
        which creates a 2D view from a 3D model. The command works as follows:
        - It operates on the currently ACTIVE 3D document
        - It creates geometry in an existing 2D document (fragment)
        - The fragment should be created BEFORE executing the command
        - The 3D document must be ACTIVE when command is executed
        
        The approach:
        1. Activate the 3D document with straightened flat pattern
        2. Create a new 2D fragment document (visible, becomes active)
        3. Re-activate the 3D document
        4. Execute command 40373 (CreateSheetFromModel)
        5. Auto-complete with StopCurrentProcess or Enter key
        6. Activate fragment and save as DXF
        
        Args:
            doc: 3D document with straightened flat pattern (already straightened)
            output_path: Output DXF file path
            
        Returns:
            True if successful
        """
        fragment = None
        
        # Command constant
        CREATE_SHEET_FROM_MODEL = 40373  # ksCMCreateSheetFromModel
        
        try:
            source_path = doc.path_name
            logger.debug(f"Exporting flat pattern from: {source_path}")
            
            # Log document state for debugging
            active_doc = self._api.active_document
            if active_doc:
                logger.debug(f"Currently active document: {active_doc.name}, type: {active_doc.document_type}")
            else:
                logger.debug("No active document")
            
            doc_count = self._api.documents.count
            logger.debug(f"Total open documents: {doc_count}")
            
            # Step 1: Ensure 3D document is active first
            doc.activate()
            logger.debug(f"Activated 3D document: {doc.name}")
            time.sleep(0.3)
            
            # Step 2: Create new 2D fragment document (INVISIBLE for command target)
            # Try creating as invisible - according to reference docs
            fragment = self._api.documents.add(DocumentType.FRAGMENT, visible=False)
            if fragment is None:
                logger.error("Failed to create 2D fragment document")
                return False
            logger.debug(f"Created 2D fragment for flat pattern view: {fragment.name}")
            
            # Step 3: CRITICAL - Keep fragment as target but 3D doc must be active
            # Command 40373 (CreateSheetFromModel) reads geometry from active 3D doc
            # and inserts into the target 2D document
            # The 3D document with the flat pattern must be the ACTIVE document
            time.sleep(0.2)
            doc.activate()  # 3D document must be active - it's the SOURCE
            time.sleep(0.3)
            logger.debug("Ensured 3D document is active (source for flat pattern)")
            
            # Check if command is available before trying to execute
            if self._api.is_command_available(CREATE_SHEET_FROM_MODEL):
                logger.debug("Command 40373 is available")
            else:
                logger.warning("Command 40373 is NOT available - trying alternative approach")
                # Try without threading - maybe the command works differently
                # when not in an interactive context
            
            # Step 4: Set up auto-completion via threading
            # The command 40373 is interactive - it waits for user to place the view
            # We use a timer to call StopCurrentProcess() which completes the command
            
            command_completed = threading.Event()
            command_success = [False]  # Use list to allow modification in nested function
            
            def auto_complete_command():
                """Background function to auto-complete the interactive command."""
                try:
                    # Wait a bit for the command to start and show its dialog/cursor
                    time.sleep(0.5)
                    
                    # Method 1: Call StopCurrentProcess to accept/complete the operation
                    # False = accept (confirm), True = cancel
                    logger.debug("Auto-completing command via StopCurrentProcess(False)...")
                    result = self._api.stop_current_process(cancel=False)
                    logger.debug(f"StopCurrentProcess result: {result}")
                    
                    # Method 2: If StopCurrentProcess didn't work, try sending Enter key
                    if not result:
                        time.sleep(0.3)
                        logger.debug("Trying keyboard Enter simulation...")
                        try:
                            # Try using win32api to send Enter key to KOMPAS window
                            import win32api  # type: ignore
                            import win32con  # type: ignore
                            # VK_RETURN = 0x0D (Enter key)
                            win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                            time.sleep(0.05)
                            win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                            logger.debug("Sent Enter key via win32api")
                        except ImportError:
                            logger.debug("win32api not available for keyboard simulation")
                        except Exception as ke:
                            logger.debug(f"Keyboard simulation failed: {ke}")
                    
                    # Method 3: Try StopCurrentProcess again after delay
                    time.sleep(0.2)
                    self._api.stop_current_process(cancel=False)
                    
                    command_success[0] = True
                except Exception as e:
                    logger.error(f"Auto-complete thread error: {e}")
                finally:
                    command_completed.set()
            
            # Start the auto-complete timer
            timer_thread = threading.Thread(target=auto_complete_command, daemon=True)
            timer_thread.start()
            
            # Step 5: Execute the interactive command
            # Try post=True first - this uses PostMessage which returns immediately
            # and allows the command to run asynchronously
            logger.debug("Executing CreateSheetFromModel command (40373) with post=True...")
            cmd_result = self._api.execute_command(CREATE_SHEET_FROM_MODEL, post=True)
            logger.debug(f"Command execution returned: {cmd_result}")
            
            # If post=True didn't work, try post=False (synchronous)
            if not cmd_result:
                logger.debug("Retrying with post=False (synchronous mode)...")
                time.sleep(0.2)
                cmd_result = self._api.execute_command(CREATE_SHEET_FROM_MODEL, post=False)
                logger.debug(f"Synchronous command returned: {cmd_result}")
            
            # Wait for auto-complete to finish (with timeout)
            command_completed.wait(timeout=5.0)
            
            if not command_completed.is_set():
                logger.warning("Auto-complete thread timed out, command may still be active")
            
            # Give KOMPAS time to process the view creation
            time.sleep(0.5)
            
            # Step 5: Activate fragment and save as DXF
            fragment.activate()
            time.sleep(0.3)
            
            logger.debug(f"Saving fragment as DXF: {output_path}")
            success = fragment.save_as(str(output_path))
            
            if success:
                # Verify the file was created and has content
                if output_path.exists():
                    file_size = output_path.stat().st_size
                    # DXF files with actual geometry should be at least a few KB
                    if file_size > 500:
                        logger.info(f"Successfully exported DXF: {output_path} ({file_size} bytes)")
                        return True
                    else:
                        logger.warning(f"DXF file too small ({file_size} bytes) - view may be empty")
                        # Try alternative: the command may not have worked
                        return False
                else:
                    logger.error("DXF file was not created")
                    return False
            else:
                logger.error("Failed to save fragment as DXF")
                return False
            
        except Exception as e:
            logger.exception(f"Fragment export error: {e}")
            return False
            
        finally:
            # Close fragment without saving (we already saved as DXF)
            if fragment is not None:
                try:
                    fragment.close(save=False)
                except:
                    pass
    
    def _generate_output_path(self, part_info: SheetPartInfo, index: int) -> Path:
        """
        Generate output file path for a part.
        
        Args:
            part_info: Part information
            index: Export index
            
        Returns:
            Path object for output file
        """
        now = datetime.now()
        
        # Build variable dictionary for filename template
        variables = {
            'designation': part_info.designation or '',
            'name': part_info.name or '',
            'material': part_info.material or '',
            'thickness': f"{part_info.thickness:.1f}" if part_info.thickness else '',
            'mass': f"{part_info.mass:.3f}" if part_info.mass else '',
            'filename': part_info.file_name or '',
            'index': str(index + 1),
            'date': now.strftime('%Y-%m-%d'),
            'time': now.strftime('%H-%M-%S'),
        }
        
        return self._settings.get_output_path(
            base_name=part_info.file_name or f"part_{index}",
            variables=variables
        )
    
    def _ensure_output_directory(self):
        """Ensure output directory exists."""
        if self._settings.output_directory:
            output_dir = Path(self._settings.output_directory)
            output_dir.mkdir(parents=True, exist_ok=True)


class DXFPostProcessor:
    """
    Post-processes DXF files.
    
    Can be used to:
    - Modify layer names/colors
    - Remove specific entities
    - Add metadata
    """
    
    def __init__(self, settings: ExportSettings):
        self._settings = settings
    
    def process_file(self, dxf_path: Path) -> bool:
        """
        Post-process a DXF file.
        
        Args:
            dxf_path: Path to DXF file
            
        Returns:
            True if successful
        """
        # This would use a DXF library like ezdxf for modifications
        # For now, return True as placeholder
        logger.info(f"Post-processing: {dxf_path}")
        return True
    
    def apply_layer_settings(self, dxf_path: Path) -> bool:
        """
        Apply layer settings from configuration.
        
        Args:
            dxf_path: Path to DXF file
            
        Returns:
            True if successful
        """
        # TODO: Implement using ezdxf library
        return True


def format_export_report(summary: ExportSummary) -> str:
    """
    Format export summary as text report.
    
    Args:
        summary: Export summary
        
    Returns:
        Formatted report string
    """
    lines = [
        "=" * 60,
        "ОТЧЕТ ОБ ЭКСПОРТЕ DXF",
        "=" * 60,
        f"Время начала: {summary.start_time.strftime('%Y-%m-%d %H:%M:%S') if summary.start_time else 'N/A'}",
        f"Время окончания: {summary.end_time.strftime('%Y-%m-%d %H:%M:%S') if summary.end_time else 'N/A'}",
        f"Длительность: {summary.duration_seconds:.1f} сек",
        "",
        f"Всего файлов: {summary.total_count}",
        f"Успешно: {summary.success_count}",
        f"Ошибок: {summary.failure_count}",
        f"Успешность: {summary.success_rate:.1f}%",
        "",
    ]
    
    if summary.failure_count > 0:
        lines.append("ОШИБКИ:")
        lines.append("-" * 40)
        for result in summary.get_failed_results():
            lines.append(f"  {result.part_info.display_name}: {result.error_message}")
        lines.append("")
    
    if summary.success_count > 0:
        lines.append("УСПЕШНО ЭКСПОРТИРОВАНО:")
        lines.append("-" * 40)
        for result in summary.get_successful_results():
            lines.append(f"  {result.part_info.display_name} -> {Path(result.output_path).name}")
    
    lines.append("=" * 60)
    
    return "\n".join(lines)
