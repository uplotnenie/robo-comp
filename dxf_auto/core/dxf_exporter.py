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
import tempfile
import shutil

from .kompas_api import (
    KompasAPI,
    KompasDocument,
    DocumentType,
    KompasCommand,
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
        
        For 3D documents with sheet metal bodies, creates an associative view
        in a 2D fragment to project the flat pattern, then saves as DXF.
        
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
            
            # Method 2: For 3D documents, use 2D fragment with associative view
            logger.debug("Document is 3D, creating 2D projection for DXF export")
            return self._export_via_2d_fragment(doc, output_path)
            
        except Exception as e:
            logger.error(f"DXF export error: {e}")
            return False
    
    def _export_via_2d_fragment(self, doc: KompasDocument, output_path: Path) -> bool:
        """
        Export 3D flat pattern via 2D fragment.
        
        This method creates a 2D projection of the straightened sheet metal body.
        
        CRITICAL SEQUENCE for command 40373 (ksCMCreateSheetFromModel):
        1. The 3D document MUST be active and visible
        2. The sheet metal body MUST be straightened (unfolded)
        3. Create a new 2D fragment (this becomes the "target" for the command)
        4. Re-activate the 3D document 
        5. Execute command 40373 - this projects the flat pattern into the last created 2D doc
        6. Activate and save the fragment as DXF
        
        Args:
            doc: 3D document with straightened flat pattern (already straightened)
            output_path: Output DXF file path
            
        Returns:
            True if successful
        """
        fragment = None
        
        try:
            source_path = doc.path_name
            logger.debug(f"Exporting flat pattern from: {source_path}")
            
            # CRITICAL: Ensure 3D document is active and visible
            doc.activate()
            logger.debug("Activated 3D document")
            
            # Step 1: Create new 2D fragment (visible so it can receive geometry)
            fragment = self._api.documents.add(DocumentType.FRAGMENT, visible=True)
            if fragment is None:
                logger.error("Failed to create 2D fragment")
                return False
            
            logger.debug("Created 2D fragment for export")
            
            # Step 2: Try command-based approach FIRST (specifically designed for flat patterns)
            # Re-activate the 3D document - the command works on the ACTIVE 3D model
            doc.activate()
            logger.debug("Re-activated 3D document for command execution")
            
            # Add small delay to ensure KOMPAS UI is ready
            import time
            time.sleep(0.3)
            
            # Execute the command to create 2D sketch from the 3D model
            success = self._execute_create_sheet_from_model(doc, fragment, output_path)
            
            if success:
                return True
            
            # Step 3: Fallback to associative view approach
            logger.info("Command-based export failed, trying associative view fallback")
            
            # For associative view, we need to save the 3D model first
            temp_dir = tempfile.mkdtemp(prefix="kompas_dxf_")
            original_name = Path(source_path).name if source_path else "temp_model.m3d"
            temp_model_path = Path(temp_dir) / original_name
            
            doc.activate()
            if doc.save_as(str(temp_model_path)):
                logger.debug(f"Saved straightened model to: {temp_model_path}")
                success = self._export_via_associative_view(fragment, temp_model_path, output_path)
                
                # Cleanup temp file
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except:
                    pass
                    
                if success:
                    return True
            
            logger.error("All export methods failed")
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
    
    def _execute_create_sheet_from_model(
        self, 
        doc: KompasDocument, 
        fragment: KompasDocument,
        output_path: Path
    ) -> bool:
        """
        Execute ksCMCreateSheetFromModel command to create 2D geometry from 3D model.
        
        This command creates a 2D sketch/geometry from the active 3D model's flat pattern
        into the last created 2D document (fragment).
        
        Args:
            doc: 3D document (must be active)
            fragment: 2D fragment to receive geometry (already created)
            output_path: Output DXF file path
            
        Returns:
            True if successful
        """
        try:
            # Check if command is available
            if not self._api.is_command_available(KompasCommand.CREATE_SHEET_FROM_MODEL):
                logger.warning("CreateSheetFromModel command not available")
                return False
            
            logger.debug("Executing CreateSheetFromModel command (40373)")
            
            # Execute the command
            # The command reads from the ACTIVE 3D document and writes to the LAST CREATED 2D document
            result = self._api.execute_command(KompasCommand.CREATE_SHEET_FROM_MODEL, False)
            
            if not result:
                logger.warning("CreateSheetFromModel command returned False")
                # Don't return yet - the command might have worked anyway
            
            # Add delay to ensure command completes
            import time
            time.sleep(0.5)
            
            # Activate the fragment to check if it has content and save it
            fragment.activate()
            
            # Try to save as DXF
            success = fragment.save_as(str(output_path))
            
            if success:
                # Verify the file was created and has some content
                if output_path.exists() and output_path.stat().st_size > 100:
                    logger.info(f"Successfully exported DXF via command: {output_path}")
                    return True
                else:
                    logger.warning("DXF file created but appears to be empty or too small")
                    return False
            else:
                logger.warning("Failed to save fragment as DXF")
                return False
                
        except Exception as e:
            logger.error(f"Command-based export error: {e}")
            return False
    
    def _export_via_associative_view(
        self,
        fragment: KompasDocument,
        source_model_path: Path,
        output_path: Path
    ) -> bool:
        """
        Export using IAssociativeView API.
        
        This creates a 2D projection view from the 3D model file.
        The source model must be saved with the flat pattern (straightened) state.
        
        Args:
            fragment: 2D fragment document (already created)
            source_model_path: Path to the saved 3D model with flat pattern
            output_path: Output DXF file path
            
        Returns:
            True if successful
        """
        try:
            # Get 2D document interface
            doc_2d = fragment.get_2d_document()
            if doc_2d is None:
                logger.error("Failed to get 2D document interface")
                return False
            
            # Get views and layers manager
            vl_manager = doc_2d.views_and_layers_manager
            if vl_manager is None:
                logger.error("Failed to get ViewsAndLayersManager")
                return False
            
            # Get views collection
            views_collection = vl_manager.views_collection
            if views_collection is None:
                logger.error("Failed to get views collection")
                return False
            
            # Add associative view
            assoc_view = views_collection.add_associative_view()
            if assoc_view is None:
                logger.error("Failed to add associative view")
                return False
            
            logger.debug("Created associative view")
            
            # Configure the view
            # Set source 3D model file
            if not assoc_view.set_source_file(str(source_model_path)):
                logger.error("Failed to set source file for associative view")
                return False
            
            # Set projection to Top view (best for flat patterns)
            # Try different projection names (Russian and English)
            projection_names = ["Сверху", "#Сверху", "Top", "#Top", "XY"]
            projection_set = False
            for proj_name in projection_names:
                if assoc_view.set_projection(proj_name):
                    projection_set = True
                    logger.debug(f"Set projection: {proj_name}")
                    break
            
            if not projection_set:
                logger.warning("Could not set projection, using default")
            
            # Set view parameters
            assoc_view.set_position(100.0, 100.0)  # Center position in fragment
            assoc_view.set_scale(1.0)  # 1:1 scale
            assoc_view.set_angle(0.0)  # No rotation
            assoc_view.set_hidden_lines(False)  # No hidden lines for flat pattern
            assoc_view.set_tangent_edges(False)  # No tangent edges
            
            # Update the view to generate geometry
            if not assoc_view.update():
                logger.error("Failed to update associative view")
                return False
            
            logger.debug("Associative view updated successfully")
            
            # Activate fragment and save as DXF
            fragment.activate()
            
            # Small delay to ensure view is fully rendered
            import time
            time.sleep(0.5)
            
            success = fragment.save_as(str(output_path))
            if success:
                logger.info(f"Successfully exported DXF via associative view: {output_path}")
            return success
            
        except Exception as e:
            logger.error(f"Associative view export error: {e}")
            return False
    
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
