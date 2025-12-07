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
                    # Unfold (straighten) the body if needed
                    if self._settings.straighten_before_export and not body.is_straightened:
                        logger.debug("Straightening sheet metal body")
                        body.is_straightened = True
                        # Rebuild to apply changes
                        doc.rebuild()
                    
                    # Export to DXF
                    success = self._export_to_dxf(doc, output_path)
                    
                    if success:
                        result.success = True
                        logger.info(f"Successfully exported: {output_path}")
                    else:
                        result.error_message = "Ошибка конвертации в DXF"
                    
                finally:
                    # Restore original state
                    if body.is_straightened != original_straightened:
                        logger.debug("Restoring original straightened state")
                        body.is_straightened = original_straightened
                    
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
        Export 3D flat pattern via 2D fragment using ksCMCreateSheetFromModel command.
        
        Process:
        1. Ensure 3D document is active and rebuilt
        2. Create a new 2D fragment (this will be the target for 2D geometry)
        3. Execute ksCMCreateSheetFromModel command (with 3D doc active)
        4. Save fragment as DXF
        5. Close fragment
        
        Args:
            doc: 3D document with straightened flat pattern
            output_path: Output file path
            
        Returns:
            True if successful
        """
        fragment = None
        try:
            # Ensure 3D document is active
            doc.activate()
            
            # Rebuild to ensure straightened state is applied
            if not doc.rebuild():
                logger.debug("Rebuild returned false, continuing anyway")
            
            # Get the 3D document path for logging
            source_path = doc.path_name
            logger.debug(f"Exporting flat pattern from: {source_path}")
            
            # Create new 2D fragment (visible for proper export)
            # This fragment will receive the 2D geometry from the command
            fragment = self._api.documents.add(DocumentType.FRAGMENT, visible=True)
            if fragment is None:
                logger.error("Failed to create 2D fragment")
                return False
            
            logger.debug(f"Created 2D fragment for export")
            
            # Use command-based export
            # The command projects the active 3D model into the last created 2D document
            return self._export_via_command(doc, fragment, output_path)
            
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
    
    def _export_via_command(
        self, 
        doc: KompasDocument, 
        fragment: KompasDocument, 
        output_path: Path
    ) -> bool:
        """
        Export using KOMPAS command.
        
        This method uses the ksCMCreateSheetFromModel command to create
        2D geometry from the active 3D model into the last created 2D fragment.
        
        IMPORTANT: The 3D document MUST be active when executing the command.
        The command projects the 3D model into the most recently created 2D document.
        
        Args:
            doc: 3D document (must be visible, will be activated)
            fragment: 2D fragment to receive geometry (already created)
            output_path: Output file path
            
        Returns:
            True if successful
        """
        try:
            # CRITICAL: Activate the 3D document BEFORE executing command
            # The command works on the ACTIVE 3D model
            doc.activate()
            logger.debug(f"Activated 3D document: {doc.name}")
            
            # Check if command is available (3D doc must be active for this check)
            if not self._api.is_command_available(KompasCommand.CREATE_SHEET_FROM_MODEL):
                logger.warning("CreateSheetFromModel command not available (is 3D doc active?)")
                return False
            
            # Execute command WHILE 3D document is still active
            # The command creates 2D sketch from active 3D model into the last created 2D document
            result = self._api.execute_command(KompasCommand.CREATE_SHEET_FROM_MODEL, False)
            
            if not result:
                logger.warning("CreateSheetFromModel command may have failed")
            
            # NOW activate the fragment to save it
            fragment.activate()
            logger.debug(f"Activated fragment for save")
            
            # Save as DXF
            success = fragment.save_as(str(output_path))
            if success:
                logger.info(f"Successfully exported DXF: {output_path}")
            return success
            
        except Exception as e:
            logger.error(f"Command export error: {e}")
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
