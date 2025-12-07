"""
Assembly Scanner Module

Scans KOMPAS-3D assemblies to find sheet metal parts.
Builds a tree structure of the assembly composition.
"""

from typing import List, Optional, Callable, Any
from dataclasses import dataclass
from pathlib import Path
import logging
import uuid

from .kompas_api import (
    KompasAPI, 
    KompasDocument, 
    KompasDocument3D, 
    Part3D,
    SheetMetalContainer,
    SheetMetalBody,
)
from models.sheet_part import SheetPartInfo, AssemblyNode

logger = logging.getLogger(__name__)


@dataclass
class ScanProgress:
    """Progress information for assembly scanning."""
    current: int = 0
    total: int = 0
    current_part: str = ""
    message: str = ""
    
    @property
    def percentage(self) -> float:
        if self.total == 0:
            return 0.0
        return (self.current / self.total) * 100


class AssemblyScanner:
    """
    Scans KOMPAS-3D assemblies to extract sheet metal parts.
    
    Usage:
        scanner = AssemblyScanner(kompas_api)
        tree = scanner.scan_assembly(document)
        sheet_parts = tree.get_all_sheet_parts()
    """
    
    def __init__(self, api: KompasAPI):
        """
        Initialize scanner.
        
        Args:
            api: KompasAPI instance
        """
        self._api = api
        self._progress_callback: Optional[Callable[[ScanProgress], None]] = None
        self._scanned_files: set = set()  # Track scanned files to avoid duplicates
    
    def set_progress_callback(self, callback: Callable[[ScanProgress], None]):
        """Set callback for progress updates."""
        self._progress_callback = callback
    
    def _report_progress(self, progress: ScanProgress):
        """Report progress to callback if set."""
        if self._progress_callback:
            self._progress_callback(progress)
    
    def scan_active_document(self) -> Optional[AssemblyNode]:
        """
        Scan the currently active document.
        
        Returns:
            AssemblyNode tree or None if no document is open
        """
        doc = self._api.active_document
        if doc is None:
            logger.warning("No active document")
            return None
        
        return self.scan_document(doc)
    
    def scan_document(self, document: KompasDocument) -> Optional[AssemblyNode]:
        """
        Scan a KOMPAS document for sheet metal parts.
        
        Args:
            document: KompasDocument to scan
            
        Returns:
            AssemblyNode tree or None if not a 3D document
        """
        if not document.is_3d:
            logger.warning(f"Document {document.name} is not a 3D model")
            return None
        
        # Get 3D document interface
        doc_3d = document.get_3d_document()
        if doc_3d is None:
            logger.error("Failed to get 3D document interface")
            return None
        
        # Get top part
        top_part = doc_3d.top_part
        if top_part is None:
            logger.error("Failed to get top part")
            return None
        
        # Clear scanned files set
        self._scanned_files.clear()
        
        # Create progress tracker
        progress = ScanProgress(message="Подсчет компонентов...")
        self._report_progress(progress)
        
        # Count total parts for progress
        total_parts = self._count_parts(top_part)
        progress.total = total_parts
        
        # Build tree
        progress.message = "Сканирование сборки..."
        self._report_progress(progress)
        
        tree = self._scan_part_recursive(
            part=top_part,
            progress=progress,
            level=0,
            parent_id=""
        )
        
        progress.message = "Сканирование завершено"
        progress.current = progress.total
        self._report_progress(progress)
        
        return tree
    
    def _count_parts(self, part: Part3D) -> int:
        """
        Recursively count total parts in hierarchy.
        
        Args:
            part: Root part to count from
            
        Returns:
            Total number of parts
        """
        count = 1
        for child in part.parts:
            count += self._count_parts(child)
        return count
    
    def _scan_part_recursive(
        self, 
        part: Part3D, 
        progress: ScanProgress,
        level: int = 0,
        parent_id: str = ""
    ) -> AssemblyNode:
        """
        Recursively scan a part and its children.
        
        Args:
            part: Part3D to scan
            progress: Progress tracker
            level: Current nesting level
            parent_id: ID of parent node
            
        Returns:
            AssemblyNode for this part
        """
        # Generate unique ID
        node_id = str(uuid.uuid4())[:8]
        
        # Get part info
        file_path = part.file_name
        designation = part.marking
        name = part.name
        
        # Update progress
        progress.current += 1
        progress.current_part = designation or name or file_path
        self._report_progress(progress)
        
        # Check for sheet metal
        is_sheet_metal = False
        sheet_part_info: Optional[SheetPartInfo] = None
        
        # Log for debugging
        is_detail = part.is_detail
        logger.debug(f"Checking part: '{name}' designation='{designation}' file='{file_path}'")
        logger.debug(f"  is_detail={is_detail}, is_standard={part.is_standard}")
        
        # Check for sheet metal regardless of is_detail to debug the issue
        # In theory, only details should have sheet metal bodies, but let's check all parts
        container = part.get_sheet_metal_container()
        logger.debug(f"  Sheet metal container: {container is not None}")
        if container:
            has_sm = container.has_sheet_metal
            logger.debug(f"  Has sheet metal: {has_sm}")
            if has_sm:
                is_sheet_metal = True
                
                # Get sheet metal body info
                bodies = container.sheet_metal_bodies
                logger.debug(f"  Sheet metal bodies count: {len(bodies)}")
                if bodies:
                    body = bodies[0]  # Usually one main sheet body
                    sheet_part_info = self._create_sheet_part_info(
                        part=part,
                        body=body,
                        level=level,
                        parent_id=parent_id
                    )
                    sheet_part_info.id = node_id
                    logger.info(f"Found sheet metal part: {name or designation}")
        
        # Create node
        node = AssemblyNode(
            id=node_id,
            name=name,
            designation=designation,
            file_path=file_path,
            is_assembly=not part.is_detail,
            is_sheet_metal=is_sheet_metal,
            is_standard=part.is_standard,
            parent_id=parent_id,
            level=level,
            quantity=part.instance_count,
            sheet_part=sheet_part_info,
        )
        
        # Store COM reference
        node._com_object = part.raw
        
        # Recursively scan children
        for child_part in part.parts:
            child_node = self._scan_part_recursive(
                part=child_part,
                progress=progress,
                level=level + 1,
                parent_id=node_id
            )
            node.add_child(child_node)
        
        return node
    
    def _create_sheet_part_info(
        self, 
        part: Part3D, 
        body: SheetMetalBody,
        level: int,
        parent_id: str
    ) -> SheetPartInfo:
        """
        Create SheetPartInfo from part and body data.
        
        Args:
            part: Part3D object
            body: SheetMetalBody object
            level: Nesting level
            parent_id: Parent node ID
            
        Returns:
            SheetPartInfo object
        """
        file_path = part.file_name
        file_name = Path(file_path).stem if file_path else ""
        
        info = SheetPartInfo(
            designation=part.marking,
            name=part.name,
            file_path=file_path,
            file_name=file_name,
            material=part.material,
            thickness=body.thickness,
            quantity=part.instance_count,
            parent_id=parent_id,
            level=level,
            mass=part.mass,
        )
        
        return info
    
    def get_sheet_parts_flat(self, tree: AssemblyNode) -> List[SheetPartInfo]:
        """
        Get flat list of all sheet metal parts from tree.
        
        Args:
            tree: AssemblyNode tree root
            
        Returns:
            List of SheetPartInfo objects
        """
        return tree.get_all_sheet_parts()
    
    def rescan_part(self, node: AssemblyNode) -> Optional[SheetPartInfo]:
        """
        Rescan a single part to refresh its properties.
        
        Args:
            node: AssemblyNode to rescan
            
        Returns:
            Updated SheetPartInfo or None
        """
        if node._com_object is None:
            logger.warning("No COM reference for part")
            return None
        
        try:
            part = Part3D(node._com_object)
            container = part.get_sheet_metal_container()
            
            if container and container.has_sheet_metal:
                bodies = container.sheet_metal_bodies
                if bodies:
                    return self._create_sheet_part_info(
                        part=part,
                        body=bodies[0],
                        level=node.level,
                        parent_id=node.parent_id
                    )
        except Exception as e:
            logger.error(f"Failed to rescan part: {e}")
        
        return None


def filter_sheet_parts(
    parts: List[SheetPartInfo],
    include_hidden: bool = False,
    include_standard: bool = False,
    min_thickness: float = 0.0,
    max_thickness: float = float('inf'),
    material_filter: str = ""
) -> List[SheetPartInfo]:
    """
    Filter sheet parts based on criteria.
    
    Args:
        parts: List of SheetPartInfo to filter
        include_hidden: Include hidden/suppressed parts
        include_standard: Include standard parts
        min_thickness: Minimum thickness filter
        max_thickness: Maximum thickness filter
        material_filter: Material name filter (substring match)
        
    Returns:
        Filtered list of SheetPartInfo
    """
    result = []
    
    for part in parts:
        # Skip based on selection
        if not part.export_selected:
            continue
        
        # Thickness filter
        if part.thickness < min_thickness or part.thickness > max_thickness:
            continue
        
        # Material filter
        if material_filter and material_filter.lower() not in part.material.lower():
            continue
        
        result.append(part)
    
    return result


def group_by_material(parts: List[SheetPartInfo]) -> dict:
    """
    Group sheet parts by material.
    
    Args:
        parts: List of SheetPartInfo
        
    Returns:
        Dictionary mapping material name to list of parts
    """
    groups = {}
    for part in parts:
        material = part.material or "Без материала"
        if material not in groups:
            groups[material] = []
        groups[material].append(part)
    return groups


def group_by_thickness(parts: List[SheetPartInfo]) -> dict:
    """
    Group sheet parts by thickness.
    
    Args:
        parts: List of SheetPartInfo
        
    Returns:
        Dictionary mapping thickness to list of parts
    """
    groups = {}
    for part in parts:
        thickness = f"{part.thickness:.1f}"
        if thickness not in groups:
            groups[thickness] = []
        groups[thickness].append(part)
    return groups
