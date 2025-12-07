"""
Sheet Part Data Model

Represents a sheet metal part extracted from a KOMPAS-3D assembly.
"""

from dataclasses import dataclass, field
from typing import Optional, List, Any
from pathlib import Path


@dataclass
class SheetPartInfo:
    """
    Basic information about a sheet metal part.
    
    This is a data transfer object that holds extracted
    properties from a KOMPAS-3D sheet metal part.
    """
    
    # Identification
    id: str = ""                      # Unique ID for this part instance
    designation: str = ""             # Обозначение (marking)
    marking: str = ""                 # Альтернативное обозначение
    name: str = ""                    # Наименование
    
    # Source file
    file_path: str = ""               # Full path to source file
    file_name: str = ""               # Just the filename
    
    # Material properties
    material: str = ""                # Material name
    thickness: float = 0.0            # Sheet thickness in mm
    
    # Quantity
    quantity: int = 1                 # Number of instances in assembly
    
    # Hierarchy
    parent_id: str = ""               # ID of parent assembly/part
    level: int = 0                    # Nesting level in assembly tree
    
    # Unfold dimensions (flat pattern)
    unfold_width: float = 0.0         # Width of unfolded part in mm
    unfold_height: float = 0.0        # Height of unfolded part in mm
    
    # Additional properties (can be extended)
    mass: float = 0.0                 # Mass in kg
    area: float = 0.0                 # Surface area
    
    # Custom properties from model variables
    custom_properties: dict = field(default_factory=dict)
    
    # Export status
    export_selected: bool = True      # Whether to export this part
    export_path: str = ""             # Generated DXF file path
    export_status: str = ""           # Export result status
    
    def __post_init__(self):
        """Generate ID if not provided."""
        if not self.id:
            # Create ID from file path and designation
            self.id = f"{self.file_name}_{self.designation}".replace(" ", "_")
    
    @property
    def display_name(self) -> str:
        """Get display name combining designation and name."""
        if self.designation and self.name:
            return f"{self.designation} - {self.name}"
        return self.designation or self.name or self.file_name or "Unknown"
    
    @property
    def part_id(self) -> str:
        """Get unique part ID (alias for id for consistency with other classes)."""
        return self.id
    
    @property
    def thickness_str(self) -> str:
        """Get thickness as formatted string."""
        if self.thickness > 0:
            return f"{self.thickness:.1f} мм"
        return ""
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            'id': self.id,
            'designation': self.designation,
            'marking': self.marking,
            'name': self.name,
            'file_path': self.file_path,
            'file_name': self.file_name,
            'material': self.material,
            'thickness': self.thickness,
            'quantity': self.quantity,
            'parent_id': self.parent_id,
            'level': self.level,
            'unfold_width': self.unfold_width,
            'unfold_height': self.unfold_height,
            'mass': self.mass,
            'area': self.area,
            'custom_properties': self.custom_properties,
            'export_selected': self.export_selected,
            'export_path': self.export_path,
            'export_status': self.export_status,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'SheetPartInfo':
        """Create from dictionary."""
        return cls(**data)


@dataclass
class SheetPart:
    """
    Sheet metal part with COM object reference.
    
    Combines part information with reference to the actual
    KOMPAS-3D COM object for operations like unfolding.
    """
    
    # Part information
    info: SheetPartInfo
    
    # COM object references (set at runtime)
    _part_object: Any = None          # IPart7 COM object
    _sheet_body_object: Any = None    # ISheetMetalBody COM object
    
    # Thumbnail/preview (PIL Image or bytes)
    thumbnail: Optional[bytes] = None
    unfold_thumbnail: Optional[bytes] = None
    
    @property
    def part_id(self) -> str:
        """Get unique part ID."""
        return self.info.id
    
    @property
    def has_com_reference(self) -> bool:
        """Check if COM objects are available."""
        return self._part_object is not None
    
    def set_com_objects(self, part: Any, sheet_body: Any = None):
        """Set COM object references."""
        self._part_object = part
        self._sheet_body_object = sheet_body
    
    def clear_com_objects(self):
        """Clear COM object references (for serialization)."""
        self._part_object = None
        self._sheet_body_object = None
    
    @property
    def part_object(self) -> Any:
        """Get IPart7 COM object."""
        return self._part_object
    
    @property
    def sheet_body_object(self) -> Any:
        """Get ISheetMetalBody COM object."""
        return self._sheet_body_object


@dataclass
class AssemblyNode:
    """
    Node in the assembly tree structure.
    
    Represents a component in the assembly hierarchy.
    """
    
    # Node identification
    id: str = ""
    name: str = ""
    designation: str = ""
    
    # Source file
    file_path: str = ""
    
    # Node type
    is_assembly: bool = False
    is_sheet_metal: bool = False
    is_standard: bool = False
    
    # Hierarchy
    parent_id: str = ""
    level: int = 0
    children: List['AssemblyNode'] = field(default_factory=list)
    
    # Quantity
    quantity: int = 1
    
    # Reference to sheet part info (if sheet metal)
    sheet_part: Optional[SheetPartInfo] = None
    
    # COM reference
    _com_object: Any = None
    
    @property
    def part_id(self) -> str:
        """Get unique part ID (alias for id)."""
        return self.id
    
    @property
    def display_name(self) -> str:
        """Get display name for tree view."""
        if self.designation and self.name:
            return f"{self.designation} - {self.name}"
        return self.designation or self.name or Path(self.file_path).stem or "Unknown"
    
    @property
    def has_children(self) -> bool:
        """Check if node has child components."""
        return len(self.children) > 0
    
    def add_child(self, child: 'AssemblyNode'):
        """Add a child node."""
        child.parent_id = self.id
        child.level = self.level + 1
        self.children.append(child)
    
    def get_all_sheet_parts(self) -> List[SheetPartInfo]:
        """
        Get all sheet metal parts from this node and descendants.
        
        Returns:
            List of SheetPartInfo objects
        """
        result = []
        
        if self.is_sheet_metal and self.sheet_part:
            result.append(self.sheet_part)
        
        for child in self.children:
            result.extend(child.get_all_sheet_parts())
        
        return result
    
    def flatten(self) -> list:
        """
        Flatten the tree into a list.
        
        Returns:
            List of all nodes in tree order
        """
        result: list = [self]
        for child in self.children:
            result.extend(child.flatten())
        return result
