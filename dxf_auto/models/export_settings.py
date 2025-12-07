"""
Export Settings Data Models

Configuration settings for DXF export process.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional
from pathlib import Path
import json


@dataclass
class LineTypeSettings:
    """
    Settings for a line type in DXF export.
    
    Controls how specific line types (contour, bend lines, etc.)
    are exported to DXF format.
    """
    
    # Identification
    key: str = ""                 # Internal key (e.g., 'contour', 'bend_up')
    name: str = ""                # Display name (e.g., 'Контур')
    description: str = ""         # Description for UI
    
    # Export settings
    enabled: bool = True          # Whether to export this line type
    layer_name: str = ""          # DXF layer name
    
    # DXF line properties
    color: int = 7                # AutoCAD Color Index (ACI) - 7 is white
    line_type: str = "CONTINUOUS" # DXF line type name
    line_weight: int = -1         # Line weight (-1 = default)
    
    def __post_init__(self):
        """Set default layer name from key."""
        if not self.layer_name and self.key:
            self.layer_name = self.key.upper()
    
    def to_dict(self) -> dict:
        """Convert to dictionary."""
        return {
            'key': self.key,
            'name': self.name,
            'description': self.description,
            'enabled': self.enabled,
            'layer_name': self.layer_name,
            'color': self.color,
            'line_type': self.line_type,
            'line_weight': self.line_weight,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'LineTypeSettings':
        """Create from dictionary."""
        return cls(**data)


# AutoCAD Color Index reference
class ACIColors:
    """AutoCAD Color Index (ACI) constants."""
    RED = 1
    YELLOW = 2
    GREEN = 3
    CYAN = 4
    BLUE = 5
    MAGENTA = 6
    WHITE = 7
    DARK_GRAY = 8
    LIGHT_GRAY = 9
    
    @classmethod
    def get_name(cls, index: int) -> str:
        """Get color name by index."""
        names = {
            1: "Красный",
            2: "Желтый", 
            3: "Зеленый",
            4: "Голубой",
            5: "Синий",
            6: "Пурпурный",
            7: "Белый/Черный",
            8: "Темно-серый",
            9: "Светло-серый",
        }
        return names.get(index, f"Цвет {index}")


# DXF Line Types
class DXFLineTypes:
    """Standard DXF line type names."""
    CONTINUOUS = "CONTINUOUS"
    DASHED = "DASHED"
    HIDDEN = "HIDDEN"
    CENTER = "CENTER"
    PHANTOM = "PHANTOM"
    DOT = "DOT"
    DASHDOT = "DASHDOT"
    
    @classmethod
    def get_all(cls) -> List[str]:
        """Get all available line types."""
        return [
            cls.CONTINUOUS,
            cls.DASHED,
            cls.HIDDEN,
            cls.CENTER,
            cls.PHANTOM,
            cls.DOT,
            cls.DASHDOT,
        ]
    
    @classmethod
    def get_display_name(cls, line_type: str) -> str:
        """Get display name for line type."""
        names = {
            "CONTINUOUS": "Сплошная",
            "DASHED": "Штриховая",
            "HIDDEN": "Невидимая",
            "CENTER": "Осевая",
            "PHANTOM": "Пунктирная",
            "DOT": "Точечная",
            "DASHDOT": "Штрихпунктирная",
        }
        return names.get(line_type, line_type)


# Default line type configurations for sheet metal
DEFAULT_LINE_TYPES: Dict[str, LineTypeSettings] = {
    'contour': LineTypeSettings(
        key='contour',
        name='Контур',
        description='Внешний контур развертки',
        enabled=True,
        layer_name='CONTOUR',
        color=ACIColors.WHITE,
        line_type=DXFLineTypes.CONTINUOUS,
    ),
    'bend_up': LineTypeSettings(
        key='bend_up',
        name='Линия сгиба вверх',
        description='Линии сгиба наружу',
        enabled=True,
        layer_name='BEND_UP',
        color=ACIColors.RED,
        line_type=DXFLineTypes.DASHED,
    ),
    'bend_down': LineTypeSettings(
        key='bend_down',
        name='Линия сгиба вниз',
        description='Линии сгиба внутрь',
        enabled=True,
        layer_name='BEND_DOWN',
        color=ACIColors.GREEN,
        line_type=DXFLineTypes.DASHED,
    ),
    'internal_cut': LineTypeSettings(
        key='internal_cut',
        name='Внутренние вырезы',
        description='Внутренние контуры для вырезания',
        enabled=True,
        layer_name='INTERNAL',
        color=ACIColors.CYAN,
        line_type=DXFLineTypes.CONTINUOUS,
    ),
    'engraving': LineTypeSettings(
        key='engraving',
        name='Гравировка',
        description='Линии для гравировки',
        enabled=False,
        layer_name='ENGRAVE',
        color=ACIColors.BLUE,
        line_type=DXFLineTypes.CONTINUOUS,
    ),
}


@dataclass
class FilenameSettings:
    """
    Settings for DXF filename generation.
    
    Uses template with variables that are replaced with
    actual part properties.
    """
    
    # Template string with {variable} placeholders
    template: str = "{designation}_{name}"
    
    # Whether to include file extension in template
    include_extension: bool = True
    
    # Character replacements for invalid filename chars
    replace_invalid: bool = True
    replacement_char: str = "_"
    
    # Invalid characters in filenames
    INVALID_CHARS = '<>:"/\\|?*'
    
    def format(self, variables: Dict[str, str]) -> str:
        """
        Format filename using template and variables.
        
        Args:
            variables: Dictionary of variable values
            
        Returns:
            Formatted filename
        """
        result = self.template
        
        # Replace variables
        for key, value in variables.items():
            placeholder = f"{{{key}}}"
            if placeholder in result:
                result = result.replace(placeholder, str(value))
        
        # Clean up any unreplaced placeholders
        import re
        result = re.sub(r'\{[^}]+\}', '', result)
        
        # Replace invalid characters
        if self.replace_invalid:
            for char in self.INVALID_CHARS:
                result = result.replace(char, self.replacement_char)
        
        # Clean up multiple underscores/dashes
        result = re.sub(r'[_-]+', '_', result)
        result = result.strip('_- ')
        
        # Add extension if needed
        if self.include_extension and not result.lower().endswith('.dxf'):
            result += '.dxf'
        
        return result
    
    def to_dict(self) -> dict:
        """Convert to dictionary."""
        return {
            'template': self.template,
            'include_extension': self.include_extension,
            'replace_invalid': self.replace_invalid,
            'replacement_char': self.replacement_char,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'FilenameSettings':
        """Create from dictionary."""
        return cls(**data)


# Available filename variables
FILENAME_VARIABLES = {
    'designation': ('Обозначение', 'Обозначение детали из свойств'),
    'name': ('Наименование', 'Наименование детали'),
    'material': ('Материал', 'Материал детали'),
    'thickness': ('Толщина', 'Толщина листа в мм'),
    'mass': ('Масса', 'Масса детали в кг'),
    'filename': ('Имя файла', 'Имя исходного файла без расширения'),
    'index': ('Номер', 'Порядковый номер при экспорте'),
    'date': ('Дата', 'Дата экспорта (YYYY-MM-DD)'),
    'time': ('Время', 'Время экспорта (HH-MM-SS)'),
}


@dataclass
class ExportSettings:
    """
    Complete export settings configuration.
    
    Combines all settings for the DXF export process.
    """
    
    # Output directory
    output_directory: str = ""
    create_subdirectories: bool = False
    subdirectory_template: str = "{date}"
    
    # File settings
    filename_settings: FilenameSettings = field(default_factory=FilenameSettings)
    overwrite_existing: bool = False
    
    # What to export
    export_hidden_parts: bool = False
    export_standard_parts: bool = False
    export_subassemblies: bool = True
    
    # Line type settings
    line_types: Dict[str, LineTypeSettings] = field(
        default_factory=lambda: {k: LineTypeSettings(**v.to_dict()) for k, v in DEFAULT_LINE_TYPES.items()}
    )
    
    # DXF format settings
    dxf_version: str = "AC1027"  # AutoCAD 2013 format
    dxf_units: str = "mm"
    
    # Processing options
    straighten_before_export: bool = True  # Unfold sheet metal
    remove_bend_lines: bool = False        # Remove fold lines from export
    
    # Convenience properties for UI compatibility
    @property
    def output_dir(self) -> str:
        """Alias for output_directory."""
        return self.output_directory
    
    @output_dir.setter
    def output_dir(self, value: str):
        """Setter for output_dir alias."""
        self.output_directory = value
    
    @property
    def create_assembly_subfolder(self) -> bool:
        """Alias for create_subdirectories."""
        return self.create_subdirectories
    
    @create_assembly_subfolder.setter
    def create_assembly_subfolder(self, value: bool):
        """Setter for create_assembly_subfolder alias."""
        self.create_subdirectories = value
    
    @property
    def filename_pattern(self) -> str:
        """Alias for filename_settings.template."""
        return self.filename_settings.template
    
    @filename_pattern.setter
    def filename_pattern(self, value: str):
        """Setter for filename_pattern alias."""
        self.filename_settings.template = value
    
    @property
    def cut_contour(self) -> LineTypeSettings:
        """Get cut contour line settings."""
        return self.line_types.get('contour', LineTypeSettings(key='contour'))
    
    @property
    def bend_lines(self) -> LineTypeSettings:
        """Get bend lines settings (combines up and down)."""
        return self.line_types.get('bend_up', LineTypeSettings(key='bend_up'))
    
    def get_output_path(self, base_name: str, variables: Optional[Dict[str, str]] = None) -> Path:
        """
        Get full output path for a DXF file.
        
        Args:
            base_name: Base filename (without extension)
            variables: Variables for filename template
            
        Returns:
            Path object for output file
        """
        # Start with output directory
        output_dir = Path(self.output_directory) if self.output_directory else Path.cwd()
        
        # Add subdirectory if configured
        if self.create_subdirectories and self.subdirectory_template:
            from datetime import datetime
            now = datetime.now()
            subdir_vars = {
                'date': now.strftime('%Y-%m-%d'),
                'time': now.strftime('%H-%M-%S'),
            }
            subdir = self.subdirectory_template.format(**subdir_vars)
            output_dir = output_dir / subdir
        
        # Generate filename
        if variables:
            filename = self.filename_settings.format(variables)
        else:
            filename = base_name + '.dxf'
        
        return output_dir / filename
    
    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            'output_directory': self.output_directory,
            'create_subdirectories': self.create_subdirectories,
            'subdirectory_template': self.subdirectory_template,
            'filename_settings': self.filename_settings.to_dict(),
            'overwrite_existing': self.overwrite_existing,
            'export_hidden_parts': self.export_hidden_parts,
            'export_standard_parts': self.export_standard_parts,
            'export_subassemblies': self.export_subassemblies,
            'line_types': {k: v.to_dict() for k, v in self.line_types.items()},
            'dxf_version': self.dxf_version,
            'dxf_units': self.dxf_units,
            'straighten_before_export': self.straighten_before_export,
            'remove_bend_lines': self.remove_bend_lines,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ExportSettings':
        """Create from dictionary."""
        settings = cls()
        
        settings.output_directory = data.get('output_directory', '')
        settings.create_subdirectories = data.get('create_subdirectories', False)
        settings.subdirectory_template = data.get('subdirectory_template', '{date}')
        settings.overwrite_existing = data.get('overwrite_existing', False)
        settings.export_hidden_parts = data.get('export_hidden_parts', False)
        settings.export_standard_parts = data.get('export_standard_parts', False)
        settings.export_subassemblies = data.get('export_subassemblies', True)
        settings.dxf_version = data.get('dxf_version', 'AC1027')
        settings.dxf_units = data.get('dxf_units', 'mm')
        settings.straighten_before_export = data.get('straighten_before_export', True)
        settings.remove_bend_lines = data.get('remove_bend_lines', False)
        
        if 'filename_settings' in data:
            settings.filename_settings = FilenameSettings.from_dict(data['filename_settings'])
        
        if 'line_types' in data:
            settings.line_types = {
                k: LineTypeSettings.from_dict(v) 
                for k, v in data['line_types'].items()
            }
        
        return settings
    
    def save(self, path: Path) -> bool:
        """Save settings to JSON file."""
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"Error saving settings: {e}")
            return False
    
    @classmethod
    def load(cls, path: Path) -> 'ExportSettings':
        """Load settings from JSON file."""
        if path.exists():
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return cls.from_dict(data)
            except Exception as e:
                print(f"Error loading settings: {e}")
        return cls()
