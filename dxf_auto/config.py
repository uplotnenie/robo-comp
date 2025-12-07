"""
DXF-Auto Configuration Module

Application-wide configuration settings and constants.
"""

import os
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Optional
import json

# Application Info
APP_NAME = "DXF-Auto"
APP_VERSION = "1.0.0"
APP_AUTHOR = "DXF-Auto Team"

# KOMPAS-3D COM Settings
KOMPAS_PROGID = "KOMPAS.Application.7"  # ProgID for KOMPAS-3D API v7

# Document Type Enum (from KOMPAS SDK DocumentTypeEnum)
class DocumentType:
    UNKNOWN = 0
    DRAWING = 1      # ksDocumentDrawing - Чертеж
    FRAGMENT = 2     # ksDocumentFragment - Фрагмент
    SPECIFICATION = 3
    PART = 4         # ksDocumentPart - Деталь
    ASSEMBLY = 5     # ksDocumentAssembly - Сборка
    TEXT = 6
    TECH_ASSEMBLY = 7


# File Extensions
KOMPAS_EXTENSIONS = {
    'assembly': '.a3d',
    'part': '.m3d',
    'drawing': '.cdw',
    'fragment': '.frw',
}

DXF_EXTENSION = '.dxf'


# Default Paths
@dataclass
class AppPaths:
    """Application paths configuration."""
    app_dir: Path = field(default_factory=lambda: Path(__file__).parent)
    config_dir: Path = field(default_factory=lambda: Path.home() / ".dxf_auto")
    output_dir: Path = field(default_factory=lambda: Path.home() / "Documents" / "DXF_Export")
    
    def __post_init__(self):
        # Ensure directories exist
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    @property
    def settings_file(self) -> Path:
        return self.config_dir / "settings.json"
    
    @property
    def recent_files(self) -> Path:
        return self.config_dir / "recent.json"


# Line Type Settings for DXF Export
@dataclass
class LineTypeConfig:
    """Configuration for a line type in DXF export."""
    name: str
    enabled: bool = True
    layer_name: str = ""
    color: int = 0  # AutoCAD color index
    line_type: str = "CONTINUOUS"
    
    def __post_init__(self):
        if not self.layer_name:
            self.layer_name = self.name


# Default Line Types for Sheet Metal
DEFAULT_LINE_TYPES = {
    'contour': LineTypeConfig(
        name='Контур',
        enabled=True,
        layer_name='CONTOUR',
        color=7,  # White
        line_type='CONTINUOUS'
    ),
    'bend_up': LineTypeConfig(
        name='Линия сгиба вверх',
        enabled=True,
        layer_name='BEND_UP',
        color=1,  # Red
        line_type='DASHED'
    ),
    'bend_down': LineTypeConfig(
        name='Линия сгиба вниз',
        enabled=True,
        layer_name='BEND_DOWN',
        color=3,  # Green
        line_type='DASHED'
    ),
    'internal': LineTypeConfig(
        name='Внутренние линии',
        enabled=False,
        layer_name='INTERNAL',
        color=4,  # Cyan
        line_type='CONTINUOUS'
    ),
}


# Filename Template Variables
FILENAME_VARIABLES = {
    '{designation}': 'Обозначение детали',
    '{name}': 'Наименование детали',
    '{material}': 'Материал',
    '{thickness}': 'Толщина листа',
    '{mass}': 'Масса',
    '{filename}': 'Имя файла источника',
    '{index}': 'Порядковый номер',
    '{date}': 'Дата экспорта',
    '{time}': 'Время экспорта',
}

DEFAULT_FILENAME_TEMPLATE = "{designation}_{name}"


# UI Configuration
@dataclass
class UIConfig:
    """UI appearance configuration."""
    window_width: int = 1200
    window_height: int = 800
    min_width: int = 800
    min_height: int = 600
    tree_width: int = 300
    table_columns: List[str] = field(default_factory=lambda: [
        'Обозначение',
        'Наименование', 
        'Материал',
        'Толщина',
        'Кол-во',
        'Имя файла DXF'
    ])
    font_family: str = "Segoe UI"
    font_size: int = 10


# Export Settings
@dataclass  
class ExportConfig:
    """DXF export configuration."""
    output_directory: str = ""
    filename_template: str = DEFAULT_FILENAME_TEMPLATE
    create_subdirectories: bool = False
    overwrite_existing: bool = False
    export_hidden: bool = False
    line_types: Dict[str, LineTypeConfig] = field(default_factory=lambda: DEFAULT_LINE_TYPES.copy())
    
    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            'output_directory': self.output_directory,
            'filename_template': self.filename_template,
            'create_subdirectories': self.create_subdirectories,
            'overwrite_existing': self.overwrite_existing,
            'export_hidden': self.export_hidden,
            'line_types': {
                k: {
                    'name': v.name,
                    'enabled': v.enabled,
                    'layer_name': v.layer_name,
                    'color': v.color,
                    'line_type': v.line_type
                } for k, v in self.line_types.items()
            }
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ExportConfig':
        """Create from dictionary."""
        config = cls()
        config.output_directory = data.get('output_directory', '')
        config.filename_template = data.get('filename_template', DEFAULT_FILENAME_TEMPLATE)
        config.create_subdirectories = data.get('create_subdirectories', False)
        config.overwrite_existing = data.get('overwrite_existing', False)
        config.export_hidden = data.get('export_hidden', False)
        
        if 'line_types' in data:
            config.line_types = {
                k: LineTypeConfig(**v) for k, v in data['line_types'].items()
            }
        
        return config


def load_settings(paths: AppPaths) -> ExportConfig:
    """Load settings from file."""
    if paths.settings_file.exists():
        try:
            with open(paths.settings_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return ExportConfig.from_dict(data)
        except Exception as e:
            print(f"Error loading settings: {e}")
    
    return ExportConfig()


def save_settings(config: ExportConfig, paths: AppPaths) -> bool:
    """Save settings to file."""
    try:
        with open(paths.settings_file, 'w', encoding='utf-8') as f:
            json.dump(config.to_dict(), f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error saving settings: {e}")
        return False
