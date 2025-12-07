"""
KOMPAS-3D API Module

Provides COM automation interface to KOMPAS-3D CAD system.
Uses the KOMPAS API version 7 (IApplication, IDocuments, etc.)

Reference: https://help.ascon.ru/KOMPAS_SDK/22/ru-RU/ap1782828.html
"""

from typing import Optional, Any, List, Tuple
from dataclasses import dataclass
from enum import IntEnum
import logging

# Windows COM imports
try:
    import win32com.client
    from win32com.client import Dispatch, GetActiveObject
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

logger = logging.getLogger(__name__)


# KOMPAS SDK Constants
class DocumentType(IntEnum):
    """KOMPAS document types (DocumentTypeEnum)."""
    UNKNOWN = 0
    DRAWING = 1       # ksDocumentDrawing - Чертеж
    FRAGMENT = 2      # ksDocumentFragment - Фрагмент
    SPECIFICATION = 3 # ksDocumentSpecification
    PART = 4          # ksDocumentPart - Деталь
    ASSEMBLY = 5      # ksDocumentAssembly - Сборка
    TEXT = 6          # ksDocumentTextual
    TECH_ASSEMBLY = 7 # ksDocumentTechnologyAssembly


class KompasCommand(IntEnum):
    """KOMPAS command IDs (ksKompasCommandEnum)."""
    REBUILD_3D = 40356           # ksCM3DRebuild
    CREATE_SHEET_FROM_MODEL = 40373  # ksCMCreateSheetFromModel


# Type stubs for KOMPAS interfaces
@dataclass
class InterfaceWrapper:
    """Base wrapper for COM interfaces."""
    _com_object: Any
    
    def __bool__(self) -> bool:
        return self._com_object is not None
    
    @property
    def raw(self) -> Any:
        """Get raw COM object."""
        return self._com_object


class KompasConnection:
    """
    Manages connection to KOMPAS-3D application via COM.
    
    Usage:
        with KompasConnection() as kompas:
            doc = kompas.active_document
            # work with document
    """
    
    PROGID = "KOMPAS.Application.7"
    
    def __init__(self, visible: bool = True, new_instance: bool = False):
        """
        Initialize KOMPAS connection.
        
        Args:
            visible: Whether to show KOMPAS window
            new_instance: If True, create new instance; if False, try to connect to running
        """
        if not HAS_WIN32:
            raise ImportError("pywin32 is required for KOMPAS-3D COM automation")
        
        self._app = None
        self._visible = visible
        self._new_instance = new_instance
        self._owns_instance = False
    
    def connect(self) -> 'KompasAPI':
        """
        Connect to KOMPAS-3D.
        
        Returns:
            KompasAPI instance
        """
        if self._app is not None:
            return KompasAPI(self._app)
        
        try:
            if not self._new_instance:
                # Try to connect to running instance
                try:
                    self._app = GetActiveObject(self.PROGID)
                    logger.info("Connected to running KOMPAS-3D instance")
                except:
                    pass
            
            if self._app is None:
                # Create new instance
                self._app = Dispatch(self.PROGID)
                self._owns_instance = True
                logger.info("Created new KOMPAS-3D instance")
            
            # Set visibility
            self._app.Visible = self._visible
            
            # Suppress message dialogs for automation
            self._app.HideMessage = 1  # Hide messages
            
            return KompasAPI(self._app)
            
        except Exception as e:
            logger.error(f"Failed to connect to KOMPAS-3D: {e}")
            raise ConnectionError(f"Cannot connect to KOMPAS-3D: {e}")
    
    def disconnect(self):
        """Disconnect from KOMPAS-3D."""
        if self._app is not None:
            try:
                # Only quit if we created the instance
                if self._owns_instance:
                    self._app.Quit()
            except:
                pass
            self._app = None
    
    def __enter__(self) -> 'KompasAPI':
        return self.connect()
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
        return False


class KompasAPI:
    """
    High-level API wrapper for KOMPAS-3D.
    
    Provides access to:
    - Application properties and methods
    - Documents collection
    - Active document
    - Converter for DXF export
    """
    
    def __init__(self, app_object: Any):
        """
        Initialize API wrapper.
        
        Args:
            app_object: COM IApplication object
        """
        self._app = app_object
    
    @property
    def application(self) -> Any:
        """Get raw IApplication COM object."""
        return self._app
    
    @property
    def visible(self) -> bool:
        """Get/set application visibility."""
        return bool(self._app.Visible)
    
    @visible.setter
    def visible(self, value: bool):
        self._app.Visible = value
    
    @property
    def active_document(self) -> Optional['KompasDocument']:
        """Get the currently active document."""
        doc = self._app.ActiveDocument
        if doc is not None:
            return KompasDocument(doc)
        return None
    
    @property
    def documents(self) -> 'DocumentsCollection':
        """Get the documents collection."""
        return DocumentsCollection(self._app.Documents)
    
    def get_converter(self, library_path: str = "") -> Optional[Any]:
        """
        Get file converter interface.
        
        Args:
            library_path: Path to converter library (optional)
            
        Returns:
            IConverter interface
        """
        try:
            return self._app.Converter(library_path)
        except Exception as e:
            logger.error(f"Failed to get converter: {e}")
            return None
    
    def execute_command(self, command_id: int, post: bool = False) -> bool:
        """
        Execute KOMPAS system command.
        
        Args:
            command_id: Command ID from ksKompasCommandEnum
            post: Use PostMessage (True) or SendMessage (False)
            
        Returns:
            True if command executed successfully
        """
        try:
            return bool(self._app.ExecuteKompasCommand(command_id, post))
        except Exception as e:
            logger.error(f"Failed to execute command {command_id}: {e}")
            return False
    
    def is_command_available(self, command_id: int) -> bool:
        """Check if a command is currently available."""
        try:
            return bool(self._app.IsKompasCommandEnable(command_id))
        except:
            return False
    
    def get_system_version(self) -> Tuple[int, int, int, int]:
        """
        Get KOMPAS-3D system version.
        
        Returns:
            Tuple of (major, minor, build, revision)
        """
        try:
            major = pythoncom.Variant(0)
            minor = pythoncom.Variant(0)  
            build = pythoncom.Variant(0)
            revision = pythoncom.Variant(0)
            self._app.GetSystemVersion(major, minor, build, revision)
            return (major.value, minor.value, build.value, revision.value)
        except:
            return (0, 0, 0, 0)


class DocumentsCollection:
    """Wrapper for IDocuments collection."""
    
    def __init__(self, docs_object: Any):
        self._docs = docs_object
    
    @property
    def count(self) -> int:
        """Get number of open documents."""
        return self._docs.Count
    
    def __len__(self) -> int:
        return self.count
    
    def __iter__(self):
        for i in range(self.count):
            yield KompasDocument(self._docs.Item(i))
    
    def open(self, path: str, visible: bool = True, read_only: bool = False) -> Optional['KompasDocument']:
        """
        Open a document.
        
        Args:
            path: Full path to the document file
            visible: Open in visible mode
            read_only: Open read-only
            
        Returns:
            KompasDocument or None if failed
        """
        try:
            doc = self._docs.Open(path, visible, read_only)
            if doc is not None:
                return KompasDocument(doc)
        except Exception as e:
            logger.error(f"Failed to open document {path}: {e}")
        return None
    
    def add(self, doc_type: DocumentType, visible: bool = True) -> Optional['KompasDocument']:
        """
        Create a new document.
        
        Args:
            doc_type: Type of document to create
            visible: Create in visible mode
            
        Returns:
            KompasDocument or None if failed
        """
        try:
            doc = self._docs.Add(int(doc_type), visible)
            if doc is not None:
                return KompasDocument(doc)
        except Exception as e:
            logger.error(f"Failed to create document: {e}")
        return None


class KompasDocument:
    """
    Wrapper for IKompasDocument and derived interfaces.
    
    Provides access to document properties and methods.
    Can represent 2D or 3D documents.
    """
    
    def __init__(self, doc_object: Any):
        self._doc = doc_object
        self._doc_type = None
    
    @property
    def raw(self) -> Any:
        """Get raw COM object."""
        return self._doc
    
    @property
    def path_name(self) -> str:
        """Get full path of the document file."""
        try:
            return str(self._doc.PathName) or ""
        except:
            return ""
    
    @property
    def name(self) -> str:
        """Get document name (without path)."""
        try:
            return str(self._doc.Name) or ""
        except:
            return ""
    
    @property 
    def document_type(self) -> DocumentType:
        """Get document type."""
        if self._doc_type is None:
            try:
                self._doc_type = DocumentType(self._doc.DocumentType)
            except:
                self._doc_type = DocumentType.UNKNOWN
        return self._doc_type
    
    @property
    def is_3d(self) -> bool:
        """Check if document is a 3D model (part or assembly)."""
        return self.document_type in (DocumentType.PART, DocumentType.ASSEMBLY)
    
    @property
    def is_2d(self) -> bool:
        """Check if document is a 2D drawing or fragment."""
        return self.document_type in (DocumentType.DRAWING, DocumentType.FRAGMENT)
    
    @property
    def is_assembly(self) -> bool:
        """Check if document is an assembly."""
        return self.document_type == DocumentType.ASSEMBLY
    
    @property
    def is_part(self) -> bool:
        """Check if document is a part."""
        return self.document_type == DocumentType.PART
    
    def close(self, save: bool = False) -> bool:
        """
        Close the document.
        
        Args:
            save: Whether to save before closing
            
        Returns:
            True if closed successfully
        """
        try:
            if save:
                self._doc.Save()
            self._doc.Close(0)  # 0 = don't prompt to save
            return True
        except Exception as e:
            logger.error(f"Failed to close document: {e}")
            return False
    
    def save(self) -> bool:
        """Save the document."""
        try:
            self._doc.Save()
            return True
        except Exception as e:
            logger.error(f"Failed to save document: {e}")
            return False
    
    def save_as(self, path: str) -> bool:
        """
        Save the document with a new name/path.
        
        Args:
            path: New file path
            
        Returns:
            True if saved successfully
        """
        try:
            self._doc.SaveAs(path)
            return True
        except Exception as e:
            logger.error(f"Failed to save document as {path}: {e}")
            return False
    
    def get_3d_document(self) -> Optional['KompasDocument3D']:
        """
        Get 3D document interface.
        
        Returns:
            KompasDocument3D if this is a 3D document, None otherwise
        """
        if not self.is_3d:
            return None
        return KompasDocument3D(self._doc)
    
    def get_2d_document(self) -> Optional['KompasDocument2D']:
        """
        Get 2D document interface.
        
        Returns:
            KompasDocument2D if this is a 2D document, None otherwise
        """
        if not self.is_2d:
            return None
        return KompasDocument2D(self._doc)


class KompasDocument3D:
    """
    Wrapper for IKompasDocument3D interface.
    
    Provides access to 3D model structure, parts, and sheet metal bodies.
    
    Note: Uses dynamic dispatch to access TopPart and other IKompasDocument3D
    properties, since the base IKompasDocument interface doesn't expose them.
    """
    
    def __init__(self, doc_object: Any):
        # Use dynamic dispatch to access IKompasDocument3D properties
        # This is necessary because ActiveDocument returns IKompasDocument,
        # but TopPart is only available on IKompasDocument3D interface.
        # Dynamic dispatch via win32com.client allows COM to resolve
        # the correct interface at runtime.
        if HAS_WIN32:
            self._doc = win32com.client.Dispatch(doc_object)
        else:
            self._doc = doc_object
    
    @property
    def raw(self) -> Any:
        return self._doc
    
    @property
    def top_part(self) -> Optional['Part3D']:
        """
        Get the top-level part/component.
        
        For parts: the part itself
        For assemblies: the root assembly component
        """
        try:
            part = self._doc.TopPart
            if part is not None:
                return Part3D(part)
        except Exception as e:
            logger.error(f"Failed to get TopPart: {e}")
        return None


class Part3D:
    """
    Wrapper for IPart7 interface.
    
    Represents a 3D component (part or assembly component).
    Provides access to:
    - Part properties (name, marking, material, etc.)
    - Child parts collection
    - Sheet metal bodies
    
    Note: Uses dynamic dispatch to access all IPart7 properties reliably.
    """
    
    def __init__(self, part_object: Any):
        # Use dynamic dispatch to access IPart7 properties reliably
        # This is necessary because the COM object might be early-bound
        # to a different interface that doesn't expose all properties.
        if HAS_WIN32:
            self._part = win32com.client.Dispatch(part_object)
        else:
            self._part = part_object
        self._sheet_metal_container = None
    
    def _debug_com_info(self) -> None:
        """Debug method to log COM object information."""
        try:
            # Log basic type info
            logger.debug(f"COM object type: {type(self._part)}")
            
            # Try to get TypeInfo
            try:
                import pythoncom
                disp = self._part._oleobj_
                type_info = disp.GetTypeInfo(0, pythoncom.LOCALE_USER_DEFAULT)
                logger.debug(f"TypeInfo: {type_info}")
            except:
                pass
            
            # List some expected properties/methods
            expected = ['Name', 'Marking', 'FileName', 'Detail', 'Parts', 
                       'SheetMetalBodies', 'SubFeatures', 'Bodies']
            for prop in expected:
                try:
                    val = getattr(self._part, prop, '<NOT FOUND>')
                    if val != '<NOT FOUND>':
                        logger.debug(f"  {prop}: exists (type={type(val).__name__})")
                    else:
                        logger.debug(f"  {prop}: NOT FOUND")
                except Exception as e:
                    logger.debug(f"  {prop}: ERROR ({type(e).__name__}: {e})")
        except Exception as e:
            logger.debug(f"_debug_com_info error: {e}")
    
    @property
    def raw(self) -> Any:
        return self._part
    
    @property
    def name(self) -> str:
        """Get part name."""
        try:
            # Try to get name from IModelObject interface
            return str(self._part.Name) or ""
        except:
            return ""
    
    @property
    def marking(self) -> str:
        """Get part designation/marking (Обозначение)."""
        try:
            return str(self._part.Marking) or ""
        except:
            return ""
    
    @property
    def file_name(self) -> str:
        """Get part source file name."""
        try:
            return str(self._part.FileName) or ""
        except:
            return ""
    
    @property
    def material(self) -> str:
        """Get part material."""
        try:
            mat = self._part.Material
            if mat:
                # Material might be an object, try to get name
                if hasattr(mat, 'Name'):
                    return str(mat.Name)
                return str(mat)
        except:
            pass
        return ""
    
    @property
    def mass(self) -> float:
        """Get part mass."""
        try:
            return float(self._part.Mass)
        except:
            return 0.0
    
    @property
    def density(self) -> float:
        """Get part density."""
        try:
            return float(self._part.Density)
        except:
            return 0.0
    
    @property
    def is_detail(self) -> bool:
        """Check if this is a detail (not a subassembly)."""
        try:
            result = bool(self._part.Detail)
            logger.debug(f"Part.Detail property: {result}")
            return result
        except Exception as e:
            logger.debug(f"Failed to get Detail property: {e}, assuming True")
            return True
    
    @property
    def is_standard(self) -> bool:
        """Check if this is a standard part."""
        try:
            return bool(self._part.Standard)
        except:
            return False
    
    @property
    def parts(self) -> List['Part3D']:
        """Get collection of child parts/components."""
        result = []
        try:
            parts_collection = self._part.Parts
            if parts_collection:
                count = parts_collection.Count
                logger.debug(f"Parts collection has {count} items")
                for i in range(count):
                    try:
                        part = parts_collection.Part(i)
                        if part:
                            child = Part3D(part)
                            logger.debug(f"  Child part[{i}]: name='{child.name}', marking='{child.marking}'")
                            result.append(child)
                    except Exception as e:
                        logger.debug(f"  Error getting part[{i}]: {e}")
                        continue
            else:
                logger.debug("Parts collection is None (leaf part)")
        except Exception as e:
            logger.debug(f"Failed to get parts collection: {e}")
        return result
    
    @property
    def instance_count(self) -> int:
        """Get number of instances of this component."""
        try:
            return int(self._part.InstanceCount)
        except:
            return 1
    
    def get_sheet_metal_container(self) -> Optional['SheetMetalContainer']:
        """
        Get sheet metal container for accessing sheet metal bodies.
        
        Uses multiple detection methods:
        1. SheetMetalBodies property (ISheetMetalContainer interface)
        2. SubFeatures(74) for o3d_sheetMetalBody type
        
        Returns:
            SheetMetalContainer or None if not a sheet metal part
        """
        if self._sheet_metal_container is not None:
            return self._sheet_metal_container
        
        part_id = self.name or self.marking or self.file_name
        logger.debug(f"Checking sheet metal for part: {part_id}")
        
        # Constant for sheet metal body type from ksObj3dTypeEnum
        O3D_SHEET_METAL_BODY = 74
        
        try:
            # Method 1: Direct SheetMetalBodies property access
            # IPart7 inherits from ISheetMetalContainer which has this property
            try:
                bodies = self._part.SheetMetalBodies
                if bodies is not None:
                    try:
                        count = bodies.Count
                        logger.debug(f"  Method 1 - SheetMetalBodies.Count = {count}")
                        if count > 0:
                            logger.info(f"Found {count} sheet metal bodies via SheetMetalBodies property")
                            self._sheet_metal_container = SheetMetalContainer(self._part)
                            return self._sheet_metal_container
                    except Exception as e:
                        logger.debug(f"  Method 1 - Error getting Count: {type(e).__name__}: {e}")
                else:
                    logger.debug(f"  Method 1 - SheetMetalBodies is None")
            except AttributeError:
                logger.debug(f"  Method 1 - SheetMetalBodies property not found")
            except Exception as e:
                logger.debug(f"  Method 1 - Error: {type(e).__name__}: {e}")
            
            # Method 2: Try SubFeatures to find sheet metal body features
            # SubFeatures(treeType, through, libObject) from IFeature7
            try:
                # Get sheet metal body features (type 74)
                sub_features = self._part.SubFeatures(O3D_SHEET_METAL_BODY, True, False)
                if sub_features is not None:
                    # SubFeatures returns SAFEARRAY or single dispatch
                    if hasattr(sub_features, '__len__'):
                        count = len(sub_features)
                    elif hasattr(sub_features, 'Count'):
                        count = sub_features.Count
                    else:
                        # Single object returned
                        count = 1
                    
                    logger.debug(f"  Method 2 - SubFeatures(74) count = {count}")
                    if count > 0:
                        logger.info(f"Found {count} sheet metal bodies via SubFeatures(74)")
                        self._sheet_metal_container = SheetMetalContainer(self._part)
                        return self._sheet_metal_container
                else:
                    logger.debug(f"  Method 2 - SubFeatures(74) is None")
            except AttributeError:
                logger.debug(f"  Method 2 - SubFeatures not available")
            except Exception as e:
                logger.debug(f"  Method 2 - Error: {type(e).__name__}: {e}")
            
            # Method 3: Try GetSubFeatures method (alternative automation syntax)
            try:
                sub_features = self._part.GetSubFeatures(O3D_SHEET_METAL_BODY, True, False)
                if sub_features is not None:
                    if hasattr(sub_features, '__len__'):
                        count = len(sub_features)
                    elif hasattr(sub_features, 'Count'):
                        count = sub_features.Count
                    else:
                        count = 1
                    
                    logger.debug(f"  Method 3 - GetSubFeatures(74) count = {count}")
                    if count > 0:
                        logger.info(f"Found {count} sheet metal bodies via GetSubFeatures(74)")
                        self._sheet_metal_container = SheetMetalContainer(self._part)
                        return self._sheet_metal_container
            except AttributeError:
                pass
            except Exception as e:
                logger.debug(f"  Method 3 - Error: {type(e).__name__}: {e}")
            
            logger.debug(f"  No sheet metal bodies found in part: {part_id}")
                
        except Exception as e:
            logger.debug(f"Error in get_sheet_metal_container: {type(e).__name__}: {e}")
        
        return None
    
    def _has_sheet_metal(self) -> bool:
        """Check if part has sheet metal bodies."""
        try:
            # Try direct property access
            bodies = self._part.SheetMetalBodies
            return bodies is not None and bodies.Count > 0
        except:
            return False
    
    def get_property_value(self, property_name: str) -> Optional[str]:
        """
        Get a property value by name.
        
        Args:
            property_name: Property name
            
        Returns:
            Property value as string, or None
        """
        try:
            # Try to get property through IPropertyKeeper
            keeper = self._part
            value = keeper.GetPropertyValue(property_name, True, False)
            return str(value) if value else None
        except:
            return None


class SheetMetalContainer:
    """
    Wrapper for ISheetMetalContainer interface.
    
    Provides access to sheet metal operations and bodies.
    
    Note: Uses dynamic dispatch for reliable COM property access.
    """
    
    def __init__(self, part_object: Any):
        # Use dynamic dispatch for reliable COM access
        if HAS_WIN32:
            self._part = win32com.client.Dispatch(part_object)
        else:
            self._part = part_object
    
    @property
    def sheet_metal_bodies(self) -> List['SheetMetalBody']:
        """Get collection of sheet metal bodies."""
        result = []
        try:
            bodies = self._part.SheetMetalBodies
            if bodies:
                count = bodies.Count
                logger.debug(f"SheetMetalBodies count: {count}")
                for i in range(count):
                    try:
                        # Try different indexing methods - KOMPAS may use Item() or SheetMetalBody()
                        body = None
                        try:
                            body = bodies.SheetMetalBody(i)
                        except:
                            try:
                                body = bodies.Item(i)
                            except:
                                try:
                                    # Some collections use 1-based indexing
                                    body = bodies.Item(i + 1)
                                except:
                                    pass
                        if body:
                            result.append(SheetMetalBody(body))
                    except Exception as e:
                        logger.debug(f"Error getting body at index {i}: {e}")
                        continue
        except Exception as e:
            logger.debug(f"Failed to get sheet metal bodies: {e}")
        return result
    
    @property
    def has_sheet_metal(self) -> bool:
        """Check if container has sheet metal bodies."""
        try:
            bodies = self._part.SheetMetalBodies
            return bodies is not None and bodies.Count > 0
        except:
            return False


class SheetMetalBody:
    """
    Wrapper for ISheetMetalBody interface.
    
    Represents a sheet metal body with properties like thickness.
    
    Note: Uses dynamic dispatch for reliable COM property access.
    """
    
    def __init__(self, body_object: Any):
        # Use dynamic dispatch for reliable COM access
        if HAS_WIN32:
            self._body = win32com.client.Dispatch(body_object)
        else:
            self._body = body_object
    
    @property
    def raw(self) -> Any:
        return self._body
    
    @property
    def thickness(self) -> float:
        """Get sheet metal thickness."""
        try:
            return float(self._body.Thickness)
        except:
            return 0.0
    
    @property
    def bend_radius(self) -> float:
        """Get default bend radius."""
        try:
            return float(self._body.Radius)
        except:
            return 0.0
    
    @property
    def bend_coefficient(self) -> float:
        """Get bend coefficient (neutral layer coefficient)."""
        try:
            return float(self._body.BendCoefficient)
        except:
            return 0.4
    
    @property
    def is_straightened(self) -> bool:
        """Check if body is in straightened (unfolded) state."""
        try:
            return bool(self._body.Straighten)
        except:
            return False
    
    @is_straightened.setter
    def is_straightened(self, value: bool):
        """Set straightened (unfolded) state."""
        try:
            self._body.Straighten = value
        except Exception as e:
            logger.error(f"Failed to set straighten state: {e}")


class KompasDocument2D:
    """
    Wrapper for IKompasDocument2D interface.
    
    Provides access to 2D document views and layers.
    """
    
    def __init__(self, doc_object: Any):
        self._doc = doc_object
    
    @property
    def raw(self) -> Any:
        return self._doc
    
    @property
    def views_and_layers_manager(self) -> Optional['ViewsAndLayersManager']:
        """Get views and layers manager."""
        try:
            manager = self._doc.ViewsAndLayersManager
            if manager:
                return ViewsAndLayersManager(manager)
        except Exception as e:
            logger.error(f"Failed to get ViewsAndLayersManager: {e}")
        return None


class ViewsAndLayersManager:
    """
    Wrapper for IViewsAndLayersManager interface.
    """
    
    def __init__(self, manager_object: Any):
        self._manager = manager_object
    
    @property
    def views(self) -> List['View2D']:
        """Get collection of views."""
        result = []
        try:
            views_collection = self._manager.Views
            if views_collection:
                count = views_collection.Count
                for i in range(count):
                    try:
                        view = views_collection.View(i)
                        if view:
                            result.append(View2D(view))
                    except:
                        continue
        except Exception as e:
            logger.debug(f"Failed to get views: {e}")
        return result


class View2D:
    """
    Wrapper for IView interface.
    """
    
    def __init__(self, view_object: Any):
        self._view = view_object
    
    @property
    def raw(self) -> Any:
        return self._view
    
    @property
    def name(self) -> str:
        """Get view name."""
        try:
            return str(self._view.Name) or ""
        except:
            return ""
    
    @property
    def layers(self) -> List['Layer2D']:
        """Get collection of layers in this view."""
        result = []
        try:
            layers_collection = self._view.Layers
            if layers_collection:
                count = layers_collection.Count
                for i in range(count):
                    try:
                        layer = layers_collection.Layer(i)
                        if layer:
                            result.append(Layer2D(layer))
                    except:
                        continue
        except Exception as e:
            logger.debug(f"Failed to get layers: {e}")
        return result


class Layer2D:
    """
    Wrapper for ILayer interface.
    """
    
    def __init__(self, layer_object: Any):
        self._layer = layer_object
    
    @property
    def raw(self) -> Any:
        return self._layer
    
    @property
    def name(self) -> str:
        """Get layer name."""
        try:
            return str(self._layer.Name) or ""
        except:
            return ""
    
    @name.setter
    def name(self, value: str):
        try:
            self._layer.Name = value
        except:
            pass
    
    @property
    def color(self) -> int:
        """Get layer color."""
        try:
            return int(self._layer.Color)
        except:
            return 0
    
    @color.setter
    def color(self, value: int):
        try:
            self._layer.Color = value
        except:
            pass
    
    @property
    def visible(self) -> bool:
        """Get layer visibility."""
        try:
            return bool(self._layer.Visible)
        except:
            return True
    
    @visible.setter
    def visible(self, value: bool):
        try:
            self._layer.Visible = value
        except:
            pass
    
    @property
    def printable(self) -> bool:
        """Get layer printable state."""
        try:
            return bool(self._layer.Printable)
        except:
            return True
    
    @printable.setter
    def printable(self, value: bool):
        try:
            self._layer.Printable = value
        except:
            pass
