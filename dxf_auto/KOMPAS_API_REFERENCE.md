# KOMPAS-3D API Reference Guide

Полное руководство по API КОМПАС-3D, используемому в приложении DXF-Auto.

## Содержание

1. [Введение](#введение)
2. [Подключение к КОМПАС-3D](#подключение-к-компас-3d)
3. [Интерфейс IApplication](#интерфейс-iapplication)
4. [Работа с документами](#работа-с-документами)
5. [3D документы и детали](#3d-документы-и-детали)
6. [Листовые тела](#листовые-тела)
7. [Развёртки](#развёртки)
8. [Экспорт в DXF](#экспорт-в-dxf)
9. [Свойства модели](#свойства-модели)
10. [Константы и перечисления](#константы-и-перечисления)
11. [Примеры кода](#примеры-кода)

---

## Введение

### О KOMPAS API

КОМПАС-3D предоставляет COM-интерфейс для автоматизации работы с CAD-системой. API версии 7 (API7) — основной интерфейс для современных версий КОМПАС.

### Требования

- Windows OS
- КОМПАС-3D v17 или новее
- Python 3.8+
- pywin32 (`pip install pywin32`)

### ProgID

```python
KOMPAS_PROGID = "KOMPAS.Application.7"
```

---

## Подключение к КОМПАС-3D

### Получение объекта приложения

```python
import win32com.client
from win32com.client import Dispatch, GetActiveObject

# Подключение к запущенному экземпляру
try:
    app = GetActiveObject("KOMPAS.Application.7")
    print("Подключено к существующему экземпляру")
except:
    # Создание нового экземпляра
    app = Dispatch("KOMPAS.Application.7")
    print("Создан новый экземпляр")

# Сделать окно видимым
app.Visible = True

# Скрыть диалоговые окна (для автоматизации)
app.HideMessage = 1  # ksHideMessageYes
```

### Параметры HideMessage

| Значение | Константа | Описание |
|----------|-----------|----------|
| 0 | ksHideMessageNo | Показывать все сообщения |
| 1 | ksHideMessageYes | Скрывать сообщения, нажимать "Да" |
| 2 | ksHideMessageNo2 | Скрывать, нажимать "Нет" |

### Завершение работы

```python
# Закрытие приложения (только если создано нами)
app.Quit()
```

---

## Интерфейс IApplication

Главный интерфейс приложения КОМПАС-3D.

### Свойства

| Свойство | Тип | Описание |
|----------|-----|----------|
| `Visible` | bool | Видимость главного окна |
| `HideMessage` | int | Режим скрытия диалогов |
| `ActiveDocument` | IKompasDocument | Активный документ |
| `Documents` | IDocuments | Коллекция открытых документов |
| `ApplicationName` | str | Название приложения |

### Методы

#### GetSystemVersion

Получение версии КОМПАС-3D.

```python
# Параметры передаются по ссылке через VARIANT
import pythoncom

major = pythoncom.Variant(0)
minor = pythoncom.Variant(0)
build = pythoncom.Variant(0)
revision = pythoncom.Variant(0)

app.ksGetSystemVersion(major, minor, build, revision)

print(f"Версия: {major.value}.{minor.value}.{build.value}.{revision.value}")
```

#### ExecuteKompasCommand

Выполнение команды КОМПАС.

```python
# Параметры: commandId, usePostMessage
result = app.ExecuteKompasCommand(command_id, False)

# Пример: Перестроить модель
REBUILD_3D = 40356  # ksCM3DRebuild
app.ExecuteKompasCommand(REBUILD_3D, False)
```

#### IsKompasCommandEnable

Проверка доступности команды.

```python
if app.IsKompasCommandEnable(command_id):
    app.ExecuteKompasCommand(command_id, False)
```

#### Converter

Получение конвертера файлов.

```python
converter = app.Converter("")  # Пустая строка = конвертер по умолчанию
```

---

## Работа с документами

### Интерфейс IDocuments

Коллекция открытых документов.

```python
docs = app.Documents

# Количество документов
count = docs.Count

# Доступ по индексу (0-based)
doc = docs.Item(0)

# Итерация
for i in range(docs.Count):
    doc = docs.Item(i)
    print(doc.Name)
```

### Открытие документа

```python
# Параметры: path, visible, readOnly
doc = docs.Open(
    r"C:\Models\Part.m3d",  # Путь к файлу
    True,                    # Видимый
    False                    # Только чтение
)
```

### Создание документа

```python
# Параметры: docType, visible
# docType из DocumentTypeEnum
doc = docs.Add(4, True)  # 4 = ksDocumentPart (деталь)
```

### Типы документов (DocumentTypeEnum)

| Значение | Константа | Расширение | Описание |
|----------|-----------|------------|----------|
| 0 | ksDocumentUnknown | - | Неизвестный |
| 1 | ksDocumentDrawing | .cdw | Чертёж |
| 2 | ksDocumentFragment | .frw | Фрагмент |
| 3 | ksDocumentSpecification | .spw | Спецификация |
| 4 | ksDocumentPart | .m3d | Деталь |
| 5 | ksDocumentAssembly | .a3d | Сборка |
| 6 | ksDocumentTextual | .kdw | Текстовый документ |

---

## Интерфейс IKompasDocument

Базовый интерфейс документа.

### Свойства

| Свойство | Тип | Описание |
|----------|-----|----------|
| `Name` | str | Имя документа (без пути) |
| `PathName` | str | Полный путь к файлу |
| `DocumentType` | int | Тип документа |
| `Changed` | bool | Документ изменён |
| `Active` | bool | Документ активен |
| `ReadOnly` | bool | Только для чтения |

### Методы

#### Save / SaveAs

```python
# Сохранение
doc.Save()

# Сохранить как
doc.SaveAs(r"C:\Output\NewFile.m3d")
```

#### Close

```python
# Параметр: askSave (спрашивать о сохранении)
doc.Close(False)  # Закрыть без вопроса
```

#### Activate

```python
doc.Activate()  # Сделать документ активным
```

---

## 3D документы и детали

### Получение 3D документа

```python
# Проверка типа
if doc.DocumentType in [4, 5]:  # Деталь или сборка
    # Получение 3D интерфейса через QueryInterface
    doc_3d = doc._oleobj_.QueryInterface(
        pythoncom.IID_IDispatch,
        pythoncom.IID_IDispatch
    )
```

### Интерфейс IKompasDocument3D

Расширенный интерфейс для 3D документов.

### Свойства

| Свойство | Тип | Описание |
|----------|-----|----------|
| `TopPart` | IPart7 | Верхний компонент |
| `Part` | IPart7 | Основная деталь |

```python
# Получение главной детали
top_part = doc_3d.TopPart
```

---

## Интерфейс IPart7

Интерфейс компонента (детали или подсборки).

### Свойства

| Свойство | Тип | Описание |
|----------|-----|----------|
| `Name` | str | Название компонента |
| `FileName` | str | Имя файла |
| `Marking` | str | Обозначение |
| `Parts` | IFeature7 | Коллекция вложенных компонентов |
| `Hidden` | bool | Компонент скрыт |
| `Standard` | bool | Стандартное изделие |

### Методы

#### Получение контейнера листовых тел

```python
# Метод возвращает ISheetMetalContainer
container = part.SheetMetalContainer
```

#### Получение свойств модели

```python
# Через IPropertyMng
prop_mng = part.PropertyMng

# Получение свойства по ID
# property_id из ksPropertyIdEnum
prop = prop_mng.GetProperty(doc, property_id)
value = prop.Value
```

---

## Листовые тела

### Интерфейс ISheetMetalContainer

Контейнер листовых тел детали.

```python
container = part.SheetMetalContainer

# Проверка наличия листовых тел
if container is not None:
    bodies = container.SheetMetalBodies
    if bodies.Count > 0:
        print("Деталь содержит листовые тела")
```

### Интерфейс ISheetMetalBodies

Коллекция листовых тел.

```python
bodies = container.SheetMetalBodies

# Количество
count = bodies.Count

# Доступ по индексу
body = bodies.Item(0)

# Итерация
for i in range(bodies.Count):
    body = bodies.Item(i)
```

### Интерфейс ISheetMetalBody

Отдельное листовое тело.

#### Свойства

| Свойство | Тип | Описание |
|----------|-----|----------|
| `Name` | str | Название тела |
| `Thickness` | float | Толщина листа (мм) |
| `IsStraightened` | bool | Тело развёрнуто |
| `BendRadius` | float | Радиус сгиба по умолчанию |

```python
body = bodies.Item(0)

# Толщина
thickness = body.Thickness
print(f"Толщина: {thickness} мм")

# Проверка развёрнутости
if body.IsStraightened:
    print("Тело развёрнуто")
```

---

## Развёртки

### Создание развёртки

Развёртка (straighten/unfold) создаётся установкой свойства `IsStraightened`.

```python
# Получение листового тела
body = container.SheetMetalBodies.Item(0)

# Сохранение исходного состояния
original_state = body.IsStraightened

# Создание развёртки
body.IsStraightened = True

# ... работа с развёрткой ...

# Восстановление состояния
body.IsStraightened = original_state
```

### Получение размеров развёртки

```python
# После разворачивания можно получить габариты
# через свойства тела или через API измерений

# Альтернативный метод через IUnfoldParam
unfold_param = body.UnfoldParam
if unfold_param:
    # Параметры развёртки
    pass
```

### Линии сгиба

В развёртке присутствуют линии сгиба, которые можно контролировать при экспорте.

```python
# Получение линий сгиба через IBendLines
bend_lines = body.BendLines

for i in range(bend_lines.Count):
    bend = bend_lines.Item(i)
    # bend_type: 1 = вверх, 2 = вниз
    bend_type = bend.BendType
    angle = bend.Angle
```

---

## Экспорт в DXF

### Метод 1: SaveAs с расширением .dxf

Для 2D документов (фрагментов, чертежей):

```python
# Создание фрагмента
fragment = docs.Add(2, False)  # 2 = ksDocumentFragment, невидимый

# ... добавление геометрии ...

# Сохранение как DXF
fragment.SaveAs(r"C:\Output\file.dxf")
fragment.Close(False)
```

### Метод 2: Использование IConverter

```python
converter = app.Converter("")

# Параметры: inputFile, outputFile, commandCode, showParams
result = converter.Convert(
    r"C:\Models\Part.m3d",    # Входной файл
    r"C:\Output\Part.dxf",    # Выходной файл
    0,                         # Код команды
    False                      # Не показывать диалог
)
```

### Метод 3: Через 2D фрагмент

Наиболее надёжный метод для экспорта развёрток:

```python
def export_unfold_to_dxf(doc_3d, output_path):
    """Экспорт развёртки через 2D фрагмент."""
    
    # 1. Получить листовое тело
    top_part = doc_3d.TopPart
    container = top_part.SheetMetalContainer
    body = container.SheetMetalBodies.Item(0)
    
    # 2. Развернуть
    body.IsStraightened = True
    
    # 3. Создать 2D фрагмент
    docs = app.Documents
    fragment = docs.Add(2, False)  # Невидимый фрагмент
    
    # 4. Выполнить команду "Создать эскиз из модели"
    # ksCMCreateSheetFromModel = 40373
    app.ExecuteKompasCommand(40373, False)
    
    # 5. Сохранить как DXF
    fragment.SaveAs(str(output_path))
    
    # 6. Закрыть фрагмент
    fragment.Close(False)
    
    # 7. Восстановить состояние
    body.IsStraightened = False
    
    return True
```

### Настройки DXF экспорта

Версии DXF формата:

| Код | Версия | Описание |
|-----|--------|----------|
| AC1006 | R10 | AutoCAD Release 10 |
| AC1009 | R11/R12 | AutoCAD R11/R12 |
| AC1012 | R13 | AutoCAD R13 |
| AC1014 | R14 | AutoCAD R14 |
| AC1015 | 2000 | AutoCAD 2000 |
| AC1018 | 2004 | AutoCAD 2004 |
| AC1021 | 2007 | AutoCAD 2007 |
| AC1024 | 2010 | AutoCAD 2010 |
| AC1027 | 2013 | AutoCAD 2013 |
| AC1032 | 2018 | AutoCAD 2018 |

---

## Свойства модели

### Интерфейс IPropertyMng

Менеджер свойств для получения атрибутов детали.

```python
prop_mng = part.PropertyMng

# Получение свойства
prop = prop_mng.GetProperty(doc, property_id)
if prop:
    value = prop.Value
```

### Стандартные свойства (ksPropertyIdEnum)

| ID | Константа | Описание |
|----|-----------|----------|
| 1 | ksPropDesignation | Обозначение |
| 2 | ksPropName | Наименование |
| 3 | ksPropMaterial | Материал |
| 4 | ksPropMass | Масса |
| 5 | ksPropDensity | Плотность |
| 6 | ksPropVolume | Объём |
| 7 | ksPropSurfaceArea | Площадь поверхности |
| 8 | ksPropAuthor | Автор |
| 9 | ksPropComment | Комментарий |

### Получение свойств детали

```python
def get_part_properties(doc, part):
    """Получение основных свойств детали."""
    prop_mng = part.PropertyMng
    
    properties = {}
    
    # Обозначение
    prop = prop_mng.GetProperty(doc, 1)
    if prop:
        properties['designation'] = prop.Value
    
    # Наименование
    prop = prop_mng.GetProperty(doc, 2)
    if prop:
        properties['name'] = prop.Value
    
    # Материал
    prop = prop_mng.GetProperty(doc, 3)
    if prop:
        properties['material'] = prop.Value
    
    # Масса
    prop = prop_mng.GetProperty(doc, 4)
    if prop:
        properties['mass'] = float(prop.Value) if prop.Value else 0.0
    
    return properties
```

---

## Работа со сборками

### Получение компонентов сборки

```python
def get_assembly_components(part):
    """Рекурсивное получение компонентов сборки."""
    components = []
    
    # Получение коллекции вложенных компонентов
    parts_feature = part.Parts
    if parts_feature is None:
        return components
    
    # Итератор компонентов
    iterator = parts_feature.ComponentIterator
    iterator.Reset()
    
    while iterator.IsValid:
        component = iterator.Component
        if component:
            components.append({
                'name': component.Name,
                'filename': component.FileName,
                'marking': component.Marking,
                'hidden': component.Hidden,
                'standard': component.Standard,
            })
            
            # Рекурсия для подсборок
            sub_components = get_assembly_components(component)
            components.extend(sub_components)
        
        iterator.MoveNext()
    
    return components
```

### Определение листовой детали

```python
def is_sheet_metal_part(part):
    """Проверка, является ли деталь листовой."""
    try:
        container = part.SheetMetalContainer
        if container is None:
            return False
        
        bodies = container.SheetMetalBodies
        return bodies.Count > 0
    except:
        return False
```

---

## Константы и перечисления

### Типы документов (ksDocumentTypeEnum)

```python
class DocumentType:
    UNKNOWN = 0          # ksDocumentUnknown
    DRAWING = 1          # ksDocumentDrawing (.cdw)
    FRAGMENT = 2         # ksDocumentFragment (.frw)
    SPECIFICATION = 3    # ksDocumentSpecification (.spw)
    PART = 4             # ksDocumentPart (.m3d)
    ASSEMBLY = 5         # ksDocumentAssembly (.a3d)
    TEXT = 6             # ksDocumentTextual (.kdw)
```

### Команды (ksKompasCommandEnum)

```python
class KompasCommand:
    REBUILD_3D = 40356                # ksCM3DRebuild
    CREATE_SHEET_FROM_MODEL = 40373   # ksCMCreateSheetFromModel
    SAVE_AS = 2                       # ksCMSaveAs
    NEW_FRAGMENT = 103                # ksCMNewFragment
```

### Типы линий сгиба

```python
class BendType:
    UP = 1      # Сгиб наружу
    DOWN = 2    # Сгиб внутрь
```

### Цвета AutoCAD (ACI)

```python
class ACIColors:
    RED = 1
    YELLOW = 2
    GREEN = 3
    CYAN = 4
    BLUE = 5
    MAGENTA = 6
    WHITE = 7
    DARK_GRAY = 8
    LIGHT_GRAY = 9
```

---

## Примеры кода

### Полный пример: Экспорт развёртки

```python
import win32com.client
from win32com.client import Dispatch, GetActiveObject
from pathlib import Path

def export_sheet_metal_to_dxf(output_dir: str) -> list:
    """
    Экспорт всех листовых деталей из активного документа.
    
    Returns:
        Список путей к созданным DXF файлам
    """
    results = []
    
    # Подключение к КОМПАС
    try:
        app = GetActiveObject("KOMPAS.Application.7")
    except:
        print("КОМПАС-3D не запущен")
        return results
    
    # Получение активного документа
    doc = app.ActiveDocument
    if doc is None:
        print("Нет открытого документа")
        return results
    
    # Проверка типа документа
    if doc.DocumentType not in [4, 5]:
        print("Документ не является деталью или сборкой")
        return results
    
    # Получение 3D интерфейса
    doc_3d = doc
    top_part = doc_3d.TopPart
    
    # Поиск листовых деталей
    sheet_parts = find_sheet_metal_parts(top_part)
    
    # Экспорт каждой детали
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    for part_info in sheet_parts:
        try:
            dxf_path = output_path / f"{part_info['name']}.dxf"
            
            if export_part_unfold(app, part_info['part'], dxf_path):
                results.append(str(dxf_path))
                print(f"✓ {part_info['name']}")
            else:
                print(f"✗ {part_info['name']}: ошибка экспорта")
                
        except Exception as e:
            print(f"✗ {part_info['name']}: {e}")
    
    return results


def find_sheet_metal_parts(part, level=0):
    """Рекурсивный поиск листовых деталей."""
    results = []
    
    # Проверка текущей детали
    if is_sheet_metal_part(part):
        results.append({
            'name': part.Name or part.FileName,
            'part': part,
            'level': level,
        })
    
    # Рекурсия по компонентам
    try:
        parts_feature = part.Parts
        if parts_feature:
            iterator = parts_feature.ComponentIterator
            iterator.Reset()
            
            while iterator.IsValid:
                component = iterator.Component
                if component and not component.Hidden:
                    sub_results = find_sheet_metal_parts(component, level + 1)
                    results.extend(sub_results)
                iterator.MoveNext()
    except:
        pass
    
    return results


def is_sheet_metal_part(part):
    """Проверка на листовое тело."""
    try:
        container = part.SheetMetalContainer
        if container:
            bodies = container.SheetMetalBodies
            return bodies.Count > 0
    except:
        pass
    return False


def export_part_unfold(app, part, output_path):
    """Экспорт развёртки одной детали."""
    try:
        container = part.SheetMetalContainer
        body = container.SheetMetalBodies.Item(0)
        
        # Запоминаем состояние
        was_straightened = body.IsStraightened
        
        # Разворачиваем
        if not was_straightened:
            body.IsStraightened = True
        
        # Создаём 2D фрагмент
        docs = app.Documents
        fragment = docs.Add(2, False)
        
        # Команда создания эскиза из модели
        app.ExecuteKompasCommand(40373, False)
        
        # Сохраняем как DXF
        fragment.SaveAs(str(output_path))
        fragment.Close(False)
        
        # Восстанавливаем
        if not was_straightened:
            body.IsStraightened = False
        
        return True
        
    except Exception as e:
        print(f"Ошибка: {e}")
        return False


# Запуск
if __name__ == "__main__":
    results = export_sheet_metal_to_dxf(r"C:\DXF_Output")
    print(f"\nЭкспортировано: {len(results)} файлов")
```

### Пример: Получение информации о детали

```python
def get_part_info(doc, part):
    """Получение полной информации о листовой детали."""
    info = {
        'name': part.Name,
        'filename': part.FileName,
        'is_sheet_metal': False,
        'thickness': None,
        'properties': {},
    }
    
    # Проверка на листовое тело
    try:
        container = part.SheetMetalContainer
        if container:
            bodies = container.SheetMetalBodies
            if bodies.Count > 0:
                info['is_sheet_metal'] = True
                info['thickness'] = bodies.Item(0).Thickness
    except:
        pass
    
    # Свойства
    try:
        prop_mng = part.PropertyMng
        
        for prop_id, prop_name in [(1, 'designation'), (2, 'name'), 
                                    (3, 'material'), (4, 'mass')]:
            prop = prop_mng.GetProperty(doc, prop_id)
            if prop:
                info['properties'][prop_name] = prop.Value
    except:
        pass
    
    return info
```

---

## Обработка ошибок

### Типичные ошибки

```python
try:
    app = GetActiveObject("KOMPAS.Application.7")
except Exception as e:
    if "Операция недоступна" in str(e):
        print("КОМПАС-3D не запущен")
    elif "Интерфейс не поддерживается" in str(e):
        print("Неверная версия API")
    else:
        print(f"Неизвестная ошибка: {e}")
```

### Проверки перед операциями

```python
def safe_export(part):
    """Безопасный экспорт с проверками."""
    
    # Проверка типа
    if not hasattr(part, 'SheetMetalContainer'):
        return False, "Не 3D компонент"
    
    container = part.SheetMetalContainer
    if container is None:
        return False, "Нет контейнера листовых тел"
    
    bodies = container.SheetMetalBodies
    if bodies.Count == 0:
        return False, "Нет листовых тел"
    
    # ... экспорт ...
    return True, "OK"
```

---

## Ресурсы

### Официальная документация

- [КОМПАС-3D SDK](https://help.ascon.ru/KOMPAS_SDK/)
- [API7 Reference](https://help.ascon.ru/KOMPAS_SDK/22/ru-RU/ap1782828.html)

### Полезные ссылки

- `help.ascon.ru` — Официальная документация ASCON
- Примеры в папке установки: `C:\Program Files\ASCON\KOMPAS-3D v22\SDK\`

---

## Changelog

| Версия | Дата | Изменения |
|--------|------|-----------|
| 1.0 | 2024-12 | Первая версия документации |

---

*Документация создана для приложения DXF-Auto*
