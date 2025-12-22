# MS Project Wrapper

Python wrapper for interacting with Microsoft Project files using win32com.

## Requirements

- Windows OS
- Microsoft Project installed
- Python 3.13+

## Installation

```bash
poetry install
```

## Usage

```python
from src import MSProjectWrapper
from datetime import datetime

# Open project file
with MSProjectWrapper(path="path/to/your/project.mpp") as ms:
    # Work with the project
    # File is automatically saved and closed when exiting the 'with' block
    pass

# Or use manually:
ms = MSProjectWrapper(path="path/to/your/project.mpp")
# ... work with project
ms.close()  # Saves and closes the project
```

### Tasks

```python
# Get all tasks
tasks = ms.tasks()
# Returns list of dicts with: task, gerarchia, inizio, fine, stabilimento, responsabile, risorse

# Get specific task
task = ms.retrieve_task("Task Name")
task_id = ms.retrieve_task_id("Task Name")

# Add task
ms.append_task(
    name="New Task",
    start=datetime(2025, 1, 1),
    end=datetime(2025, 1, 31),
    warehouse="Warehouse A",
    manager="John Doe",
    parent=ms.retrieve_task("Parent Task Name")  # None for top-level task
)

# Delete task
ms.delete_task(task_id)
```

### Resources

```python
# Get all resources
resources = ms.resources()
# Returns list of dicts with: categoria, risorsa, modello, passo, distanza_interasse, diametro, max, stabilimento, note

# Get resource ID
resource_id = ms.retrieve_resource_id("Resource Name")

# Add resource
ms.append_resource(
    name="New Resource",
    category="Tastatore",  # Valid categories: Tastatore, Mandrino, Maschera, Testa, Generatore
    warehouse="Warehouse B",
    diameter=12.5,  # Optional
    pitch=2.0,  # Optional
    center_to_center=3.5,  # Optional
    model="Model123",  # Optional
    max=100.0,  # Optional
    note="Some notes"  # Optional
)

# Delete resource
ms.delete_resource(resource_id)

# Assign resources to task
ms.assign_resources(task, [resource_id1, resource_id2])

# Check availability
is_available = ms.check_availability(
    resource_id=resource_id,
    start=datetime(2025, 1, 1),
    end=datetime(2025, 1, 31)
)

# Query resources with filters
results = ms.query(
    category="Tastatore",  # Optional
    warehouse="Warehouse A",  # Optional
    min_diameter=10.0,  # Optional
    max_diameter=20.0,  # Optional
    pitch=2.0,  # Optional
    center_to_center=3.5,  # Optional
    model="Model123",  # Optional
    start=datetime(2025, 1, 1),  # Optional (requires end)
    end=datetime(2025, 1, 31)  # Optional (requires start)
)
```

## Custom Fields

### Tasks
- `Text1`: Stabilimento (Warehouse)
- `Text2`: Responsabile (Manager)

### Resources
- `Text1`: Stabilimento (Warehouse)
- `Text2`: Categoria (Category) - Valid values: Tastatore, Mandrino, Maschera, Testa, Generatore
- `Text3`: Modello (Model)
- `Number1`: Passo (Pitch)
- `Number2`: Distanza Interasse (Center-to-center distance)
- `Number3`: Diametro (Diameter)
- `Number4`: Max
- `Notes`: Note aggiuntive