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

ms = MSProjectWrapper(path="path/to/your/project.mpp")
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
# Returns list of dicts with: risorsa, passo, distanza_interasse, diametro, stabilimento

# Get resource ID
resource_id = ms.retrieve_resource_id("Resource Name")

# Add resource
# - diameter: for mandrino/tastatore
# - pitch and center_to_center: for maschera
# - all None: for others (Testa/Generatore)
ms.append_resource(
    name="New Resource",
    warehouse="Warehouse B",
    diameter=12.5,  # Optional
    pitch=2.0,  # Optional
    center_to_center=3.5  # Optional
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
    category="Category",  # Optional
    warehouse="Warehouse A",  # Optional
    min_diameter=10.0,  # Optional
    max_diameter=20.0,  # Optional
    pitch=2.0,  # Optional
    center_to_center=3.5,  # Optional
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
- `Text2`: Categoria (Category)
- `Number1`: Passo (Pitch)
- `Number2`: Distanza Interasse (Center-to-center distance)
- `Number3`: Diametro (Diameter)