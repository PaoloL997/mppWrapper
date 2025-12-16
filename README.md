# MS Project Wrapper

Python wrapper for interacting with Microsoft Project files using win32com.

## Requirements

- Windows OS
- Microsoft Project installed
- Python 3.10+

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
tasks = ms.get_tasks()

# Get specific task
task = ms.get_task("Task Name")
task_id = ms.get_task_id("Task Name")

# Add task
ms.add_task(
    name="New Task",
    start=datetime(2025, 1, 1),
    end=datetime(2025, 1, 31),
    warehouse="Warehouse A",
    manager="John Doe",
    parent=ms.get_task("Parent Task Name")  # None if u want to add a parent task
)

# Delete task
ms.delete_task(task_id)
```

### Resources

```python
# Get all resources
resources = ms.get_resources()

# Get resource ID
resource_id = ms.get_resource_id("Resource Name")

# Add resource
# - diameter: for mandrino/tastatore
# - pitch and center_to_center: for maschera
# - all None: for others (Testa/Generatore)
ms.add_resource(
    name="New Resource",
    warehouse="Warehouse B",
    diameter=12.5,  # Optional
    pitch=2.0,  # Optional
    center_to_center=3.5  # Optional
)

# Delete resource
ms.delete_resource(resource_id)

# Assign resources to task
ms.add_resources_to_task(task, [resource_id1, resource_id2])

# Check availability
is_available = ms.check_resource_availability(
    resource_id=resource_id,
    start=datetime(2025, 1, 1),
    end=datetime(2025, 1, 31)
)
```