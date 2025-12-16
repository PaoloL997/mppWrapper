import os
import win32com.client
from datetime import datetime


class MSProjectWrapper:
    """Wrapper class for interacting with Microsoft Project files.
    
    This class provides methods to read, create, modify, and delete tasks and resources
    in Microsoft Project files using the win32com client.
    """
    
    def __init__(self, path: str):
        """Initialize the MSProject wrapper and open a project file. 

        Args:
            path: The file path to the Microsoft Project file (.mpp).
        Raises:
            FileNotFoundError: If the specified project file does not exist.
        """
        self.app = win32com.client.Dispatch("MSProject.Application")
        self.app.Visible = True
        
        if not os.path.exists(path):
            raise FileNotFoundError(f"{path} not found.")
        self.app.FileOpen(path)
        self.project = self.app.ActiveProject
    
    def get_tasks(self):
        """Retrieve all tasks from the project.
        
        Returns:
            A list of dictionaries containing task information including:
                - task: Task name
                - gerarchia: Outline number (i.e., x = parent task, x.y = subtask)
                - inizio: Start date
                - fine: Finish date
                - stabilimento: Warehouse (Text1 field)
                - responsabile: Manager (Text2 field)
                - risorse: List of assigned resource names
        """
        out = []
        for task in self.project.Tasks:
            if task is not None:
                resources = []
                for res in task.Assignments:
                    resources.append(res.ResourceName)
                out.append(
                    {
                        "task": task.Name,
                        "gerarchia":task.OutlineNumber,
                        "inizio": task.Start,
                        "fine": task.Finish,
                        "stabilimento": task.Text1,
                        "responsabile": task.Text2,
                        "risorse": resources
                    }
                )
        return out

    def get_resources(self):
        """Retrieve all resources from the project.
        
        Returns:
            A list of dictionaries containing resource information including:
                - risorsa: Resource name
                - passo: Pitch (Number1 field)
                - distanza_interasse: Center-to-center distance (Number2 field)
                - diametro: Diameter (Number3 field)
                - stabilimento: Warehouse (Text1 field)
        """
        out = []
        for res in self.project.Resources:
            if res is not None:
                out.append(
                    {
                        "risorsa":res.Name,
                        "passo": res.Number1,
                        "distanza_interasse": res.Number2,
                        "diametro": res.Number3,
                        "stabilimento": res.Text1
                    }
                )
        return out
    
    def get_task(self, name: str):
        """Retrieve a specific task by name.
        
        Args:
            name: The name of the task to retrieve. 
        Returns:
            The task object if found, or a NameError if the task cannot be found.
        """
        try:
            task = self.project.Tasks(name)
            if task is not None:
                return task
        except Exception as e:
            return NameError(f"Unable to find task {name}: {e}")

    def get_resource_id(self, resource_name: str):
        """Get the ID of a resource by its name.
        
        Args:
            resource_name: The name of the resource.  
        Returns:
            The resource ID if found, None otherwise.
        """
        for res in self.project.Resources:
            if res is not None and res.Name == resource_name:
                return res.ID
        return None
    
    def get_task_id(self, task_name: str):
        """Get the ID of a task by its name.
        
        Args:
            task_name: The name of the task.
        Returns:
            The task ID if found, None otherwise.
        """
        for task in self.project.Tasks:
            if task is not None and task.Name == task_name:
                return task.ID
        return None
    
    def add_task(
            self,
            name: str,
            start: datetime,
            end: datetime,
            warehouse: str,
            manager: str,
            parent: win32com.client.CDispatch | None = None
            ):
        """Add a new task to the project.
        
        Args:
            name: The name of the new task.
            start: The start date of the task.
            end: The finish date of the task.
            warehouse: The warehouse identifier (stored in Text1 field).
            manager: The manager name (stored in Text2 field).
            parent: Optional parent task to indent under. If None, creates a top-level task.
        """
        new_task = self.project.Tasks.Add(name)
        new_task.Start = start
        new_task.Finish = end
        new_task.Text1 = warehouse
        new_task.Text2 = manager

        if parent: 
            while new_task.OutlineLevel < parent.OutlineLevel:
                new_task.OutlineIndent()
        else:
            while new_task.OutlineLevel > 1:
                new_task.OutlineOutdent() # Altrimenti utilizza quella precedente

    def add_resources_to_task(
            self,
            task: win32com.client.CDispatch,
            resourceIDs: list[int]
            ):
        """Assign resources to a task.
        
        Args:
            task: The task object to assign resources to.
            resourceIDs: A list of resource IDs to assign to the task.
        """
        for res in resourceIDs:
            task.Assignments.Add(task.ID, res)

    def add_resource(
            self,
            name: str,
            warehouse: str,
            diameter: float | None = None,
            pitch: float | None = None,
            center_to_center: float | None = None,
            ):
        """Add a new resource to the project.
        
        Args:
            name: The name of the new resource.
            warehouse: The warehouse identifier (stored in Text1 field).
            diameter: Optional diameter value (stored in Number3 field).
            pitch: Optional pitch value (stored in Number1 field).
            center_to_center: Optional center-to-center distance (stored in Number2 field).
        """
        new_resource = self.project.Resources.Add(name)
        if pitch:
            new_resource.Number1 = pitch
        if center_to_center:
            new_resource.Number2 = center_to_center
        if diameter:
            new_resource.Number3 = diameter
        new_resource.Text1 = warehouse

    def check_resource_availability(
            self,
            resource_id: int,
            start: datetime,
            end: datetime
        ):
        """Check if a resource is available during a specified time period.
        
        Args:
            resource_id: The ID of the resource to check.
            start: The start date and time of the period to check.
            end: The end date and time of the period to check.
            
        Returns:
            True if the resource is available (no conflicting assignments),
            False if there are overlapping task assignments.
        """
        resource = self.project.Resources(resource_id)
        for assignments in resource.Assignments:
            task = assignments.Task
            if task is None:
                continue
            task_start = task.Start.replace(tzinfo=None)
            task_end = task.Finish.replace(tzinfo=None)
            if (start < task_end) and (end > task_start):
                return False
        return True
        
    def delete_task(self, task_id: int):
        """Delete a task from the project.
        
        Args:
            task_id: The ID of the task to delete.
        Returns:
            True if the task was successfully deleted, False if the task was not found.
        """
        task = self.project.Tasks(task_id)
        if task:
            task.delete()
            return True
        return False

    def delete_resource(self, resource_id: int):
        """Delete a resource from the project.
        
        Args:
            resource_id: The ID of the resource to delete.
        Returns:
            True if the resource was successfully deleted, False if the resource was not found.
        """
        res = self.project.Resources(resource_id)
        if res:
            res.delete()
            return True
        return False
        