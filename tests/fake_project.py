"""Minimal in-memory imitation of the MS Project COM object model."""

from __future__ import annotations

import datetime as dt


class FakeCollection(list):
    """1-based, callable like COM collections."""

    def __call__(self, index):
        if not 1 <= index <= len(self):
            raise IndexError(index)
        return self[index - 1]

    @property
    def Count(self):
        return len(self)


class FakeAssignment:
    def __init__(self, task, resource, units):
        self.task, self.resource, self.Units = task, resource, units
        self.UniqueID = id(self) % 10000
        self.Work = 480

    TaskID = property(lambda s: s.task.ID)
    TaskName = property(lambda s: s.task.Name)
    ResourceID = property(lambda s: s.resource.ID)
    ResourceName = property(lambda s: s.resource.Name)
    Start = property(lambda s: s.task.Start)
    Finish = property(lambda s: s.task.Finish)

    def Delete(self):
        self.task.Assignments.remove(self)
        self.resource.Assignments.remove(self)


class FakeAssignments(FakeCollection):
    def __init__(self, task):
        super().__init__()
        self.task = task

    def Add(self, ResourceID=None, Units=1.0):
        res = self.task.project.Resources(ResourceID)
        a = FakeAssignment(self.task, res, Units)
        self.append(a)
        res.Assignments.append(a)
        return a


class FakeDependency:
    def __init__(self, frm, to, type_, lag):
        self.From, self.To, self.Type, self.Lag = frm, to, type_, lag


class FakeTask:
    _uid = 0

    def __init__(self, project, name):
        FakeTask._uid += 1
        self.project = project
        self.UniqueID = FakeTask._uid
        self.Name = name
        self.OutlineLevel = 1
        self.Duration = 480
        self.Start = dt.datetime(2026, 10, 5, 8, 0)
        self.Finish = dt.datetime(2026, 10, 5, 17, 0)
        self.PercentComplete = 0
        self.Manual = True
        self.Milestone = False
        self.Critical = False
        self.Notes = ""
        self.Priority = 500
        self.Active = True
        self.ConstraintType = 0
        self.ConstraintDate = "NA"
        self.Deadline = "NA"
        self.ResourceNames = ""
        self.Assignments = FakeAssignments(self)
        self.preds = []  # list of (task, type, lag)

    @property
    def ID(self):
        return self.project.Tasks.index(self) + 1

    @property
    def Duration(self):
        return self._duration

    @Duration.setter
    def Duration(self, value):
        # Real Project parses strings like "5d"; support days/hours here.
        if isinstance(value, str):
            n, unit = float(value[:-1]), value[-1]
            value = n * (480 if unit == "d" else 60)
        self._duration = value

    @property
    def Summary(self):
        tasks = self.project.Tasks
        i = tasks.index(self)
        return i + 1 < len(tasks) and tasks[i + 1].OutlineLevel > self.OutlineLevel

    @property
    def Predecessors(self):
        return ",".join(str(p.ID) for p, _, _ in self.preds)

    @property
    def TaskDependencies(self):
        deps = [FakeDependency(p, self, t, lag) for p, t, lag in self.preds]
        for t in self.project.Tasks:
            for p, ty, lag in t.preds:
                if p is self:
                    deps.append(FakeDependency(self, t, ty, lag))
        return deps

    def LinkPredecessors(self, task, link=1, lag=0):
        self.preds.append((task, link, lag))

    def UnlinkPredecessors(self, task):
        self.preds = [p for p in self.preds if p[0] is not task]

    def OutlineIndent(self):
        self.OutlineLevel += 1

    def OutlineOutdent(self):
        self.OutlineLevel = max(1, self.OutlineLevel - 1)

    def Delete(self):
        self.project.Tasks.remove(self)


class FakeTasks(FakeCollection):
    def __init__(self, project):
        super().__init__()
        self.project = project

    def Add(self, Name, Before=None):
        t = FakeTask(self.project, Name)
        if Before is None:
            self.append(t)
        else:
            self.insert(Before - 1, t)
        return t


class FakeResource:
    def __init__(self, project, name):
        self.project = project
        self.Name = name
        self.UniqueID = len(project.Resources) + 1
        self.Type = 0
        self.MaxUnits = 1.0
        self.StandardRate = "$0.00/hr"
        self.Initials = name[:1]
        self.Group = ""
        self.EMailAddress = ""
        self.Work = 0
        self.Cost = 0
        self.Overallocated = False
        self.Assignments = FakeCollection()

    @property
    def ID(self):
        return self.project.Resources.index(self) + 1

    def Delete(self):
        self.project.Resources.remove(self)


class FakeResources(FakeCollection):
    def __init__(self, project):
        super().__init__()
        self.project = project

    def Add(self, Name, Before=None):
        r = FakeResource(self.project, Name)
        self.append(r)
        return r


class FakeProject:
    def __init__(self, app, name):
        self.app = app
        self.Name = name
        self.FullName = name
        self.Path = ""
        self.Saved = True
        self.HoursPerDay = 8
        self.ProjectStart = dt.datetime(2026, 10, 5, 8, 0)
        self.ProjectFinish = dt.datetime(2026, 10, 30, 17, 0)
        self.StatusDate = "NA"
        self.Tasks = FakeTasks(self)
        self.Resources = FakeResources(self)
        self.ProjectSummaryTask = FakeTask(self, name)

    def Activate(self):
        self.app.ActiveProject = self


class FakeApp:
    Name = "Microsoft Project"
    Version = "16.0"

    def __init__(self):
        self.Visible = True
        self.Projects = FakeCollection()
        self.ActiveProject = None
        self.calls = []
        projects = self.Projects
        app = self

        def add(*args):
            p = FakeProject(app, f"Project{len(projects) + 1}")
            projects.append(p)
            return p

        projects.Add = add

    def FileSaveAs(self, **kw):
        self.calls.append(("FileSaveAs", kw))
        self.ActiveProject.FullName = kw["Name"]
        self.ActiveProject.Path = kw["Name"]

    def FileSave(self):
        self.calls.append(("FileSave", {}))

    def FileCloseEx(self, Save=0):
        self.calls.append(("FileCloseEx", {"Save": Save}))
        self.Projects.remove(self.ActiveProject)
        self.ActiveProject = self.Projects[0] if self.Projects else None

    def FileOpenEx(self, Name, ReadOnly=False):
        p = self.Projects.Add()
        p.Name = Name.rsplit("/", 1)[-1]
        p.FullName = p.Path = Name
        p.Activate()
        return True

    def CalculateProject(self):
        self.calls.append(("CalculateProject", {}))

    def BaselineSave(self, **kw):
        self.calls.append(("BaselineSave", kw))

    def DocumentExport(self, **kw):
        self.calls.append(("DocumentExport", kw))
