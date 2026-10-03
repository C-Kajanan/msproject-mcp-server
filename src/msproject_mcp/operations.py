"""Operations against the MS Project COM object model.

Every public function takes the Project ``Application`` COM object as its first
argument and returns plain JSON-serialisable data. They run on the COM thread
(see :mod:`msproject_mcp.com_bridge`) and are free of MCP concerns, which keeps
them testable against a fake object model.
"""

from __future__ import annotations

import datetime as dt
import os
from typing import Any, Callable, Iterator

from . import constants as c


class ProjectError(ValueError):
    """A user-facing error (bad ID, no project open, invalid argument...)."""


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def _safe(getter: Callable[[], Any], default: Any = None) -> Any:
    """Read a COM property that may raise for some task types."""
    try:
        return getter()
    except Exception:
        return default


def _iso(value: Any) -> str | None:
    """Format a COM date. Project returns the string 'NA' for empty dates."""
    if isinstance(value, dt.datetime):
        # pywin32 may tag VT_DATE values as UTC even though Project dates are
        # local wall-clock times; drop tzinfo so the value is shown verbatim.
        return value.replace(tzinfo=None).isoformat(timespec="minutes")
    if isinstance(value, dt.date):
        return value.isoformat()
    return None


def _parse_date(value: str, field: str) -> dt.datetime:
    try:
        return dt.datetime.fromisoformat(value)
    except (TypeError, ValueError) as exc:
        raise ProjectError(
            f"Invalid {field} {value!r}; use ISO format, e.g. 2026-10-05 or 2026-10-05T08:00"
        ) from exc


def _num(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _minutes_per_day(project: Any) -> float:
    hours = _num(_safe(lambda: project.HoursPerDay)) or 8.0
    return hours * 60.0


def _minutes_to_days(minutes: Any, project: Any) -> float | None:
    m = _num(minutes)
    if m is None:
        return None
    return round(m / _minutes_per_day(project), 3)


def _minutes_to_hours(minutes: Any) -> float | None:
    m = _num(minutes)
    return None if m is None else round(m / 60.0, 3)


def _projects_count(app: Any) -> int:
    return int(_safe(lambda: app.Projects.Count, 0) or 0)


def _get_project(app: Any, project_name: str | None = None) -> Any:
    if _projects_count(app) == 0:
        raise ProjectError("No project is open. Use open_project or create_project first.")
    if project_name:
        for p in app.Projects:
            if p.Name == project_name or p.FullName == project_name:
                return p
        raise ProjectError(f"No open project named {project_name!r}.")
    return app.ActiveProject


def _iter_tasks(project: Any) -> Iterator[Any]:
    # Blank rows in the Gantt view come back as None.
    for t in project.Tasks:
        if t is not None:
            yield t


def _iter_resources(project: Any) -> Iterator[Any]:
    for r in project.Resources:
        if r is not None:
            yield r


def _get_task(project: Any, task_id: int) -> Any:
    task = _safe(lambda: project.Tasks(task_id))
    if task is None:
        raise ProjectError(f"No task with ID {task_id} in project {project.Name!r}.")
    return task


def _get_resource(project: Any, resource: int | str) -> Any:
    if isinstance(resource, str) and not resource.isdigit():
        for r in _iter_resources(project):
            if r.Name.lower() == resource.lower():
                return r
        raise ProjectError(f"No resource named {resource!r} in project {project.Name!r}.")
    res = _safe(lambda: project.Resources(int(resource)))
    if res is None:
        raise ProjectError(f"No resource with ID {resource} in project {project.Name!r}.")
    return res


def _project_summary(project: Any, active_name: str | None = None) -> dict[str, Any]:
    return {
        "name": project.Name,
        "full_name": _safe(lambda: project.FullName),
        "saved": bool(_safe(lambda: project.Saved, False)),
        "active": project.Name == active_name if active_name else None,
    }


def _task_dict(task: Any, project: Any, detailed: bool = False) -> dict[str, Any]:
    constraint = _safe(lambda: task.ConstraintType)
    data: dict[str, Any] = {
        "id": task.ID,
        "unique_id": _safe(lambda: task.UniqueID),
        "name": task.Name,
        "wbs": _safe(lambda: task.WBS),
        "outline_level": _safe(lambda: task.OutlineLevel),
        "summary": bool(_safe(lambda: task.Summary, False)),
        "milestone": bool(_safe(lambda: task.Milestone, False)),
        "critical": bool(_safe(lambda: task.Critical, False)),
        "manually_scheduled": bool(_safe(lambda: task.Manual, False)),
        "start": _iso(_safe(lambda: task.Start)),
        "finish": _iso(_safe(lambda: task.Finish)),
        "duration_days": _minutes_to_days(_safe(lambda: task.Duration), project),
        "percent_complete": _safe(lambda: task.PercentComplete),
        "predecessors": _safe(lambda: task.Predecessors) or "",
        "resource_names": _safe(lambda: task.ResourceNames) or "",
    }
    if detailed:
        data.update(
            {
                "successors": _safe(lambda: task.Successors) or "",
                "active": bool(_safe(lambda: task.Active, True)),
                "priority": _safe(lambda: task.Priority),
                "constraint_type": c.CONSTRAINT_TYPE_NAMES.get(constraint, constraint),
                "constraint_date": _iso(_safe(lambda: task.ConstraintDate)),
                "deadline": _iso(_safe(lambda: task.Deadline)),
                "actual_start": _iso(_safe(lambda: task.ActualStart)),
                "actual_finish": _iso(_safe(lambda: task.ActualFinish)),
                "baseline_start": _iso(_safe(lambda: task.BaselineStart)),
                "baseline_finish": _iso(_safe(lambda: task.BaselineFinish)),
                "total_slack_days": _minutes_to_days(_safe(lambda: task.TotalSlack), project),
                "work_hours": _minutes_to_hours(_safe(lambda: task.Work)),
                "cost": _num(_safe(lambda: task.Cost)),
                "notes": _safe(lambda: task.Notes) or "",
            }
        )
    return data


def _resource_dict(res: Any) -> dict[str, Any]:
    rtype = _safe(lambda: res.Type)
    return {
        "id": res.ID,
        "unique_id": _safe(lambda: res.UniqueID),
        "name": res.Name,
        "type": c.RESOURCE_TYPE_NAMES.get(rtype, rtype),
        "initials": _safe(lambda: res.Initials),
        "group": _safe(lambda: res.Group),
        "email": _safe(lambda: res.EMailAddress),
        "max_units": _num(_safe(lambda: res.MaxUnits)),
        "standard_rate": _safe(lambda: res.StandardRate),
        "work_hours": _minutes_to_hours(_safe(lambda: res.Work)),
        "cost": _num(_safe(lambda: res.Cost)),
        "overallocated": bool(_safe(lambda: res.Overallocated, False)),
    }


def _assignment_dict(a: Any) -> dict[str, Any]:
    return {
        "unique_id": _safe(lambda: a.UniqueID),
        "task_id": _safe(lambda: a.TaskID),
        "task_name": _safe(lambda: a.TaskName),
        "resource_id": _safe(lambda: a.ResourceID),
        "resource_name": _safe(lambda: a.ResourceName),
        "units": _num(_safe(lambda: a.Units)),
        "work_hours": _minutes_to_hours(_safe(lambda: a.Work)),
        "start": _iso(_safe(lambda: a.Start)),
        "finish": _iso(_safe(lambda: a.Finish)),
    }


def _link_type(link_type: str) -> int:
    key = link_type.upper()
    if key not in c.LINK_TYPES:
        raise ProjectError(f"Invalid link_type {link_type!r}; use one of FS, SS, FF, SF.")
    return c.LINK_TYPES[key]


def _apply_task_fields(
    task: Any,
    *,
    name: str | None = None,
    duration: str | None = None,
    start: str | None = None,
    finish: str | None = None,
    percent_complete: int | None = None,
    notes: str | None = None,
    priority: int | None = None,
    manually_scheduled: bool | None = None,
    active: bool | None = None,
    milestone: bool | None = None,
    deadline: str | None = None,
    constraint_type: str | None = None,
    constraint_date: str | None = None,
) -> None:
    # Order matters: scheduling mode first so dates/durations are interpreted
    # under the intended mode.
    if manually_scheduled is not None:
        task.Manual = manually_scheduled
    if name is not None:
        task.Name = name
    if duration is not None:
        # Project parses duration strings such as "5d", "3h", "2w", "0d".
        task.Duration = duration
    if start is not None:
        task.Start = _parse_date(start, "start")
    if finish is not None:
        task.Finish = _parse_date(finish, "finish")
    if milestone is not None:
        task.Milestone = milestone
    if percent_complete is not None:
        if not 0 <= percent_complete <= 100:
            raise ProjectError("percent_complete must be between 0 and 100.")
        task.PercentComplete = percent_complete
    if notes is not None:
        task.Notes = notes
    if priority is not None:
        if not 0 <= priority <= 1000:
            raise ProjectError("priority must be between 0 and 1000.")
        task.Priority = priority
    if active is not None:
        task.Active = active
    if deadline is not None:
        task.Deadline = _parse_date(deadline, "deadline") if deadline else "NA"
    if constraint_type is not None:
        key = constraint_type.lower()
        if key not in c.CONSTRAINT_TYPES:
            raise ProjectError(
                f"Invalid constraint_type {constraint_type!r}; use one of "
                + ", ".join(c.CONSTRAINT_TYPES)
            )
        task.ConstraintType = c.CONSTRAINT_TYPES[key]
    if constraint_date is not None:
        task.ConstraintDate = _parse_date(constraint_date, "constraint_date")


# --------------------------------------------------------------------------
# Application / file operations
# --------------------------------------------------------------------------


def get_status(app: Any) -> dict[str, Any]:
    active = _safe(lambda: app.ActiveProject.Name) if _projects_count(app) else None
    return {
        "application": _safe(lambda: app.Name),
        "version": _safe(lambda: app.Version),
        "visible": bool(_safe(lambda: app.Visible, False)),
        "active_project": active,
        "open_projects": [_project_summary(p, active) for p in app.Projects],
    }


def open_project(app: Any, path: str, read_only: bool = False) -> dict[str, Any]:
    full = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(full):
        raise ProjectError(f"File not found: {full}")
    ok = app.FileOpenEx(Name=full, ReadOnly=read_only)
    if ok is False:
        raise ProjectError(f"Microsoft Project could not open {full}.")
    return get_project_info(app)


def create_project(
    app: Any, start_date: str | None = None, title: str | None = None
) -> dict[str, Any]:
    project = app.Projects.Add(False)
    project.Activate()
    if start_date:
        project.ProjectStart = _parse_date(start_date, "start_date")
    if title:
        project.ProjectSummaryTask.Name = title
    return get_project_info(app)


def activate_project(app: Any, project_name: str) -> dict[str, Any]:
    _get_project(app, project_name).Activate()
    return get_project_info(app)


def save_project(
    app: Any,
    path: str | None = None,
    file_format: str = "mpp",
    project_name: str | None = None,
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    project.Activate()
    if path:
        fmt = c.SAVE_FORMATS.get(file_format.lower())
        if fmt is None:
            raise ProjectError(
                f"Unsupported file_format {file_format!r}; use one of {', '.join(c.SAVE_FORMATS)}."
            )
        full = os.path.abspath(os.path.expanduser(path))
        app.FileSaveAs(Name=full, FormatID=fmt)
    else:
        if not _safe(lambda: project.Path):
            raise ProjectError("Project has never been saved; provide a path.")
        app.FileSave()
    return {"saved": True, "full_name": _safe(lambda: app.ActiveProject.FullName)}


def export_pdf(app: Any, path: str, project_name: str | None = None) -> dict[str, Any]:
    _get_project(app, project_name).Activate()
    full = os.path.abspath(os.path.expanduser(path))
    app.DocumentExport(FileName=full, FileType=c.PJ_PDF)
    return {"exported": True, "path": full}


def close_project(
    app: Any, save: bool = False, project_name: str | None = None
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    name = project.Name
    project.Activate()
    app.FileCloseEx(Save=c.PJ_SAVE if save else c.PJ_DO_NOT_SAVE)
    return {"closed": name, "saved": save}


def get_project_info(app: Any, project_name: str | None = None) -> dict[str, Any]:
    project = _get_project(app, project_name)
    summary = _safe(lambda: project.ProjectSummaryTask)
    tasks = list(_iter_tasks(project))
    return {
        "name": project.Name,
        "full_name": _safe(lambda: project.FullName),
        "title": _safe(lambda: summary.Name) if summary is not None else None,
        "start": _iso(_safe(lambda: project.ProjectStart)),
        "finish": _iso(_safe(lambda: project.ProjectFinish)),
        "status_date": _iso(_safe(lambda: project.StatusDate)),
        "calendar": _safe(lambda: project.Calendar.Name),
        "hours_per_day": _num(_safe(lambda: project.HoursPerDay)),
        "percent_complete": _safe(lambda: summary.PercentComplete) if summary else None,
        "duration_days": _minutes_to_days(_safe(lambda: summary.Duration), project)
        if summary
        else None,
        "cost": _num(_safe(lambda: summary.Cost)) if summary else None,
        "task_count": len(tasks),
        "resource_count": sum(1 for _ in _iter_resources(project)),
        "saved": bool(_safe(lambda: project.Saved, False)),
    }


def calculate_project(app: Any, project_name: str | None = None) -> dict[str, Any]:
    _get_project(app, project_name).Activate()
    app.CalculateProject()
    return get_project_info(app)


def save_baseline(app: Any, project_name: str | None = None) -> dict[str, Any]:
    _get_project(app, project_name).Activate()
    app.BaselineSave(All=True)
    return {"baseline_saved": True}


# --------------------------------------------------------------------------
# Tasks
# --------------------------------------------------------------------------


def list_tasks(
    app: Any,
    name_contains: str | None = None,
    include_summary: bool = True,
    critical_only: bool = False,
    max_outline_level: int | None = None,
    offset: int = 0,
    limit: int = 200,
    project_name: str | None = None,
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    needle = name_contains.lower() if name_contains else None
    matched = []
    for t in _iter_tasks(project):
        if needle and needle not in (t.Name or "").lower():
            continue
        if not include_summary and _safe(lambda: t.Summary, False):
            continue
        if critical_only and not _safe(lambda: t.Critical, False):
            continue
        if max_outline_level is not None and (_safe(lambda: t.OutlineLevel, 1) or 1) > max_outline_level:
            continue
        matched.append(t)
    if offset < 0 or limit < 0:
        raise ProjectError("offset and limit must be non-negative.")
    page = matched[offset : offset + limit]
    return {
        "project": project.Name,
        "total": len(matched),
        "offset": offset,
        "returned": len(page),
        "tasks": [_task_dict(t, project) for t in page],
    }


def get_task(app: Any, task_id: int, project_name: str | None = None) -> dict[str, Any]:
    project = _get_project(app, project_name)
    task = _get_task(project, task_id)
    data = _task_dict(task, project, detailed=True)
    data["assignments"] = [_assignment_dict(a) for a in task.Assignments]
    return data


def _last_descendant_position(project: Any, parent: Any) -> int | None:
    """ID to insert *before* so a new task becomes the parent's last child.

    Returns None when the parent's subtree runs to the end of the project.
    """
    level = parent.OutlineLevel
    for t in _iter_tasks(project):
        if t.ID > parent.ID and t.OutlineLevel <= level:
            return t.ID
    return None


def add_task(
    app: Any,
    name: str,
    duration: str | None = None,
    start: str | None = None,
    parent_id: int | None = None,
    before_id: int | None = None,
    predecessors: list[int] | None = None,
    notes: str | None = None,
    manually_scheduled: bool = False,
    project_name: str | None = None,
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    if parent_id is not None and before_id is not None:
        raise ProjectError("Specify parent_id or before_id, not both.")

    parent = _get_task(project, parent_id) if parent_id is not None else None
    position = before_id
    if parent is not None:
        position = _last_descendant_position(project, parent)
    elif before_id is not None:
        _get_task(project, before_id)

    task = project.Tasks.Add(name, position) if position is not None else project.Tasks.Add(name)
    if parent is not None:
        task.OutlineLevel = parent.OutlineLevel + 1

    _apply_task_fields(
        task,
        manually_scheduled=manually_scheduled,
        duration=duration,
        start=start,
        notes=notes,
    )
    for pred_id in predecessors or []:
        task.LinkPredecessors(_get_task(project, pred_id), c.LINK_TYPES["FS"])
    return _task_dict(task, project, detailed=True)


def update_task(app: Any, task_id: int, project_name: str | None = None, **fields: Any) -> dict[str, Any]:
    project = _get_project(app, project_name)
    task = _get_task(project, task_id)
    _apply_task_fields(task, **{k: v for k, v in fields.items() if v is not None})
    return _task_dict(task, project, detailed=True)


def delete_task(app: Any, task_id: int, project_name: str | None = None) -> dict[str, Any]:
    project = _get_project(app, project_name)
    task = _get_task(project, task_id)
    name = task.Name
    is_summary = bool(_safe(lambda: task.Summary, False))
    task.Delete()
    return {
        "deleted": {"id": task_id, "name": name},
        "note": ("Subtasks were deleted too. " if is_summary else "")
        + "Task IDs after the deleted row have shifted.",
    }


def indent_task(app: Any, task_id: int, project_name: str | None = None) -> dict[str, Any]:
    project = _get_project(app, project_name)
    task = _get_task(project, task_id)
    task.OutlineIndent()
    return _task_dict(task, project)


def outdent_task(app: Any, task_id: int, project_name: str | None = None) -> dict[str, Any]:
    project = _get_project(app, project_name)
    task = _get_task(project, task_id)
    task.OutlineOutdent()
    return _task_dict(task, project)


def link_tasks(
    app: Any,
    predecessor_id: int,
    successor_id: int,
    link_type: str = "FS",
    lag: str | None = None,
    project_name: str | None = None,
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    if predecessor_id == successor_id:
        raise ProjectError("A task cannot be linked to itself.")
    pred = _get_task(project, predecessor_id)
    succ = _get_task(project, successor_id)
    if lag:
        succ.LinkPredecessors(pred, _link_type(link_type), lag)
    else:
        succ.LinkPredecessors(pred, _link_type(link_type))
    return _task_dict(succ, project)


def unlink_tasks(
    app: Any, predecessor_id: int, successor_id: int, project_name: str | None = None
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    pred = _get_task(project, predecessor_id)
    succ = _get_task(project, successor_id)
    succ.UnlinkPredecessors(pred)
    return _task_dict(succ, project)


def get_task_dependencies(
    app: Any, task_id: int, project_name: str | None = None
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    task = _get_task(project, task_id)
    deps = []
    for d in task.TaskDependencies:
        deps.append(
            {
                "from_id": d.From.ID,
                "from_name": d.From.Name,
                "to_id": d.To.ID,
                "to_name": d.To.Name,
                "type": c.LINK_TYPE_NAMES.get(d.Type, d.Type),
                "lag_days": _minutes_to_days(_safe(lambda: d.Lag), project),
            }
        )
    return {"task_id": task_id, "dependencies": deps}


def get_critical_path(app: Any, project_name: str | None = None) -> dict[str, Any]:
    project = _get_project(app, project_name)
    tasks = [
        _task_dict(t, project)
        for t in _iter_tasks(project)
        if _safe(lambda: t.Critical, False) and not _safe(lambda: t.Summary, False)
    ]
    return {"project": project.Name, "critical_tasks": tasks}


# --------------------------------------------------------------------------
# Resources & assignments
# --------------------------------------------------------------------------


def list_resources(app: Any, project_name: str | None = None) -> dict[str, Any]:
    project = _get_project(app, project_name)
    return {
        "project": project.Name,
        "resources": [_resource_dict(r) for r in _iter_resources(project)],
    }


def _apply_resource_fields(
    res: Any,
    *,
    name: str | None = None,
    resource_type: str | None = None,
    max_units: float | None = None,
    standard_rate: str | None = None,
    initials: str | None = None,
    group: str | None = None,
    email: str | None = None,
) -> None:
    if name is not None:
        res.Name = name
    if resource_type is not None:
        key = resource_type.lower()
        if key not in c.RESOURCE_TYPES:
            raise ProjectError("resource_type must be one of work, material, cost.")
        res.Type = c.RESOURCE_TYPES[key]
    if max_units is not None:
        res.MaxUnits = max_units
    if standard_rate is not None:
        # Project accepts rate strings such as "50/h" or "400/d".
        res.StandardRate = standard_rate
    if initials is not None:
        res.Initials = initials
    if group is not None:
        res.Group = group
    if email is not None:
        res.EMailAddress = email


def add_resource(
    app: Any,
    name: str,
    resource_type: str = "work",
    max_units: float | None = None,
    standard_rate: str | None = None,
    initials: str | None = None,
    group: str | None = None,
    email: str | None = None,
    project_name: str | None = None,
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    res = project.Resources.Add(name)
    _apply_resource_fields(
        res,
        resource_type=resource_type,
        max_units=max_units,
        standard_rate=standard_rate,
        initials=initials,
        group=group,
        email=email,
    )
    return _resource_dict(res)


def update_resource(
    app: Any, resource: int | str, project_name: str | None = None, **fields: Any
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    res = _get_resource(project, resource)
    _apply_resource_fields(res, **{k: v for k, v in fields.items() if v is not None})
    return _resource_dict(res)


def delete_resource(
    app: Any, resource: int | str, project_name: str | None = None
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    res = _get_resource(project, resource)
    info = {"id": res.ID, "name": res.Name}
    res.Delete()
    return {"deleted": info}


def assign_resource(
    app: Any,
    task_id: int,
    resource: int | str,
    units: float = 1.0,
    project_name: str | None = None,
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    task = _get_task(project, task_id)
    if _safe(lambda: task.Summary, False):
        raise ProjectError("Resources should be assigned to detail tasks, not summary tasks.")
    res = _get_resource(project, resource)
    assignment = task.Assignments.Add(ResourceID=res.ID, Units=units)
    return _assignment_dict(assignment)


def list_assignments(
    app: Any,
    task_id: int | None = None,
    resource: int | str | None = None,
    project_name: str | None = None,
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    if task_id is not None:
        source = _get_task(project, task_id).Assignments
    elif resource is not None:
        source = _get_resource(project, resource).Assignments
    else:
        source = [a for t in _iter_tasks(project) for a in t.Assignments]
    return {"assignments": [_assignment_dict(a) for a in source]}


def remove_assignment(
    app: Any, task_id: int, resource: int | str, project_name: str | None = None
) -> dict[str, Any]:
    project = _get_project(app, project_name)
    task = _get_task(project, task_id)
    res = _get_resource(project, resource)
    for a in task.Assignments:
        if a.ResourceID == res.ID:
            a.Delete()
            return {"removed": {"task_id": task_id, "resource": res.Name}}
    raise ProjectError(f"Resource {res.Name!r} is not assigned to task {task_id}.")
