---
name: build-schedule
description: >
  This skill should be used when the user asks to "build a schedule in MS Project",
  "create a project plan", "add tasks to my .mpp", "link these tasks", "set a baseline",
  "find the critical path", or otherwise wants Microsoft Project desktop driven through
  the msproject tools.
---

# Build and edit schedules in Microsoft Project

Drive Microsoft Project desktop through the `msproject` MCP tools. The tools act on the
copy of Project running on the user's Windows computer, so every change is visible in the
Project window as it happens.

## Before changing anything

1. Call `get_status`. It attaches to Project (starting it if needed) and lists the open
   projects. If it reports that Project is not available, tell the user the plugin needs
   Windows with Microsoft Project desktop installed, and stop.
2. Work on the right file. Use `open_project` with an absolute Windows path, or
   `create_project` for a new one. Tools act on the active project unless `project_name`
   is passed.
3. When editing an existing file, read it first with `get_project_info` and `list_tasks`
   so the user's structure is understood before it is changed.

## Conventions the tools expect

- Tasks are addressed by ID, the row number shown in Project. IDs shift whenever a row is
  inserted or deleted. Call `list_tasks` again after any structural change and never
  reuse IDs from before it. `unique_id` does not change.
- Durations and lags are strings: `"5d"`, `"4h"`, `"2w"`, `"0d"` for a milestone,
  `"-3d"` for a lead.
- Dates are ISO 8601: `2026-10-05` or `2026-10-05T08:00`.
- Units: `1.0` is 100%.

## Building a schedule

Work top-down and in this order, because links and assignments need the task IDs to be
settled first.

1. Create the project with its start date. Do not set start dates on individual tasks
   unless the user gives a fixed date: passing `start` to `add_task` puts a constraint
   on an auto-scheduled task, and constrained tasks stop following their predecessors.
2. Add the summary tasks, then add each subtask with `parent_id` set to its summary
   task. This appends the subtask as the last child, so add children in order. Use
   `before_id` to insert a row above an existing one. `indent_task` and `outdent_task`
   fix the outline level of a task that is already in place.
3. Re-read with `list_tasks`, then link. Use `link_tasks` for anything other than a
   plain finish-to-start link, or when a lag or lead is needed. Link the working tasks,
   not the summary tasks.
4. Add resources with `add_resource`, then `assign_resource`.
5. Call `calculate_project`, then check the result with `get_project_info` and
   `get_critical_path`. Compare the finish date with what the user expected and say so
   if it differs.
6. Call `save_baseline` only when the user asks for a baseline or the plan is agreed.
   Saving a baseline over an existing one replaces it.
7. Save with `save_project`. Nothing is written to disk until this is called. A project
   that has never been saved needs a `path`.

## Things to check before reporting success

- Every non-summary task except the first has a predecessor, unless the user wanted it
  unlinked. Tasks with no predecessor start on the project start date.
- Milestones have a duration of `"0d"`.
- The file was saved, and the reply states the full path it was saved to.

## When something goes wrong

- A call that hangs usually means Project is showing a dialog. Ask the user to look at
  the Project window and dismiss it.
- `delete_task` on a summary task deletes its subtasks too. Confirm with the user before
  deleting a summary task.
- `close_project` discards unsaved changes unless `save` is true.
- The tools cover tasks, links, resources, assignments, baselines and files. They do not
  edit calendars or working-time exceptions. If the user needs holidays or a custom
  calendar, say that it has to be set in Project itself (Project > Change Working Time).
