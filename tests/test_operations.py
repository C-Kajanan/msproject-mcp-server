import pytest

from msproject_mcp import operations as ops
from fake_project import FakeApp


@pytest.fixture
def app():
    a = FakeApp()
    ops.create_project(a, start_date="2026-10-05", title="Website")
    return a


def test_no_project_open_raises():
    with pytest.raises(ops.ProjectError, match="No project is open"):
        ops.list_tasks(FakeApp())


def test_create_project_info(app):
    info = ops.get_project_info(app)
    assert info["title"] == "Website"
    assert info["start"] == "2026-10-05T00:00"
    assert info["task_count"] == 0


def test_add_and_list_tasks(app):
    ops.add_task(app, "Design", duration="5d")
    t = ops.add_task(app, "Build", duration="10d", predecessors=[1], notes="n")
    assert t["id"] == 2
    assert t["manually_scheduled"] is False
    assert t["predecessors"] == "1"
    assert t["notes"] == "n"
    listed = ops.list_tasks(app, name_contains="bui")
    assert [x["name"] for x in listed["tasks"]] == ["Build"]
    assert listed["tasks"][0]["duration_days"] == 10.0


def test_add_task_under_parent_goes_after_existing_children(app):
    ops.add_task(app, "Phase 1")
    ops.add_task(app, "A", parent_id=1)
    ops.add_task(app, "Phase 2")
    ops.add_task(app, "B", parent_id=1)
    names = [(t["name"], t["outline_level"]) for t in ops.list_tasks(app)["tasks"]]
    assert names == [("Phase 1", 1), ("A", 2), ("B", 2), ("Phase 2", 1)]
    assert ops.get_task(app, 1)["summary"] is True
    assert ops.list_tasks(app, include_summary=False)["total"] == 3


def test_parent_and_before_conflict(app):
    ops.add_task(app, "X")
    with pytest.raises(ops.ProjectError):
        ops.add_task(app, "Y", parent_id=1, before_id=1)


def test_update_task_validation(app):
    ops.add_task(app, "X")
    t = ops.update_task(app, 1, percent_complete=50, constraint_type="must_start_on",
                        constraint_date="2026-11-01", name="Y")
    assert t["percent_complete"] == 50
    assert t["constraint_type"] == "must_start_on"
    assert t["constraint_date"] == "2026-11-01T00:00"
    assert t["name"] == "Y"
    with pytest.raises(ops.ProjectError):
        ops.update_task(app, 1, percent_complete=150)
    with pytest.raises(ops.ProjectError):
        ops.update_task(app, 1, start="next tuesday")
    with pytest.raises(ops.ProjectError, match="No task with ID 9"):
        ops.update_task(app, 9, name="Z")


def test_link_unlink_dependencies(app):
    ops.add_task(app, "A")
    ops.add_task(app, "B")
    ops.link_tasks(app, 1, 2, link_type="ss", lag="2d")
    deps = ops.get_task_dependencies(app, 2)["dependencies"]
    assert deps[0]["from_id"] == 1 and deps[0]["type"] == "SS"
    assert ops.unlink_tasks(app, 1, 2)["predecessors"] == ""
    with pytest.raises(ops.ProjectError):
        ops.link_tasks(app, 1, 1)
    with pytest.raises(ops.ProjectError):
        ops.link_tasks(app, 1, 2, link_type="XX")


def test_indent_outdent_delete(app):
    ops.add_task(app, "A")
    ops.add_task(app, "B")
    assert ops.indent_task(app, 2)["outline_level"] == 2
    assert ops.outdent_task(app, 2)["outline_level"] == 1
    res = ops.delete_task(app, 1)
    assert res["deleted"]["name"] == "A"
    assert ops.list_tasks(app)["tasks"][0]["id"] == 1


def test_resources_and_assignments(app):
    ops.add_task(app, "Build")
    r = ops.add_resource(app, "Alice", standard_rate="50/h", max_units=0.5)
    assert r["type"] == "work" and r["max_units"] == 0.5
    ops.add_resource(app, "Concrete", resource_type="material")
    a = ops.assign_resource(app, 1, "alice", units=0.5)
    assert a["resource_name"] == "Alice" and a["units"] == 0.5
    assert len(ops.list_assignments(app)["assignments"]) == 1
    assert len(ops.list_assignments(app, resource="Alice")["assignments"]) == 1
    ops.remove_assignment(app, 1, "1")
    assert ops.list_assignments(app, task_id=1)["assignments"] == []
    with pytest.raises(ops.ProjectError):
        ops.remove_assignment(app, 1, "Alice")
    assert ops.update_resource(app, "Alice", group="Dev")["group"] == "Dev"
    ops.delete_resource(app, "Concrete")
    assert [x["name"] for x in ops.list_resources(app)["resources"]] == ["Alice"]
    with pytest.raises(ops.ProjectError):
        ops.add_resource(app, "Bad", resource_type="robot")


def test_save_close_and_files(app, tmp_path):
    with pytest.raises(ops.ProjectError, match="never been saved"):
        ops.save_project(app)
    out = tmp_path / "plan.xml"
    ops.save_project(app, path=str(out), file_format="xml")
    assert app.calls[-1] == ("FileSaveAs", {"Name": str(out), "FormatID": "MSProject.XML"})
    ops.save_project(app)
    assert app.calls[-1][0] == "FileSave"
    with pytest.raises(ops.ProjectError):
        ops.save_project(app, path=str(out), file_format="docx")
    ops.export_pdf(app, str(tmp_path / "plan.pdf"))
    assert app.calls[-1][0] == "DocumentExport"
    ops.save_baseline(app)
    assert app.calls[-1] == ("BaselineSave", {"All": True})
    assert ops.close_project(app, save=True)["saved"] is True
    assert app.calls[-1] == ("FileCloseEx", {"Save": 1})
    assert ops.get_status(app)["open_projects"] == []


def test_open_project_missing_file(app, tmp_path):
    with pytest.raises(ops.ProjectError, match="File not found"):
        ops.open_project(app, str(tmp_path / "nope.mpp"))
    f = tmp_path / "real.mpp"
    f.write_bytes(b"")
    info = ops.open_project(app, str(f))
    assert info["name"] == "real.mpp"
    assert ops.activate_project(app, "Project1")["name"] == "Project1"


def test_critical_path(app):
    ops.add_task(app, "A")
    ops.add_task(app, "B")
    app.ActiveProject.Tasks(2).Critical = True
    assert [t["name"] for t in ops.get_critical_path(app)["critical_tasks"]] == ["B"]
