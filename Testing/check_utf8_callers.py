"""Check actual SVG and gettext callers in owned disposable Access databases.

Requires Windows, Access, and pywin32. The installed add-in is only a library.
Use a rebuilt, closed development add-in in a disposable checkout for gettext.
"""
if not __debug__:
    raise SystemExit("UTF-8 caller checks require assertions. Run without -O, -OO, or PYTHONOPTIMIZE.")

import argparse
import gc
import hashlib
import json
from pathlib import Path
import subprocess
import xml.etree.ElementTree as ET


# The prepared development fixture contains these two translation languages.
EXPECTED_PO_FILES = {"en_TEST.po", "pt_BR.po"}


def validate_preparation(mode, database, library, disposable_host_marker=None):
    if mode == "svg":
        assert not database.exists(), "Use a new disposable database"
        assert not Path(str(database) + ".src").exists(), "Use a path with no stale source folder"
        return []
    assert database.is_file() and (database.parent / "Translation").is_dir()
    assert database != library, "The installed library must never be the host"
    assert disposable_host_marker is not None, "Gettext requires --disposable-host-marker for this host"
    marker = json.loads(disposable_host_marker.read_text(encoding="utf-8"))
    assert marker.get("purpose") == "utf8-gettext-caller", "Incorrect disposable host marker purpose"
    designated = Path(marker.get("database", ""))
    assert designated.is_absolute() and designated.resolve() == database, "Disposable marker names a different host"

    def git(*args):
        return subprocess.run(["git", "-C", str(database.parent), *args],
                              check=True, capture_output=True).stdout

    root = Path(git("rev-parse", "--show-toplevel").decode("utf-8").strip()).resolve()
    assert root == database.parent, "Gettext host must be in the disposable checkout root"
    folder = database.parent / "Translation"
    assert folder.resolve() == folder, "Translation folder must be inside the disposable checkout"
    stale = sorted(p for pattern in ("*.po", "*.pot") for p in folder.glob(pattern))
    assert all(p.is_file() and not p.is_symlink() for p in stale), "Translation inputs must be regular files"
    tracked = {name.decode("utf-8") for name in
               git("ls-tree", "-r", "--name-only", "-z", "HEAD", "--", "Translation").split(b"\0")
               if name and Path(name.decode("utf-8")).parent == Path("Translation")
               and Path(name.decode("utf-8")).suffix.lower() in (".po", ".pot")}
    actual = {p.relative_to(root).as_posix() for p in stale}
    assert actual == tracked, "Translation inputs include missing, untracked, or ignored POT/PO files"
    assert EXPECTED_PO_FILES <= {p.name for p in stale}, "Prepare both en_TEST.po and pt_BR.po translation inputs"
    # Diff against HEAD includes staged and unstaged edits, with Git's normal
    # working-tree encoding and line-ending conversions.
    try:
        git("diff", "--exit-code", "HEAD", "--", *sorted(tracked))
        git("diff", "--cached", "--exit-code", "HEAD", "--", *sorted(tracked))
    except subprocess.CalledProcessError as exc:
        raise AssertionError("Translation POT/PO inputs have staged or unstaged edits; use a clean disposable checkout") from exc
    # Git can hide working-tree edits behind assume-unchanged/skip-worktree.
    # Hash actual input through Git's normal conversions without writing objects.
    for name in sorted(tracked):
        assert git("hash-object", "--path=" + name, name).strip() == git("rev-parse", "HEAD:" + name).strip(), \
            "Translation POT/PO input differs from HEAD: " + name
    return stale


def validate_gettext_artifacts(records):
    names = {record["name"] for record in records}
    assert EXPECTED_PO_FILES <= names, "Missing expected PO outputs: " + ", ".join(sorted(EXPECTED_PO_FILES - names))
    assert "MSAccessVCS.pot" in names, "Missing expected MSAccessVCS.pot output"
    for record in records:
        assert record["bytes"] and not record["bom"]
        assert 'msgid ""' in record["text"] and 'msgstr ""' in record["text"]


def first(value):
    return value[0] if isinstance(value, tuple) else value


def artifact(path):
    data = path.read_bytes()
    return {"name": path.name, "bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
            "bom": data.startswith(b"\xef\xbb\xbf"), "text": data.decode("utf-8")}


def run(mode, database, library, disposable_host_marker=None):
    database = database.resolve()
    library = library.resolve()
    stale = validate_preparation(mode, database, library, disposable_host_marker)
    # Keep Windows/COM imports after preparation so refusals cannot start Access.
    import winreg
    import pythoncom
    import win32api
    import win32com.client
    import win32event
    import win32process

    result = {"mode": mode, "database": str(database), "library": str(library)}
    app = handle = stamp = translation = label = form = None
    saved_contribute = None
    setting = r"Software\VB and VBA Program Settings\MSAccessVCS\Language"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, setting) as key:
            saved_contribute = winreg.QueryValueEx(key, "Contribute To Translations")
    except FileNotFoundError:
        pass
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("Access.Application")
        app = win32com.client.gencache.EnsureDispatch(app)
        window = app.hWndAccessApp
        window = window() if callable(window) else window
        pid = win32process.GetWindowThreadProcessId(window)[1]
        result["pid"] = pid
        handle = win32api.OpenProcess(0x1000 | 0x100000, False, pid)
        stamp = win32process.GetProcessTimes(handle)["CreationTime"]
        result["create_time"] = str(stamp)
        if mode == "svg":
            database.parent.mkdir(parents=True, exist_ok=True)
            app.NewCurrentDatabase(str(database))
        else:
            app.OpenCurrentDatabase(str(database))
        app.Visible = True
        assert Path(app.CurrentProject.FullName).resolve() == database
        lib = str(library.with_suffix(""))

        def api(method, *args):
            return first(app.Run(lib + ".API", method, *args))

        app.Run(lib + ".SetInteractionMode", 1)
        result["version"] = api("GetVCSVersion")
        if mode == "svg":
            constants = win32com.client.constants
            form = app.CreateForm()
            original_name = form.Name
            label = app.CreateControl(original_name, constants.acLabel)
            label.Properties.Item("Caption").Value = "UTF-8 omega " + chr(0x3A9)
            app.DoCmd.Save(constants.acForm, original_name)
            app.DoCmd.Close(constants.acForm, original_name, constants.acSaveYes)
            label = form = None
            app.DoCmd.Rename("frmUtf8Probe", constants.acForm, original_name)
            result["option"] = json.loads(api("SetOption", "ExportLayoutSvg", True))
            assert result["option"].get("success") is True, result["option"]
            result["export"] = json.loads(api("ExportObject", "Forms", "frmUtf8Probe"))
            assert result["export"].get("success") is True, result["export"]
            files = list(Path(str(database) + ".src").rglob("*.svg"))
            result["svg_files"] = [str(path) for path in files]
            assert len(files) == 1, "Expected one actual form SVG"
            record = artifact(files[0])
            ET.fromstring(record["text"])
            assert record["bytes"] and not record["bom"] and chr(0x3A9) in record["text"]
            result["artifacts"] = [record]
        else:
            # The fixture uses the development project's clsTranslation directly.
            # Import only into this idle disposable current project, never into
            # a loaded library. Keep the fixture through close and reopen.
            fixture = Path(__file__).parent / "Fixtures/modules/vcs_test_utf8_translation.bas"
            project = next(p for p in app.VBE.VBProjects if Path(str(p.FileName)).resolve() == database)
            assert not any(c.Name == "vcs_test_utf8_translation" for c in project.VBComponents)
            component = project.VBComponents.Import(str(fixture.resolve()))
            result["fixture_name"] = component.Name
            assert component.Name == "vcs_test_utf8_translation"
            component = project = None
            app.DoCmd.Save(win32com.client.constants.acModule, "vcs_test_utf8_translation")
            result["fixture_compile"] = api("CompileVBA", True)
            assert result["fixture_compile"] and app.IsCompiled
            result["translation_path"] = str(database.parent / "Translation")
            # Remove committed POT/PO files so only this run's callers can recreate them.
            # Recheck immediately before deletion as well as before opening Access.
            stale = validate_preparation(mode, database, library, disposable_host_marker)
            result["deleted_before_probe"] = [p.name for p in stale]
            for path in stale:
                path.unlink()
            # Path discovery requires the master POT to exist. This placeholder
            # has no gettext headers, so the checks below require a fresh export.
            template = database.parent / "Translation" / "MSAccessVCS.pot"
            template.write_text("# UTF-8 caller probe: replace this placeholder.\n", encoding="utf-8")
            result["translation_probe"] = first(app.Run(str(database.with_suffix("")) + ".Utf8TranslationProbe", True))
            assert result["translation_probe"] == "OK", result["translation_probe"]
            files = sorted((database.parent / "Translation").glob("*.po"))
            files += sorted((database.parent / "Translation").glob("*.pot"))
            result["artifacts"] = [artifact(path) for path in files]
            validate_gettext_artifacts(result["artifacts"])
        result["end_session"] = api("EndSession")
        result["success"] = True
    except Exception as exc:
        result["error"] = repr(exc)
        result["success"] = False
    finally:
        translation = label = form = None
        if mode == "gettext":
            try:
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, setting) as key:
                    if saved_contribute is None:
                        try:
                            winreg.DeleteValue(key, "Contribute To Translations")
                        except FileNotFoundError:
                            pass
                    else:
                        winreg.SetValueEx(key, "Contribute To Translations", 0,
                                         saved_contribute[1], saved_contribute[0])
                result["contribute_setting_restored"] = True
            except Exception as exc:
                result["restoration_error"] = repr(exc)
        if app is not None:
            # app is the instance this script created with DispatchEx. Quit it even
            # when its process handle was never opened, unless a recorded identity
            # check shows that it is no longer that process.
            owned = True
            try:
                if handle is not None:
                    window = app.hWndAccessApp
                    window = window() if callable(window) else window
                    owned = win32process.GetWindowThreadProcessId(window)[1] == result["pid"]
                    if owned and stamp is not None:
                        owned = win32process.GetProcessTimes(handle)["CreationTime"] == stamp
                    assert owned, "Owned Access process identity changed"
                assert Path(app.CurrentProject.FullName).resolve() == database
                result["end_session_cleanup"] = first(app.Run(str(library.with_suffix("")) + ".API", "EndSession"))
                if mode == "gettext" and result.get("fixture_name") == "vcs_test_utf8_translation":
                    project = next(p for p in app.VBE.VBProjects if Path(str(p.FileName)).resolve() == database)
                    app.VBE.ActiveVBProject = project
                    app.DoCmd.RunCommand(win32com.client.constants.acCmdCompileAndSaveAllModules)
                    project = None
                    result["compile_and_save_before_close"] = bool(app.IsCompiled)
            except Exception as exc:
                result["close_error"] = repr(exc)
            if owned:
                try:
                    app.Quit(2)
                    result["graceful_quit"] = True
                except Exception as exc:
                    result["close_error"] = repr(exc)
        app = None
        gc.collect()
        if handle is not None:
            result["original_handle_exit"] = win32event.WaitForSingleObject(handle, 5000) == win32event.WAIT_OBJECT_0
            handle.Close()
        pythoncom.CoUninitialize()
    if mode == "gettext" and result.get("fixture_name") == "vcs_test_utf8_translation" and result.get("original_handle_exit"):
        pythoncom.CoInitialize()
        reopened = reopen_handle = None
        try:
            reopened = win32com.client.DispatchEx("Access.Application")
            reopened = win32com.client.gencache.EnsureDispatch(reopened)
            reopened.OpenCurrentDatabase(str(database))
            reopened.Visible = True
            window = reopened.hWndAccessApp
            window = window() if callable(window) else window
            reopen_pid = win32process.GetWindowThreadProcessId(window)[1]
            result["reopened_pid"] = reopen_pid
            reopen_handle = win32api.OpenProcess(0x1000 | 0x100000, False, reopen_pid)
            assert Path(reopened.CurrentProject.FullName).resolve() == database
            project = next(p for p in reopened.VBE.VBProjects if Path(str(p.FileName)).resolve() == database)
            module = project.VBComponents("vcs_test_utf8_translation").CodeModule
            text = str(module.Lines(1, module.CountOfLines))
            result["reopened_fixture_readable"] = "Utf8TranslationProbe" in text
            result["reopened_compiled_before_check"] = bool(reopened.IsCompiled)
            reopened.VBE.ActiveVBProject = project
            reopened.DoCmd.RunCommand(win32com.client.constants.acCmdCompileAndSaveAllModules)
            result["reopened_is_compiled"] = bool(reopened.IsCompiled)
            result["reopened_pid"] = reopen_pid
            module = project = None
            reopened.Quit(2)
            reopened = None
        except Exception as exc:
            result["reopen_error"] = repr(exc)
        finally:
            if reopened is not None:
                try:
                    if reopen_handle is not None:
                        window = reopened.hWndAccessApp
                        window = window() if callable(window) else window
                        assert win32process.GetWindowThreadProcessId(window)[1] == result["reopened_pid"]
                    reopened.Quit(2)
                except Exception as exc:
                    result["reopen_close_error"] = repr(exc)
            reopened = None
            gc.collect()
            if reopen_handle is not None:
                result["reopened_handle_exit"] = win32event.WaitForSingleObject(reopen_handle, 5000) == win32event.WAIT_OBJECT_0
                reopen_handle.Close()
            pythoncom.CoUninitialize()
    if mode == "gettext":
        result["success"] = bool(result.get("success") and result.get("contribute_setting_restored")
                                 and result.get("reopened_fixture_readable")
                                 and result.get("reopened_is_compiled")
                                 and result.get("reopened_handle_exit"))
    result["success"] = bool(result.get("success") and result.get("original_handle_exit")
                             and "close_error" not in result and "restoration_error" not in result)
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=["svg", "gettext"])
    parser.add_argument("database", type=Path)
    parser.add_argument("library", type=Path)
    parser.add_argument("receipt", type=Path)
    parser.add_argument("--disposable-host-marker", type=Path,
                        help="JSON designation with purpose utf8-gettext-caller and the absolute database path")
    args = parser.parse_args()
    assert not args.receipt.exists(), "Use a new receipt path"
    result = run(args.mode, args.database, args.library, args.disposable_host_marker)
    args.receipt.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in result.items() if k != "artifacts"}, indent=2))
    raise SystemExit(0 if result.get("success") and result.get("original_handle_exit") else 1)
