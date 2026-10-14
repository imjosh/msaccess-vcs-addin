"""Check actual SVG and gettext callers in owned disposable Access databases.

Requires Windows, Access, and pywin32. The installed add-in is only a library.
Use a rebuilt, closed development add-in in a disposable checkout for gettext.
"""
import argparse
import gc
import hashlib
import json
from pathlib import Path
import winreg
import xml.etree.ElementTree as ET

import pythoncom
import win32api
import win32com.client
import win32event
import win32process


def first(value):
    return value[0] if isinstance(value, tuple) else value


def artifact(path):
    data = path.read_bytes()
    return {"name": path.name, "bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
            "bom": data.startswith(b"\xef\xbb\xbf"), "text": data.decode("utf-8")}


def run(mode, database, library):
    database = database.resolve()
    library = library.resolve()
    if mode == "svg":
        assert not database.exists(), "Use a new disposable database"
    else:
        assert database.exists() and (database.parent / "Translation").is_dir()
        assert database != library, "The installed library must never be the host"
    result = {"mode": mode, "database": str(database), "library": str(library)}
    app = handle = translation = label = form = None
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
        handle = win32api.OpenProcess(0x1000 | 0x100000, False, pid)
        stamp = win32process.GetProcessTimes(handle)["CreationTime"]
        result.update(pid=pid, create_time=str(stamp))
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
            result["export"] = json.loads(api("ExportObject", "Forms", "frmUtf8Probe"))
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
            result["fixture_compile"] = api("CompileVBA", True)
            assert result["fixture_compile"] and app.IsCompiled
            result["translation_path"] = str(database.parent / "Translation")
            result["translation_probe"] = first(app.Run(str(database.with_suffix("")) + ".Utf8TranslationProbe", True))
            assert result["translation_probe"] == "OK", result["translation_probe"]
            files = sorted((database.parent / "Translation").glob("*.po"))
            files += sorted((database.parent / "Translation").glob("*.pot"))
            records = [artifact(path) for path in files]
            assert any(r["name"].endswith(".pot") for r in records)
            assert any(r["name"].endswith(".po") for r in records)
            for record in records:
                assert record["bytes"] and not record["bom"]
                assert 'msgid ""' in record["text"] and 'msgstr ""' in record["text"]
                assert "charset=UTF-8" in record["text"]
            result["artifacts"] = records
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
        if app is not None and handle is not None:
            try:
                window = app.hWndAccessApp
                window = window() if callable(window) else window
                assert win32process.GetWindowThreadProcessId(window)[1] == result["pid"]
                assert win32process.GetProcessTimes(handle)["CreationTime"] == stamp
                assert Path(app.CurrentProject.FullName).resolve() == database
                result["end_session_cleanup"] = first(app.Run(str(library.with_suffix("")) + ".API", "EndSession"))
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
            reopened = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Access.Application"))
            reopened.OpenCurrentDatabase(str(database))
            reopened.Visible = True
            window = reopened.hWndAccessApp
            window = window() if callable(window) else window
            reopen_pid = win32process.GetWindowThreadProcessId(window)[1]
            reopen_handle = win32api.OpenProcess(0x1000 | 0x100000, False, reopen_pid)
            assert Path(reopened.CurrentProject.FullName).resolve() == database
            project = next(p for p in reopened.VBE.VBProjects if Path(str(p.FileName)).resolve() == database)
            module = project.VBComponents("vcs_test_utf8_translation").CodeModule
            text = str(module.Lines(1, module.CountOfLines))
            result["reopened_fixture_readable"] = "Utf8TranslationProbe" in text
            result["reopened_is_compiled"] = bool(reopened.IsCompiled)
            result["reopened_pid"] = reopen_pid
            module = project = None
            reopened.Quit(2)
        except Exception as exc:
            result["reopen_error"] = repr(exc)
        finally:
            reopened = None
            gc.collect()
            if reopen_handle is not None:
                result["reopened_handle_exit"] = win32event.WaitForSingleObject(reopen_handle, 5000) == win32event.WAIT_OBJECT_0
                reopen_handle.Close()
            pythoncom.CoUninitialize()
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=["svg", "gettext"])
    parser.add_argument("database", type=Path)
    parser.add_argument("library", type=Path)
    parser.add_argument("receipt", type=Path)
    args = parser.parse_args()
    assert not args.receipt.exists(), "Use a new receipt path"
    result = run(args.mode, args.database, args.library)
    args.receipt.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in result.items() if k != "artifacts"}, indent=2))
    raise SystemExit(0 if result.get("success") and result.get("original_handle_exit") else 1)
