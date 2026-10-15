"""Portable shutdown regression checks. Run with python -m unittest discover -s Testing.

Compile the driver's nested shutdown functions in a test closure so these checks
exercise the production code without importing Windows COM or creating a host.
Native lifecycle controls separately cover archiving and disposal.
"""
import ast
from pathlib import Path
from types import SimpleNamespace
import time
import unittest
from unittest.mock import Mock


DRIVER = Path(__file__).with_name('check_disposable_suite.py')


class ShutdownTests(unittest.TestCase):
    def setUp(self):
        self.info = {'pid': 123, 'create_time': 456}
        self.app = Mock()
        self.app.CurrentProject.FullName = str(DRIVER)
        self.handle = Mock()
        self.pid = Mock(return_value=123)
        self.stamp = Mock(return_value=456)
        self.wait = Mock(return_value=0)
        self.pids = Mock(return_value=set())
        self.scope = {
            'Path': Path, 'gc': Mock(), 'time': time, 'EXIT_WAIT_MS': 60000,
            'pid_from_access_app': self.pid, 'process_create_time': self.stamp,
            'list_access_pids_or_none': self.pids,
            'win32event': SimpleNamespace(WaitForSingleObject=self.wait,
                                         WAIT_OBJECT_0=0, WAIT_TIMEOUT=258),
        }

    def shutdown(self, handle=True):
        tree = ast.parse(DRIVER.read_text(encoding='utf-8-sig'))
        run = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'run')
        names = {'identity_mismatch', 'wait_for_no_access', 'shutdown_owned'}
        functions = [n for n in run.body if isinstance(n, ast.FunctionDef) and n.name in names]
        self.assertEqual(len(functions), len(names))
        factory = ast.parse('''
def make_shutdown(app, handle, info):
    data = {'phase': info}
    project = module = current_phase = None
''').body[0]
        factory.body.extend(functions)
        factory.body.append(ast.parse('return shutdown_owned').body[0])
        module = ast.fix_missing_locations(ast.Module(body=[factory], type_ignores=[]))
        exec(compile(module, str(DRIVER), 'exec'), self.scope)
        self.scope['make_shutdown'](self.app, self.handle if handle else None, self.info)('phase')

    def assert_quit_and_exit(self):
        self.app.Quit.assert_called_once_with(2)
        self.wait.assert_called_once_with(self.handle, 60000)
        self.handle.Close.assert_called_once_with()
        self.assertEqual(self.info['exit_wait_ms'], 60000)
        self.assertTrue(self.info['original_handle_exit'])
        self.assertTrue(self.info['closure_confirmed'])
        self.assertNotIn('identity_mismatch', self.info)

    def test_matching_identity(self):
        self.shutdown()
        self.assert_quit_and_exit()

    def test_unavailable_pid(self):
        self.pid.return_value = None
        self.shutdown()
        self.assert_quit_and_exit()
        self.assertTrue(self.info['pid_check_unavailable'])

    def test_unavailable_creation_time(self):
        self.stamp.return_value = None
        self.shutdown()
        self.assert_quit_and_exit()
        self.assertTrue(self.info['create_time_check_unavailable'])

    def test_lookup_exceptions(self):
        self.pid.side_effect = RuntimeError('PID lookup failed')
        self.stamp.side_effect = RuntimeError('creation-time lookup failed')
        self.shutdown()
        self.assert_quit_and_exit()
        self.assertIn('PID lookup failed', self.info['pid_check_error'])
        self.assertIn('creation-time lookup failed', self.info['create_time_check_error'])

    def test_contradictory_identity(self):
        for field in ['pid', 'create_time']:
            with self.subTest(field=field):
                self.setUp()
                lookup = self.pid if field == 'pid' else self.stamp
                lookup.return_value = 999
                self.shutdown()
                self.app.Quit.assert_not_called()
                self.wait.assert_called_once_with(self.handle, 0)
                self.assertEqual(self.info['identity_mismatch'], field)

    def test_unavailable_pid_does_not_hide_creation_time_mismatch(self):
        self.pid.return_value = None
        self.stamp.return_value = 999
        self.shutdown()
        self.app.Quit.assert_not_called()
        self.assertEqual(self.info['identity_mismatch'], 'create_time')

    def test_incomplete_startup_still_quits(self):
        self.info.clear()
        self.shutdown(handle=False)
        self.app.Quit.assert_called_once_with(2)
        self.assertEqual(self.info['exit_wait_ms'], 60000)
        self.assertTrue(self.info['process_list_exit'])
        self.assertTrue(self.info['closure_confirmed'])

    def test_unavailable_identity_without_exit_evidence(self):
        self.stamp.return_value = None
        self.wait.return_value = 258
        self.pids.return_value = {123}
        self.shutdown()
        self.app.Quit.assert_called_once_with(2)
        self.wait.assert_called_once_with(self.handle, 60000)
        self.assertTrue(self.info['exit_timeout'])
        self.assertFalse(self.info['closure_confirmed'])


if __name__ == '__main__':
    unittest.main()
