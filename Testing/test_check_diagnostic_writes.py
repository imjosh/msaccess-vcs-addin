"""Controlled supervisor and safety tests. No Access application is launched."""
import copy
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch, Mock

spec = importlib.util.spec_from_file_location('driver', Path(__file__).with_name('check_diagnostic_writes.py'))
driver = importlib.util.module_from_spec(spec)
spec.loader.exec_module(driver)


class DriverTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.out = Path(self.temp.name)

    def valid_record(self):
        names = ['AppendsLines', 'FailsQuietly', 'Locked', 'WithLiveErrorWrites', 'WithLiveErrorFailsQuietly']
        return {'mode': 'suite', 'filters': [], 'qualified': True, 'initial_compile': True,
                'fresh_reopen_compiled': True, 'components_preserved': True, 'after_pids': [],
                'library_preserved': True, 'references_preserved': True,
                'initial_process': {'original_handle_exit': True}, 'reopen_process': {'original_handle_exit': True},
                'tests': {'allPassed': True, 'summary': {'passed': 5, 'failed': 0, 'errored': 0},
                          'tests': {'modTestFileAccess.TestTryAppendToFile'+n:
                                    {'status': 'PASSED', 'assertions': [{'passed': True}]} for n in names}}}

    def test_invalid_labels_and_paths(self):
        for label in ('../bad', 'a/b', 'a\\b', '..', '.', 'NUL', 'COM1', 'bad.', 'bad:ads', ''):
            with self.subTest(label=label), self.assertRaises(RuntimeError):
                driver.validate_inputs(self.out, label)
        for out in (driver.REPO, driver.REPO/'evidence', driver.REPO.parent):
            with self.subTest(out=out), self.assertRaises(RuntimeError):
                driver.validate_inputs(out, 'safe')

    def test_competing_or_unknown_inventory(self):
        for inventory in ({123}, None):
            with patch.object(driver, 'list_access_pids_or_none', return_value=inventory), self.assertRaises(RuntimeError):
                driver.run(self.out, 'busy', 'suite', [], 'sync')
        self.assertFalse((self.out/'busy.json').exists())

    def test_qualification_prerequisites(self):
        valid = self.valid_record()
        self.assertTrue(all(driver.qualification_checks(valid).values()))
        for key in ('initial_compile', 'fresh_reopen_compiled', 'components_preserved'):
            record = copy.deepcopy(valid)
            record[key] = False
            self.assertFalse(all(driver.qualification_checks(record).values()))
        for key in ('library_preserved', 'references_preserved'):
            self.assertFalse(all(driver.qualification_checks({**valid, key: False}).values()))
        for name in ('initial_process', 'reopen_process'):
            record = copy.deepcopy(valid)
            record[name]['original_handle_exit'] = False
            self.assertFalse(all(driver.qualification_checks(record).values()))
        for key, value in (('after_pids', None), ('after_pids', [123]), ('error', 'failed'), ('close_error', 'blocked')):
            record = {**valid, key: value}
            self.assertFalse(all(driver.qualification_checks(record).values()))

    def test_missing_regression_and_failed_suite(self):
        record = self.valid_record()
        self.assertTrue(all(driver.behavior_checks(record).values()))
        record['tests']['tests'].pop(next(iter(record['tests']['tests'])))
        self.assertFalse(all(driver.behavior_checks(record).values()))
        record = self.valid_record()
        record['tests']['summary']['failed'] = 1
        self.assertFalse(all(driver.behavior_checks(record).values()))
        record = self.valid_record()
        next(iter(record['tests']['tests'].values()))['assertions'].append({'passed': True, 'message': 'SKIP: unavailable'})
        self.assertFalse(all(driver.behavior_checks(record).values()))

    def test_stalled_and_failed_workers_retain_evidence(self):
        for label, code, expected_timeout in [('stalled', 'import time; time.sleep(60)', True),
                                               ('failed', 'raise SystemExit(7)', False)]:
            host = self.out/'hosts'/label/'repo'
            host.mkdir(parents=True)
            (host/'preserved.txt').write_text('evidence')
            driver.save_record(self.out, label, {'active_phase': 'controlled.blocked_call', 'qualified': False})
            start = time.monotonic()
            with self.assertRaises(RuntimeError):
                driver.supervise([sys.executable, '-c', code], self.out, label, .25, 5)
            record = json.loads((self.out/(label+'.json')).read_bytes())
            self.assertLess(time.monotonic()-start, 9)
            self.assertFalse(record['qualified'])
            self.assertEqual(record['timed_out'], expected_timeout)
            self.assertEqual(record['failed_phase'], 'controlled.blocked_call')
            self.assertEqual(record['recovery_worker']['exit_code'], 0)
            self.assertTrue(record['retained_failed_host'])
            self.assertTrue((host/'preserved.txt').is_file())
            if os.environ.get('DIAGNOSTIC_TEST_EVIDENCE'):
                evidence = Path(os.environ['DIAGNOSTIC_TEST_EVIDENCE'])
                evidence.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(self.out/(label+'.json'), evidence/(label+'.json'))

    def test_stalled_recovery_is_bounded(self):
        driver.save_record(self.out, 'recovery', {'qualified': False, 'active_phase': 'shutdown'})
        real_popen = subprocess.Popen
        calls = 0
        def launch(command, **kwargs):
            nonlocal calls
            calls += 1
            if calls == 2:
                command = [sys.executable, '-c', 'import time; time.sleep(60)']
            return real_popen(command, **kwargs)
        start = time.monotonic()
        with patch.object(driver.subprocess, 'Popen', side_effect=launch), self.assertRaises(RuntimeError):
            driver.supervise([sys.executable, '-c', 'import time; time.sleep(60)'], self.out, 'recovery', .25, .25)
        self.assertLess(time.monotonic()-start, 4)
        record = json.loads((self.out/'recovery.json').read_bytes())
        self.assertIsNone(record['recovery_worker']['exit_code'])
        self.assertFalse(record['qualified'])

    def test_disposal_failure_reports_actual_retention(self):
        folder = self.out/'hosts'/'partial'
        folder.mkdir(parents=True)
        self.assertTrue(folder.resolve().is_relative_to(self.out.resolve()))
        driver.save_record(self.out, 'partial', {'qualified': True, 'active_phase': 'remove_worktree', 'archive_sha256': 'saved'})
        (self.out/'partial-host.zip').write_bytes(b'closed archive')
        command = [sys.executable, '-c', 'import shutil,sys; shutil.rmtree(sys.argv[1]); sys.exit(7)', str(folder)]
        with self.assertRaises(RuntimeError):
            driver.supervise(command, self.out, 'partial', 5, 5)
        record = json.loads((self.out/'partial.json').read_bytes())
        self.assertFalse(record['qualified'])
        self.assertFalse(record['retained_failed_host'])
        self.assertTrue(record['archived_host_available'])

    def test_child_optimization_matches_supervisor(self):
        command = driver.python_command('-c', 'import sys; print(sys.flags.optimize)')
        result = subprocess.check_output(command, text=True)
        self.assertEqual(int(result), sys.flags.optimize)

    def test_recovery_refuses_reused_pid_and_wrong_image(self):
        record = {'initial_process': {'pid': 123, 'create_time': 1}}
        with patch.object(driver, 'process_create_time', return_value=2), patch.object(driver.win32api, 'OpenProcess') as opening:
            result = driver.recover_processes(record, 1)
            self.assertEqual(result[0]['outcome'], 'recovery_refused_or_failed')
            opening.assert_not_called()
        with patch.object(driver, 'process_create_time', return_value=1), \
                patch.object(driver.win32api, 'OpenProcess', return_value=Mock()), \
                patch.object(driver, 'process_image', return_value='C:/python.exe'), \
                patch.object(driver.win32api, 'TerminateProcess') as terminate:
            self.assertEqual(driver.recover_processes(record, 1)[0]['outcome'], 'recovery_refused_or_failed')
            terminate.assert_not_called()

    def test_cleanup_requires_qualification_archive_and_containment(self):
        label = 'cleanup'
        folder = self.out/'hosts'/label/'repo'
        folder.mkdir(parents=True)
        record = self.valid_record()
        record['host'] = str(folder/'Version Control.accda')
        archive = self.out/(label+'-host.zip')
        archive.write_bytes(b'archive')
        record['archive_sha256'] = driver.sha(archive)
        with patch.object(driver, 'list_access_pids_or_none', return_value=set()):
            self.assertEqual(driver.cleanup_target(self.out, label, record), (folder.parent, folder))
            with patch.object(Path, 'is_junction', side_effect=lambda: True), self.assertRaises(RuntimeError):
                driver.cleanup_target(self.out, label, record)
            for key, value in [('qualified', False), ('timed_out', True), ('initial_compile', False),
                               ('archive_sha256', 'bad'), ('host', str(self.out/'outside.accda'))]:
                with self.subTest(key=key), self.assertRaises(RuntimeError):
                    driver.cleanup_target(self.out, label, {**record, key: value})
            driver.save_record(self.out, label, record)
            with patch.object(driver.subprocess, 'run') as git:
                driver.cleanup(self.out, label)
                git.assert_called_once()
            self.assertFalse(folder.parent.exists())
            self.assertTrue(json.loads((self.out/(label+'.json')).read_bytes())['disposed'])


if __name__ == '__main__':
    unittest.main()
