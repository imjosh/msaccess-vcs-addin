"""F1 native cases for check_policy_workflow.py. All setup runs outside assertions."""
import asyncio
import inspect
import json
from pathlib import Path
import struct
import time
import uuid

import pythoncom

POLICIES = ['block', 'prefer_source', 'prefer_database', 'skip', 'decline']


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def index_components(path):
    """Inspect supported binary history without modifying serialized indexes."""
    if not path.exists():
        return {'version': None, 'components': {}}
    data = path.read_bytes()
    require(data[:4] == b'VCSI', 'Unexpected index magic')
    version = int.from_bytes(data[4:6], 'little')
    require(version in [3, 4, 5], f'Unexpected index version {version}')
    pos = 6

    def number():
        nonlocal pos
        value = struct.unpack_from('<I', data, pos)[0]
        pos += 4
        return value

    def text():
        nonlocal pos
        length = number()
        value = data[pos:pos + length].decode('utf-16-le')
        pos += length
        return value

    pos += 32
    text()
    pos += 8
    for _ in range(number()):
        text()
        text()
    components = {}
    for _ in range(number()):
        category = text()
        entries = {}
        for _ in range(number()):
            name = text()
            flags = data[pos]
            pos += 1
            item = {'dates': list(struct.unpack_from('<dddd', data, pos))}
            pos += 32
            for flag, key in [(1, 'FileHash'), (2, 'OtherHash'), (4, 'MetaHash'),
                              (8, 'FilePropertiesHash'), (16, 'FolderAnnotation'),
                              (32, 'AllFilesHash'), (64, 'DatabaseFilesHash'), (128, 'SourceFilesHash')]:
                if flags & flag:
                    item[key] = text()
            entries[name] = item
        components[category] = entries
    return {'version': version, 'components': components}


class Cases:
    def __init__(self, app, driver, folder, endpoint, record, save, legacy=False):
        self.app, self.driver, self.folder, self.endpoint = app, driver, folder, endpoint
        self.record, self.save, self.legacy = record, save, legacy
        self.host = self.source = None

    def api(self, method, *args):
        value = self.app.Run(self.driver + '.API', method, *args)
        value = value[0] if isinstance(value, tuple) else value
        return json.loads(value) if isinstance(value, str) and value.startswith('{') else value

    def new(self, key):
        initial_host = self.host is None
        if not initial_host:
            self.app.CloseCurrentDatabase()
        self.host = self.folder / (key + '.accdb')
        self.source = Path(str(self.host) + '.src')
        self.app.NewCurrentDatabase(str(self.host))
        if initial_host:
            self.app.Visible = True
            self.app.UserControl = True
        require(Path(self.app.CurrentProject.FullName).resolve() == self.host.resolve(), 'Wrong disposable host')
        db = self.app.CurrentDb()
        db.CreateQueryDef('qryPolicy', 'SELECT 1 AS FixtureValue;')
        db.CreateQueryDef('qryUnrelated', 'SELECT 17 AS FixtureValue;')
        db.Properties.Append(db.CreateProperty('F1PolicyValue', 10, 'baseline'))
        db.Execute('CREATE TABLE tblPolicy (ID LONG CONSTRAINT PrimaryKey PRIMARY KEY, FixtureValue LONG)')
        db.Execute('INSERT INTO tblPolicy (ID, FixtureValue) VALUES (1, 1)')
        db = None
        # Access persists first-open NavPane defaults on close. Stabilize that real
        # import-defining metadata before seeding history, rather than ignoring it.
        self.app.CloseCurrentDatabase()
        self.app.OpenCurrentDatabase(str(self.host))
        require(Path(self.app.CurrentProject.FullName).resolve() == self.host.resolve(), 'Wrong initialized fixture host')
        db = self.app.CurrentDb()
        rows = db.OpenRecordset('SELECT COUNT(*) FROM tblPolicy')
        require(rows.Fields(0).Value == 1, 'Initialized fixture DAO table is unreadable')
        rows.Close()
        rows = db = None
        self.source.mkdir(exist_ok=True)
        options = {'Info': {}, 'Options': {'ExportFolder': '', 'ExportFormatVersion': '5.1.0',
                   'TablesToExportData': {'tblPolicy': {'Format': 'Tab Delimited'}}}}
        (self.source / 'vcs-options.json').write_bytes(b'\xef\xbb\xbf' + json.dumps(options).encode())
        self.app.Run(self.driver + '.Preload')
        registered = self.api('RegisterSession', 'f1-' + uuid.uuid4().hex)
        require(registered.get('success') is True, f'Session setup failed: {registered}')
        configured = self.api('SetOption', 'ExportFormatVersion', 50100)
        require(configured.get('success') is True, f'Format setup failed: {configured}')
        item = {'host': str(self.host), 'fixture_first_open_defaults_stabilized_before_history': True}
        self.record['cases'][key] = item
        return item

    def reopen(self):
        self.app.CloseCurrentDatabase()
        self.app.OpenCurrentDatabase(str(self.host))
        require(Path(self.app.CurrentProject.FullName).resolve() == self.host.resolve(), 'Wrong reopened disposable host')
        db = self.app.CurrentDb()
        rows = db.OpenRecordset('SELECT COUNT(*) AS FixtureRows FROM tblPolicy')
        count = rows.Fields(0).Value
        require(count == 1, f'Reopened disposable DAO fixture is unreadable: {count}')
        rows.Close()
        rows = db = None
        self.record['cases'][self.host.stem]['reopened_host'] = {'path': str(self.host), 'dao_fixture_rows': count}
        self.save()

    def path(self, kind='query'):
        return self.source / {'query': 'queries/qryPolicy.sql', 'db_property': 'dbs-properties.json',
                              'table_data': 'tables/tblPolicy.txt'}[kind]

    def name(self, kind):
        return 'tblPolicy' if kind == 'table_data' else 'qryPolicy'

    def set_source(self, kind, value):
        if kind == 'query':
            data = f'SELECT {value} AS FixtureValue;\r\n'
        elif kind == 'table_data':
            data = f'ID\tFixtureValue\r\n1\t{value}\r\n'
        else:
            data = json.loads(self.path(kind).read_text(encoding='utf-8-sig'))
            data['Items']['F1PolicyValue']['Value'] = str(value)
            data = json.dumps(data, indent=2) + '\r\n'
        self.path(kind).write_bytes(b'\xef\xbb\xbf' + data.encode())

    def set_db(self, kind, value):
        db = self.app.CurrentDb()
        if kind == 'query':
            db.QueryDefs('qryPolicy').SQL = f'SELECT {value} AS FixtureValue;'
        elif kind == 'table_data':
            db.Execute(f'UPDATE tblPolicy SET FixtureValue={value} WHERE ID=1')
        else:
            db.Properties('F1PolicyValue').Value = str(value)
        db = None

    def db_value(self, kind='query', name='qryPolicy'):
        db = self.app.CurrentDb()
        if kind in ['query', 'table_data']:
            rows = db.OpenRecordset('SELECT FixtureValue FROM tblPolicy WHERE ID=1' if kind == 'table_data' else name)
            value = rows.Fields(0).Value
            rows.Close()
            rows = None
        else:
            value = db.Properties('F1PolicyValue').Value
        db = None
        return str(value)

    def history(self):
        return index_components(self.source / 'vcs-index.idx')

    async def mcp(self, method, kind, name, policy):
        from msaccess_vcs_mcp import tools
        from msaccess_vcs_mcp.operation_manager import OperationManager
        OperationManager.get_instance().set_event_loop(asyncio.get_running_loop())
        tool = tools.vcs_import_object if method == 'ImportObject' else tools.vcs_export_object
        value = tool(str(self.host), object_type=kind, object_name=name,
                     noninteractive=True, decision_policy=policy)
        return await value if inspect.isawaitable(value) else value

    def call(self, transport, method, kind, name, policy):
        if transport == 'mcp':
            return asyncio.run(self.mcp(method, kind, name, policy))
        args = (kind, name) if self.legacy else (kind, name, policy)
        if transport == 'external':
            return self.api(method, *args)
        start = len(self.endpoint.callbacks)
        identity = uuid.uuid4().hex
        admitted = self.app.Run(self.driver + '.APIAsync', self.endpoint.info(identity), method, *args)
        admitted = json.loads(admitted[0] if isinstance(admitted, tuple) else admitted)
        sync_result = admitted.get('result') if admitted.get('sync') else None
        if isinstance(sync_result, str):
            sync_result = json.loads(sync_result)
        self.last_dispatch = {'admission': admitted, 'sync_result': sync_result}
        if not (admitted.get('sync') or admitted.get('async') or admitted.get('started')):
            return admitted
        end = time.monotonic() + 60
        while time.monotonic() < end:
            pythoncom.PumpWaitingMessages()
            terminal = [c for c in self.endpoint.callbacks[start:] if c['type'] in ['complete', 'error', 'cancelled']]
            if terminal:
                break
            time.sleep(.03)
        require(len(terminal) == 1 and terminal[0]['operation_id'] == identity, f'Wrong terminal callback: {terminal}')
        if isinstance(sync_result, dict):
            for key in ['success', 'transferred', 'skipped', 'transferred_count', 'skipped_count', 'decisions', 'error_pattern']:
                if key in sync_result:
                    require(terminal[0].get(key) == sync_result[key], f'Synchronous/callback field differs: {key}')
        return terminal[0]

    def seed(self, kind):
        result = self.call('external', 'ExportObject', 'query', 'qryUnrelated', 'prefer_database')
        require(result.get('success') is True, f'Unrelated seed failed: {result}')
        result = self.call('external', 'ExportObject', kind, self.name(kind), 'prefer_database')
        require(result.get('success') is True, f'Seed failed: {result}')
        return result

    def outcome(self, result, blocked=False, skipped=False, transferred=False):
        require(bool(result.get('decision_required')) == blocked, f'Decision mismatch: {result}')
        require(result.get('success') is (not blocked), f'Success mismatch: {result}')
        if not self.legacy:
            require(result.get('transferred') is transferred and result.get('skipped') is skipped, f'Transfer mismatch: {result}')
        if blocked or skipped:
            require(bool(result.get('decisions')), f'Journal missing: {result}')

    def matrix(self, transports):
        for transport in ['external'] if self.legacy else transports:
            for kind in ['query', 'db_property', 'table_data']:
                for direction in ['import', 'export']:
                    for policy in ['legacy'] if self.legacy else POLICIES:
                        key = '-'.join([transport, kind, direction, policy])
                        item = self.new(key)
                        item['seed'] = self.seed(kind)
                        self.set_source(kind, 99)
                        self.set_db(kind, 2)
                        original = self.path(kind).read_bytes()
                        before = self.history()
                        result = self.call(transport, direction.title() + 'Object', kind, self.name(kind), policy)
                        item['result'] = result
                        if transport == 'async':
                            item['dispatch'] = self.last_dispatch
                        self.save()
                        transferred = self.legacy or policy == ('prefer_source' if direction == 'import' else 'prefer_database')
                        blocked = not self.legacy and policy == 'block'
                        self.outcome(result, blocked, not blocked and not transferred, transferred)
                        if not self.legacy:
                            require(result.get('transferred_count') == (1 if transferred else 0), f'Single-object completed count mismatch: {result}')
                            require(result.get('skipped_count') == (1 if not blocked and not transferred else 0), f'Single-object preserved count mismatch: {result}')
                            require(result.get('decision_policy') == policy, f'Request policy was not retained: {result}')
                            if transport == 'mcp':
                                require(result.get('initiating_source') == 'mcp', f'MCP source identity was not retained: {result}')
                        require(self.db_value(kind) == str(99 if transferred and direction == 'import' else 2), f'Database changed incorrectly: {key}')
                        changed = self.path(kind).read_bytes() != original
                        require(changed == (transferred and direction == 'export'), f'Source changed incorrectly: {key}')
                        item['history_before'], item['history_after'] = before, self.history()
                        if not transferred:
                            require(item['history_after'] == before, f'Skipped/blocked history advanced: {key}')
                        require(item['history_after']['components'].get('Queries', {}).get('qryUnrelated.sql') == before['components'].get('Queries', {}).get('qryUnrelated.sql'), f'Unrelated history changed: {key}')
                        self.reopen()
                        item['reopened_value'], item['reopened_history'] = self.db_value(kind), self.history()
                        require(item['reopened_history'] == item['history_after'], 'Persisted history differs')
                        if not self.legacy and not transferred:
                            repeat = self.call(transport, direction.title() + 'Object', kind, self.name(kind), 'block')
                            item['reopened_conflict'] = repeat
                            self.outcome(repeat, blocked=True)
                        self.save()

    def cycles_and_migration(self):
        for kind in ['query', 'db_property', 'table_data']:
            item = self.new(kind + '-repeated-cycles')
            item['seed'], item['cycles'] = self.seed(kind), []
            for value in [7, 8, 9]:
                self.set_source(kind, value)
                result = self.call('external', 'ImportObject', kind, self.name(kind), 'block')
                self.outcome(result, transferred=True)
                require(self.db_value(kind) == str(value), 'Ordinary edit was not imported')
                item['cycles'].append(result)
                self.reopen()
            self.set_db(kind, 10)
            result = self.call('external', 'ExportObject', kind, self.name(kind), 'block')
            self.outcome(result, transferred=True)
            item['one_sided_export'] = result
            self.reopen()
            self.set_db(kind, 11)
            self.set_source(kind, 12)
            result = self.call('external', 'ImportObject', kind, self.name(kind), 'block')
            self.outcome(result, blocked=True)
            item['new_conflict_after_reopen'] = result
            self.save()
            for identical in [False, True]:
                item = self.new(kind + ('-identical-missing-history' if identical else '-divergent-missing-history'))
                item['seed'] = self.seed(kind)
                if not identical:
                    self.set_source(kind, 99)
                    self.set_db(kind, 2)
                index = self.source / 'vcs-index.idx'
                item['removed_index_sha256'] = __import__('hashlib').sha256(index.read_bytes()).hexdigest()
                index.rename(self.source / 'historical-disabled-index.idx')
                result = self.call('external', 'ImportObject', kind, self.name(kind), 'block')
                self.outcome(result, blocked=not identical, transferred=identical)
                item['result'] = result
                self.save()

    def deletion_matrix(self):
        for policy in POLICIES:
            item = self.new('import-edited-object-deletion-' + policy)
            item['seed'] = self.seed('query')
            self.set_db('query', 2)
            self.path().unlink()
            result = self.api('ImportByType', 'queries', True, policy)
            item['result'] = result
            require(self.db_value() == '2', 'Edited database object was deleted')
            require(bool(result.get('decision_required')) == (policy in ['block', 'prefer_source']), f'Import deletion mismatch: {result}')
            require(result.get('transferred_count') == (0 if policy in ['block', 'prefer_source'] else 1), f'Import batch completed count mismatch: {result}')
            require(result.get('skipped_count') == (0 if policy in ['block', 'prefer_source'] else 1), f'Import batch preservation count mismatch: {result}')
            self.save()
            item = self.new('export-changed-tracked-orphan-' + policy)
            item['seed'] = self.seed('query')
            self.set_source('query', 99)
            original = self.path().read_bytes()
            self.app.CurrentDb().QueryDefs.Delete('qryPolicy')
            result = self.api('ExportByType', 'queries', True, policy)
            item['result'] = result
            require(self.path().read_bytes() == original, 'Changed tracked orphan was deleted')
            require(bool(result.get('decision_required')) == (policy in ['block', 'prefer_database']), f'Export deletion mismatch: {result}')
            require(result.get('transferred_count') == (0 if policy in ['block', 'prefer_database'] else 1), f'Export batch completed count mismatch: {result}')
            require(result.get('skipped_count') == (0 if policy in ['block', 'prefer_database'] else 1), f'Export batch preservation count mismatch: {result}')
            self.save()
            item = self.new('export-untracked-source-' + policy)
            item['seed'] = self.seed('query')
            added = self.source / 'queries/qryTeammate.sql'
            original = b'\xef\xbb\xbfSELECT 99 AS FixtureValue;\r\n'
            added.write_bytes(original)
            result = self.api('ExportByType', 'queries', True, policy)
            item['result'] = result
            require(result.get('success') is True and not result.get('decision_required'), f'Untracked file blocked export: {result}')
            require(added.read_bytes() == original, 'Untracked file changed')
            require(result.get('transferred_count') == 2 and result.get('skipped_count') == 1, f'Untracked-source batch counts mismatch: {result}')
            self.save()
        for tracked in [False, True]:
            for policy in POLICIES:
                key = ('tracked' if tracked else 'untracked') + '-orphan-metadata-without-primary-' + policy
                item = self.new(key)
                item['seed'] = self.seed('query')
                metadata = self.path().with_suffix('.json')
                require(metadata.exists(), 'Query seed did not create its metadata companion')
                if tracked:
                    self.app.CurrentDb().QueryDefs.Delete('qryPolicy')
                    self.path().unlink()
                    data = json.loads(metadata.read_text(encoding='utf-8-sig'))
                    data['Items']['Description'] = 'Edited orphan metadata fixture'
                    metadata.write_bytes(b'\xef\xbb\xbf' + json.dumps(data, indent=2).encode())
                else:
                    added = metadata.with_name('qryTeammate.json')
                    added.write_bytes(metadata.read_bytes())
                    metadata = added
                original = metadata.read_bytes()
                before = self.history()
                result = self.api('ExportByType', 'queries', True, policy)
                item.update(result=result, metadata=str(metadata), history_before=before, history_after=self.history())
                blocked = tracked and policy in ['block', 'prefer_database']
                require(bool(result.get('decision_required')) == blocked, f'Orphan metadata policy mismatch: {result}')
                require(metadata.read_bytes() == original, 'Policy deleted or rewrote an orphan metadata companion')
                require(result.get('transferred_count') == (0 if blocked else 1 if tracked else 2), f'Orphan metadata transfer count mismatch: {result}')
                require(result.get('skipped_count') == (0 if blocked else 1), f'Orphan metadata preservation count mismatch: {result}')
                decisions = result.get('decisions', [])
                require(any(d.get('classification') == ('tracked_changed_orphan' if tracked else 'untracked_orphan') and d.get('tracked') is tracked for d in decisions), f'Orphan metadata journal classification missing: {result}')
                if blocked:
                    require(self.history() == before, 'Blocked orphan metadata advanced history')
                self.save()
        for direction in ['import', 'export']:
            item = self.new(direction + '-mixed-overwrite-and-blocked-deletion')
            item['seed'] = self.seed('query')
            self.set_source('query', 99)
            self.set_db('query', 2)
            db = self.app.CurrentDb()
            other = self.source / 'queries/qryUnrelated.sql'
            if direction == 'import':
                db.QueryDefs('qryUnrelated').SQL = 'SELECT 18 AS FixtureValue;'
                other.unlink()
                result = self.api('ImportByType', 'queries', True, 'prefer_source')
                require(self.db_value() == '2' and self.db_value(name='qryUnrelated') == '18', 'Mixed import partially applied resolutions')
            else:
                other.write_bytes(b'\xef\xbb\xbfSELECT 18 AS FixtureValue;\r\n')
                db.QueryDefs.Delete('qryUnrelated')
                original = self.path().read_bytes()
                result = self.api('ExportByType', 'queries', True, 'prefer_database')
                require(self.path().read_bytes() == original and other.exists(), 'Mixed export partially applied resolutions')
            db = None
            item['result'] = result
            require(result.get('decision_required') is True, f'Mixed deletion did not block: {result}')
            self.save()
        item = self.new('new-object-and-source')
        result = self.call('external', 'ExportObject', 'query', 'qryPolicy', 'block')
        self.outcome(result, transferred=True)
        item['new_source'] = result
        added = self.source / 'queries/qryNew.sql'
        added.write_bytes(b'\xef\xbb\xbfSELECT 55 AS FixtureValue;\r\n')
        result = self.call('external', 'ImportObject', 'query', 'qryNew', 'block')
        self.outcome(result, transferred=True)
        require(self.db_value(name='qryNew') == '55', 'New object import failed')
        item['new_object'] = result
        self.save()
        for direction in ['import', 'export']:
            for policy in POLICIES:
                item = self.new('grouped-table-' + direction + '-' + policy)
                seed = self.call('external', 'ExportObject', 'table', 'tblPolicy', 'prefer_database')
                self.outcome(seed, transferred=True)
                require(seed.get('transferred_count') == 2, f'Table definition/data group did not transfer both members: {seed}')
                item['seed'] = seed
                self.set_source('table_data', 99)
                self.set_db('table_data', 2)
                original_data = self.path('table_data').read_bytes()
                schema = self.source / 'tbldefs/tblPolicy.xml'
                original_schema = schema.read_bytes()
                before = self.history()
                result = self.call('external', direction.title() + 'Object', 'table', 'tblPolicy', policy)
                item['result'] = result
                winning = policy == ('prefer_source' if direction == 'import' else 'prefer_database')
                blocked = policy == 'block'
                self.outcome(result, blocked=blocked, skipped=not blocked and not winning, transferred=not blocked)
                require(result.get('transferred_count') == (0 if blocked else 2 if winning else 1), f'Grouped actual transfer count mismatch: {result}')
                require(result.get('skipped_count') == (1 if not blocked and not winning else 0), f'Grouped preservation count mismatch: {result}')
                require(self.db_value('table_data') == str(99 if direction == 'import' and winning else 2), 'Grouped policy changed the wrong database content')
                if blocked:
                    require(schema.read_bytes() == original_schema and self.path('table_data').read_bytes() == original_data, 'Blocked dependent conflict partially transferred the group')
                    require(self.history() == before, 'Blocked group advanced history')
                self.reopen()
                item['reopened_history'] = self.history()
                self.save()
