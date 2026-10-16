"""Actual native fixtures for complete artifact and grouped dependency policies."""
import hashlib
import json
from pathlib import Path
import re
import shutil
import uuid
import zipfile

import win32com.client

from request_policy_cases import POLICIES, require


FIXTURES = Path(__file__).with_name('Testing.accdb.src')


def digest_tree(folder):
    return {str(p.relative_to(folder)).lower(): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in folder.rglob('*') if p.is_file()}


def bitmap_pixels(data):
    """Decode the fixture's uncompressed RGB pixels independently of BMP headers."""
    require(data[:2] == b'BM', 'Picture proof must be a BMP')
    offset = int.from_bytes(data[10:14], 'little')
    width = int.from_bytes(data[18:22], 'little', signed=True)
    height = int.from_bytes(data[22:26], 'little', signed=True)
    bits = int.from_bytes(data[28:30], 'little')
    compression = int.from_bytes(data[30:34], 'little')
    require(width > 0 and height and bits == 24 and compression == 0,
            'Picture fixture proof requires an uncompressed 24-bit RGB BMP')
    stride = ((width * 3 + 3) // 4) * 4
    require(offset + stride * abs(height) <= len(data), 'Truncated picture fixture')
    rows = [data[offset + y * stride:offset + y * stride + width * 3]
            for y in range(abs(height))]
    if height > 0:
        rows.reverse()
    rgb = bytes(channel for row in rows for x in range(0, len(row), 3)
                for channel in (row[x + 2], row[x + 1], row[x]))
    return (width, abs(height)), rgb


def bitmap_evidence(path, reference=None):
    data = path.read_bytes()
    dimensions, pixels = bitmap_pixels(data)
    result = {'path': str(path), 'bytes': len(data),
              'sha256': hashlib.sha256(data).hexdigest(), 'dimensions': list(dimensions),
              'rgb_sha256': hashlib.sha256(pixels).hexdigest()}
    if reference is not None:
        original = reference.read_bytes()
        original_dimensions, original_pixels = bitmap_pixels(original)
        require(dimensions == original_dimensions, 'Picture normalization changed dimensions')
        result.update(changed_bytes=sum(a != b for a, b in zip(original, data))
                      + abs(len(original) - len(data)),
                      changed_rgb_pixels=sum(original_pixels[i:i + 3] != pixels[i:i + 3]
                                             for i in range(0, len(pixels), 3)))
    return result


def bitmap_variant(data, step):
    value = bytearray(data)
    require(value[:2] == b'BM', 'Native fixture must be a BMP')
    offset = int.from_bytes(value[10:14], 'little')
    require(offset < len(value), 'Invalid BMP pixel offset')
    value[offset] ^= step
    return bytes(value)


def menu_fixture(cases, name):
    menu = cases.source / 'menus' / (name + '.json')
    images = menu.parent / (name + '_Images')
    images.mkdir(parents=True)
    data = json.loads((FIXTURES / 'menus/Demo Popup.json').read_text(encoding='utf-8-sig'))
    data['Items']['NameLocal'] = name
    data['Items']['Controls'][0]['ImagePath'] = 'rel:' + name + '_Images\\Custom Button'
    data['Items']['Controls'][0]['OnAction'] = ''
    menu.write_bytes(b'\xef\xbb\xbf' + json.dumps(data, indent=2).encode())
    for source in (FIXTURES / 'menus/Demo Popup_Images').iterdir():
        shutil.copy2(source, images / source.name)
    return menu, images


def set_native_command_bar_picture(cases, name, picture, source_picture, item):
    helper = Path(__file__).with_name('Fixtures') / 'modules/vcs_policy_artifact_fixture.bas'
    import_folder = cases.folder / (cases.host.stem + '-picture-fixture-import')
    import_folder.mkdir(exist_ok=False)
    import_copy = import_folder / helper.name
    # VBE.Import expects BOM-free system text, as the product module importer does.
    # Keep tracked UTF-8 fixture bytes and CRLF intact; convert only the owned copy.
    import_copy.write_bytes(helper.read_bytes().decode('utf-8-sig').encode('mbcs'))
    item['native_picture_helper'] = {
        'expected_module': 'vcs_policy_artifact_fixture', 'expected_type': 1,
        'canonical_fixture': str(helper.resolve()),
        'fixture_sha256': hashlib.sha256(helper.read_bytes()).hexdigest(),
        'transport_copy': str(import_copy), 'transport_encoding': 'system mbcs without BOM',
        'transport_sha256': hashlib.sha256(import_copy.read_bytes()).hexdigest(),
        'host': str(cases.host)}
    cases.save()
    project = next(p for p in cases.app.VBE.VBProjects
                   if Path(p.FileName).resolve() == cases.host.resolve())
    module = project.VBComponents.Import(str(import_copy.resolve()))
    item['native_picture_helper'].update(module=module.Name, vbcomponent_type=int(module.Type),
                                         project=str(project.FileName))
    cases.save()
    require(module.Name == 'vcs_policy_artifact_fixture', f'Unexpected native picture fixture module: {item["native_picture_helper"]}')
    require(int(module.Type) == 1, 'Native picture fixture must be a real standard module')
    cases.app.VBE.ActiveVBProject = project
    cases.app.DoCmd.Save(win32com.client.constants.acModule, module.Name)
    item['native_picture_helper']['saved_vbcomponent_type'] = int(project.VBComponents(module.Name).Type)
    cases.save()
    require(item['native_picture_helper']['saved_vbcomponent_type'] == 1, 'Saving changed the native picture fixture identity')
    module = project = None
    # Independently assign each raw variant to its own sacrificial copied button.
    # Office transforms pixels during button assignment, beyond LoadPicture/SavePicture.
    # The tested button is untouched until both independent proofs exist.
    canonical = {}
    for variant, raw in [('database', picture), ('source', source_picture)]:
        normalized = cases.folder / (cases.host.stem + '-canonical-' + variant + '.bmp')
        probe_name = name + ' independent ' + variant
        # Import assigns the source Picture and then its transparency Mask.
        # Database mutation assigns Picture only and keeps the existing mask.
        mask = source_picture.with_name(source_picture.name.replace('_Picture.bmp', '_Mask.bmp'))
        require(mask.is_file(), 'Source picture fixture must include its import-defining mask')
        mask_argument = str(mask) if variant == 'source' else ''
        result = cases.app.Run('F1CanonicalizeButtonPicture', str(cases.host), name,
                               probe_name, str(raw), str(normalized), mask_argument)
        result = result[0] if isinstance(result, tuple) else result
        diagnostic = cases.app.Run('F1PictureFixtureLastError')
        diagnostic = diagnostic[0] if isinstance(diagnostic, tuple) else diagnostic
        details = {'fixture_result': result, 'fixture_error': diagnostic,
                   'method': ('independent copied button Picture then source Mask then SavePicture'
                              if variant == 'source' else
                              'independent copied button Picture assignment then SavePicture'),
                   'mask_input': mask_argument,
                   'mask_sha256': hashlib.sha256(mask.read_bytes()).hexdigest(),
                   'temporary_probe_bar': probe_name, 'probe_deleted_before_success': result is True,
                   'raw': bitmap_evidence(raw), 'canonical_path': str(normalized)}
        item.setdefault('native_picture_canonicalization', {})[variant] = details
        cases.save()
        require(result is True and not diagnostic, f'Native picture normalization failed: {details}')
        details['canonical'] = bitmap_evidence(normalized, raw)
        canonical[variant] = normalized
        cases.save()
    item['independent_picture_variants'] = bitmap_evidence(canonical['source'], canonical['database'])
    cases.save()
    require(item['independent_picture_variants']['changed_rgb_pixels'] > 0,
            'Requested picture variants became pixel-identical during native button assignment')
    proof = cases.folder / (cases.host.stem + '-assigned-picture.bmp')
    result = cases.app.Run('F1SetCommandBarPicture', str(cases.host), name, str(picture), str(proof))
    result = result[0] if isinstance(result, tuple) else result
    diagnostic = cases.app.Run('F1PictureFixtureLastError')
    diagnostic = diagnostic[0] if isinstance(diagnostic, tuple) else diagnostic
    item['native_picture_assignment'] = {'fixture_result': result, 'fixture_error': diagnostic,
                                          'proof': str(proof)}
    cases.save()
    require(result is True and not diagnostic, f'Native picture fixture failed: {item["native_picture_assignment"]}')
    expected = hashlib.sha256(canonical['database'].read_bytes()).hexdigest()
    actual = hashlib.sha256(proof.read_bytes()).hexdigest()
    item['native_picture_assignment'].update(expected_sha256=expected, actual_sha256=actual,
                                           actual=bitmap_evidence(proof, canonical['database']))
    cases.save()
    require(actual == expected, 'Access-local fixture did not assign the normalized requested actual picture')
    require(digest_tree(source_picture.parent) == item['source_image_directory_before_fixture'],
            'Native fixture normalization changed canonical source images')
    return canonical


def command_bars(cases, transports):
    for transport in transports:
        for direction in ['import', 'export']:
            for policy in POLICIES:
                key = '-'.join(['artifact-menu', transport, direction, policy])
                item = cases.new(key)
                name = 'F1 Menu ' + uuid.uuid4().hex[:8]
                menu, images = menu_fixture(cases, name)
                imported = cases.call('external', 'ImportObject', 'command_bar', name, 'block')
                item['seed_import'] = imported
                cases.save()
                cases.outcome(imported, transferred=True)
                seeded = cases.call('external', 'ExportObject', 'command_bar', name, 'block')
                item['seed_export'] = seeded
                cases.save()
                cases.outcome(seeded, transferred=True)
                picture = next(images.glob('*_Picture.bmp'))
                baseline = picture.read_bytes()
                db_picture = cases.folder / (key + '-database.bmp')
                db_picture.write_bytes(bitmap_variant(baseline, 0x40))
                picture.write_bytes(bitmap_variant(baseline, 0x80))
                source_before = digest_tree(menu.parent)
                # Load/assign IPictureDisp inside Access; out-of-process assignment
                # cannot reliably marshal Office Picture properties.
                item['source_image_directory_before_fixture'] = digest_tree(images)
                canonical = set_native_command_bar_picture(cases, name, db_picture, picture, item)
                history_before = cases.history()
                result = cases.call(transport, direction.title() + 'Object', 'command_bar', name, policy)
                item.update(seed_import=imported, seed_export=seeded, result=result,
                            source_before=source_before, source_after=digest_tree(menu.parent))
                cases.save()
                winning = policy == ('prefer_source' if direction == 'import' else 'prefer_database')
                blocked = policy == 'block'
                cases.outcome(result, blocked=blocked, skipped=not blocked and not winning, transferred=winning)
                if not winning or direction == 'import':
                    require(item['source_after'] == source_before, 'Menu comparison changed canonical image content')
                if not winning:
                    require(cases.history() == history_before, 'Menu preservation advanced history')
                # A deliberate database export makes the actual native picture observable.
                observed = cases.call('external', 'ExportObject', 'command_bar', name, 'prefer_database')
                cases.outcome(observed, transferred=True)
                expected_variant = 'source' if direction == 'import' and winning else 'database'
                expected = canonical[expected_variant].read_bytes()
                item['observed_picture'] = bitmap_evidence(picture, canonical[expected_variant])
                item['expected_picture_variant'] = expected_variant
                item['observed_picture_sha256'] = hashlib.sha256(picture.read_bytes()).hexdigest()
                item['expected_picture_sha256'] = hashlib.sha256(expected).hexdigest()
                cases.save()
                require(item['observed_picture_sha256'] == item['expected_picture_sha256'], 'Native menu contains the wrong picture')
                cases.reopen()
                item['reopened_history'] = cases.history()
                cases.save()


def theme_xml_variant(data, color):
    text = data.decode('utf-8-sig')
    changed, count = re.subn(r'(?<=val=")[0-9A-Fa-f]{6}(?=")', color, text, count=1)
    require(count == 1, 'Theme fixture has no editable RGB color')
    return changed.encode('utf-8')


def replace_theme_attachment(cases, name, path):
    db = cases.app.CurrentDb()
    rows = db.OpenRecordset("SELECT [Data] FROM MSysResources WHERE [Type]='thmx' AND [Name]='" + name.replace("'", "''") + "'")
    require(not rows.EOF, 'Imported theme resource was not found')
    rows.Edit()
    attachment = rows.Fields('Data').Value
    if not attachment.EOF:
        attachment.Delete()
    attachment.AddNew()
    attachment.Fields('FileData').LoadFromFile(str(path))
    attachment.Update()
    attachment.Close()
    rows.Update()
    rows.Close()
    attachment = rows = db = None


def read_theme_attachment(cases, name, path):
    db = cases.app.CurrentDb()
    rows = db.OpenRecordset("SELECT [Data] FROM MSysResources WHERE [Type]='thmx' AND [Name]='" + name.replace("'", "''") + "'")
    attachment = rows.Fields('Data').Value
    attachment.Fields('FileData').SaveToFile(str(path))
    attachment.Close()
    rows.Close()
    attachment = rows = db = None



def theme_resource_snapshot(cases, name):
    """Retain read-only DAO identities even when a seed API reply reports failure."""
    result = {'requested_name': name, 'rows': [], 'matching_name_extension_type_count': 0}
    db = rows = None
    try:
        db = cases.app.CurrentDb()
        rows = db.OpenRecordset('SELECT [Name], [Extension], [Type] FROM MSysResources')
        while not rows.EOF:
            row = {key: rows.Fields(key).Value for key in ['Name', 'Extension', 'Type']}
            result['rows'].append(row)
            if (str(row['Name']).lower() == name.lower() and
                    str(row['Extension']).lower() == 'thmx' and str(row['Type']).lower() == 'thmx'):
                result['matching_name_extension_type_count'] += 1
            rows.MoveNext()
        result['row_count'] = len(result['rows'])
    except Exception as error:
        # A diagnostic query cannot replace the original API failure. The caller
        # persists this observed error and still requires a successful snapshot.
        result['query_error'] = repr(error)
    finally:
        if rows is not None:
            try:
                rows.Close()
            except Exception as error:
                result['close_error'] = repr(error)
        rows = db = None
    return result


def themes(cases, transports):
    for transport in transports:
        for direction in ['import', 'export']:
            for policy in POLICIES:
                key = '-'.join(['artifact-theme', transport, direction, policy])
                item = cases.new(key)
                name = 'F1 Theme ' + uuid.uuid4().hex[:8]
                folder = cases.source / 'themes' / name
                folder.mkdir(parents=True)
                fixture = FIXTURES / 'themes/Angles.thmx'
                with zipfile.ZipFile(fixture) as archive:
                    for member in archive.namelist():
                        require((folder / member).resolve().is_relative_to(folder.resolve()), 'Unsafe fixture archive path')
                    archive.extractall(folder)
                options_file = cases.source / 'vcs-options.json'
                options = json.loads(options_file.read_text(encoding='utf-8-sig'))
                options['Options']['ExtractThemeFiles'] = True
                options_file.write_bytes(b'\xef\xbb\xbf' + json.dumps(options).encode())
                configured = cases.api('SetOption', 'ExtractThemeFiles', True)
                require(configured.get('success') is True, 'Theme extraction setup failed')
                item['seed_source_before'] = digest_tree(folder)
                imported = cases.call('external', 'ImportObject', 'theme', name, 'block')
                item['seed_import'] = imported
                cases.save()
                item['resources_after_seed_import'] = theme_resource_snapshot(cases, name)
                cases.save()
                cases.outcome(imported, transferred=True)
                resource_snapshot = item['resources_after_seed_import']
                require(not resource_snapshot.get('query_error') and not resource_snapshot.get('close_error'),
                        f'Theme resource identity query failed: {resource_snapshot}')
                require(resource_snapshot['matching_name_extension_type_count'] == 1,
                        f'First extracted-theme import did not create exactly one matching resource: {resource_snapshot}')
                item['seed_source_after'] = digest_tree(folder)
                item['seed_import_history'] = cases.history()
                seed_attachment = cases.folder / (key + '-seed-observed.thmx')
                read_theme_attachment(cases, name, seed_attachment)
                with zipfile.ZipFile(seed_attachment) as archive:
                    item['seed_attachment_files'] = {
                        member.filename.replace('\\', '/').lower(): hashlib.sha256(archive.read(member)).hexdigest()
                        for member in archive.infolist() if not member.is_dir()}
                cases.save()
                require(item['seed_source_after'] == item['seed_source_before'],
                        'First extracted-theme import changed canonical source files')
                require(not folder.with_suffix('.thmx').exists(),
                        'First extracted-theme import wrote a canonical upload archive')
                expected_attachment = {key.replace('\\', '/'): value for key, value in item['seed_source_before'].items()}
                require(item['seed_attachment_files'] == expected_attachment,
                        'First extracted-theme import did not preserve the full actual attachment content')
                entries = [entry for key, entry in item['seed_import_history'].get('components', {}).get('Themes', {}).items()
                           if key.lower() == (folder.name + '.thmx').lower()]
                require(len(entries) == 1 and entries[0].get('SourceFilesHash') and entries[0].get('DatabaseFilesHash'),
                        'First extracted-theme import did not persist both complete-content baselines')
                seeded = cases.call('external', 'ExportObject', 'theme', name, 'block')
                item['seed_export'] = seeded
                cases.save()
                cases.outcome(seeded, transferred=True)
                xml = next(p for p in folder.rglob('*.xml') if 'clrScheme' in p.read_text(encoding='utf-8-sig'))
                original = xml.read_bytes()
                source_xml = theme_xml_variant(original, '112233')
                database_xml = theme_xml_variant(original, '445566')
                xml.write_bytes(source_xml)
                db_archive = cases.folder / (key + '-database.thmx')
                with zipfile.ZipFile(db_archive, 'w', zipfile.ZIP_DEFLATED) as archive:
                    for path in folder.rglob('*'):
                        if path.is_file():
                            archive.writestr(path.relative_to(folder).as_posix(), database_xml if path == xml else path.read_bytes())
                replace_theme_attachment(cases, name, db_archive)
                source_before = digest_tree(folder)
                history_before = cases.history()
                result = cases.call(transport, direction.title() + 'Object', 'theme', name, policy)
                item.update(seed_import=imported, seed_export=seeded, result=result,
                            source_before=source_before, source_after=digest_tree(folder))
                cases.save()
                winning = policy == ('prefer_source' if direction == 'import' else 'prefer_database')
                blocked = policy == 'block'
                cases.outcome(result, blocked=blocked, skipped=not blocked and not winning, transferred=winning)
                if not winning or direction == 'import':
                    require(item['source_after'] == source_before, 'Theme comparison/import changed canonical source')
                require(not folder.with_suffix('.thmx').exists(), 'Extracted theme import wrote a canonical upload archive')
                if not winning:
                    require(cases.history() == history_before, 'Theme preservation advanced history')
                observed = cases.folder / (key + '-observed.thmx')
                read_theme_attachment(cases, name, observed)
                with zipfile.ZipFile(observed) as archive:
                    actual = archive.read(xml.relative_to(folder).as_posix())
                expected = source_xml if direction == 'import' and winning else database_xml
                require(actual == expected, 'Native theme resource has the wrong XML content')
                item['observed_xml_sha256'] = hashlib.sha256(actual).hexdigest()
                cases.reopen()
                item['reopened_history'] = cases.history()
                cases.save()


def access_text(path):
    data = path.read_bytes()
    return data.decode('utf-16') if data.startswith((b'\xff\xfe', b'\xfe\xff')) else data.decode('utf-8-sig')


def grouped_table_macros(cases):
    template = (FIXTURES / 'tdmacros/tblSaveXML.xml').read_text(encoding='utf-8-sig')
    action = re.search(r'<Action Name="SetLocalVar">.*?</Action>', template, flags=re.S).group()
    macro = '<?xml version="1.0"?>\r\n<DataMacros xmlns="http://schemas.microsoft.com/office/accessservices/2010/12/application"><DataMacro Event="AfterUpdate"><Statements>' + action + '</Statements></DataMacro></DataMacros>'
    for direction in ['import', 'export']:
        for policy in POLICIES:
            item = cases.new('grouped-table-macro-' + direction + '-' + policy)
            macro_source = cases.source / 'tdmacros/tblPolicy.xml'
            macro_source.parent.mkdir(parents=True)
            macro_source.write_text(macro, encoding='utf-8-sig')
            imported = cases.call('external', 'ImportObject', 'table_data_macro', 'tblPolicy', 'block')
            cases.outcome(imported, transferred=True)
            seeded = cases.call('external', 'ExportObject', 'table', 'tblPolicy', 'prefer_database')
            require(seeded.get('transferred_count') == 3, f'Table/data/macro group missing a member: {seeded}')
            baseline = macro_source.read_text(encoding='utf-8-sig')
            source_macro = baseline.replace('[ID]', '3')
            database_macro = baseline.replace('[ID]', '2')
            require(source_macro != baseline and database_macro != baseline, 'Macro fixture mutation failed')
            macro_source.write_text(source_macro, encoding='utf-8-sig')
            database_path = cases.folder / ('database-macro-' + uuid.uuid4().hex + '.xml')
            database_path.write_text(database_macro, encoding='utf-8-sig')
            cases.app.LoadFromText(win32com.client.constants.acTableDataMacro, 'tblPolicy', str(database_path))
            cases.set_source('table_data', 99)
            cases.set_db('table_data', 2)
            before = digest_tree(cases.source)
            history = cases.history()
            result = cases.call('external', direction.title() + 'Object', 'table', 'tblPolicy', policy)
            item.update(seed_import=imported, seed_export=seeded, result=result)
            cases.save()
            winning = policy == ('prefer_source' if direction == 'import' else 'prefer_database')
            blocked = policy == 'block'
            cases.outcome(result, blocked=blocked, skipped=not blocked and not winning, transferred=not blocked)
            require(result.get('transferred_count') == (0 if blocked else 3 if winning else 1), f'Grouped macro transfer count mismatch: {result}')
            require(result.get('skipped_count') == (2 if not blocked and not winning else 0), f'Grouped macro preservation count mismatch: {result}')
            if blocked:
                # Logs/overrides may be written, so compare the actual transfer groups.
                after = digest_tree(cases.source)
                for path, value in before.items():
                    if path.startswith(('tbldefs', 'tdmacros', 'tables')):
                        require(after.get(path) == value, 'Blocked macro dependency partially transferred')
                require(cases.history() == history, 'Blocked macro dependency advanced history')
            observed = cases.folder / ('observed-macro-' + uuid.uuid4().hex + '.xml')
            cases.app.SaveAsText(win32com.client.constants.acTableDataMacro, 'tblPolicy', str(observed))
            actual = access_text(observed)
            require(('>3<' if direction == 'import' and winning else '>2<') in actual, 'Native table macro has wrong content')
            require(cases.db_value('table_data') == str(99 if direction == 'import' and winning else 2), 'Grouped table data has wrong value')
            cases.reopen()
            item['reopened_history'] = cases.history()
            cases.save()


def run_artifacts(cases, transports, artifact_types=None):
    selected = artifact_types if artifact_types is not None else ['command_bar', 'theme', 'table_data_macro']
    if 'command_bar' in selected:
        command_bars(cases, transports)
    if 'theme' in selected:
        themes(cases, transports)
    if 'table_data_macro' in selected:
        grouped_table_macros(cases)
