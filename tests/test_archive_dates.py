"""Dated archive imports use synthetic fixtures in temporary HR_DATA only.

Run these tests inside Docker. Importing the fixture module, rather than its
TestCase class, keeps unittest discovery from running the older suite twice.
"""
import copy
import hashlib
import inspect
import io
import json
import shutil
import unittest
from datetime import date, datetime, timezone
from pathlib import Path
from unittest.mock import patch

try:
    import test_archive_batch as archive_fixtures
except ModuleNotFoundError:
    from tests import test_archive_batch as archive_fixtures

from app import batch, cli, common, service


class ArchiveDateTest(unittest.TestCase):
    setUp = archive_fixtures.ArchiveBatchTest.setUp
    tearDown = archive_fixtures.ArchiveBatchTest.tearDown
    write_archive = archive_fixtures.ArchiveBatchTest.write_archive
    add_candidate = archive_fixtures.ArchiveBatchTest.add_candidate

    def import_on(self, as_of=None, registry=None, output=None):
        result = batch.import_archive(self.archive, self.template, output,
                                      as_of=as_of, partner_registry=registry)
        return common.active(), result

    def registry_package(self, tag, as_of):
        directory = self.root / ('registry-' + tag)
        sources = directory / 'sources'
        sources.mkdir(parents=True)
        excerpt = 'Официальный партнер: ООО «Синтетический партнер»'
        snapshot = sources / (tag + '.txt')
        snapshot.write_text(excerpt, encoding='utf-8')
        raw = sources / (tag + '.html')
        raw.write_text('<h2>' + excerpt + '</h2>', encoding='utf-8')
        index = sources / (tag + '-index.html')
        index.write_text('<p>Каталог: ' + tag + '</p>', encoding='utf-8')
        registry = {
            'schema_version': 1,
            'amendment': {
                'id': 'synthetic-' + tag, 'rule_row': 262, 'months': 6,
                'as_of': as_of, 'approved_at': as_of,
                'approved_by': 'synthetic reviewer',
                'user_instruction': 'Все официальные партнеры ОЭЗ',
                'approved_domains': ['official.example'],
            },
            'entities': [{
                'entity_id': 'synthetic-partner',
                'canonical_name': 'ООО «Синтетический партнер»',
                'aliases': ['ООО «Синтетический партнер»'],
                'relation': 'official_partner', 'certainty': 'probable',
                'temporal_basis': 'Проверен синтетический снимок на дату запуска.',
                'evidence': [{
                    'url': 'https://official.example/partners',
                    'title': 'Синтетический официальный каталог',
                    'excerpt': excerpt,
                    'snapshot': snapshot.relative_to(directory).as_posix(),
                    'sha256': common.sha(snapshot),
                    'raw_snapshot': raw.relative_to(directory).as_posix(),
                    'raw_sha256': common.sha(raw),
                    'checked_at': as_of, 'reviewer': 'synthetic reviewer',
                }],
            }],
            'supplementary_sources': [{
                'snapshot': index.relative_to(directory).as_posix(),
                'sha256': common.sha(index),
            }],
        }
        path = directory / 'registry.json'
        common.dump(path, registry)
        return path

    def inherit_registry(self, registry):
        destination = self.previous / 'inputs/amendments/rule262'
        if destination.exists():
            shutil.rmtree(destination)
        shutil.copytree(registry.parent, destination)
        meta = common.read(self.previous / 'portable.json')
        meta['rules_hash'] = service.rules_hash(self.previous / 'inputs')
        common.dump(self.previous / 'portable.json', meta)

    def state_files(self):
        # SQLite files here belong only to the temporary fixture. WAL sidecars
        # are transient and are not durable application state.
        return {
            p.relative_to(self.root).as_posix(): common.sha(p)
            for base in ('state', 'cvs', 'exports')
            for p in (self.root / base).rglob('*')
            if p.is_file() and not p.name.endswith(('-wal', '-shm'))
        }

    def assert_rejected_without_mutation(self, operation, message=None):
        before = self.state_files()
        if message:
            with self.assertRaisesRegex(ValueError, message):
                operation()
        else:
            with self.assertRaises(ValueError):
                operation()
        self.assertEqual(self.state_files(), before)

    @staticmethod
    def moscow_boundary_clock():
        class FrozenDateTime(datetime):
            @classmethod
            def now(cls, tz=None):
                instant = datetime(2026, 11, 3, 21, 15, tzinfo=timezone.utc)
                return instant.astimezone(tz) if tz else instant.replace(tzinfo=None)
        return FrozenDateTime

    def test_default_date_uses_moscow_calendar_and_dated_output(self):
        with patch.object(batch, 'datetime', self.moscow_boundary_clock()):
            run, result = self.import_on()
        self.assertFalse(result['resumed'])
        self.assertEqual(common.read(run / 'manifest.json')['as_of'], '2026-11-04')
        self.assertEqual(common.read(run / 'archive_manifest.json')['assessment_date'],
                         '2026-11-04')
        self.assertEqual(Path(common.read(run / 'portable.json')['output_path']).name,
                         'newCVs_2026-11-04.xlsx')

    def test_explicit_iso_date_overrides_clock_and_accepts_leap_day(self):
        with patch.object(batch, 'datetime', self.moscow_boundary_clock()):
            run, _ = self.import_on('2028-02-29')
        self.assertEqual(common.read(run / 'manifest.json')['as_of'], '2028-02-29')
        self.assertEqual(common.read(run / 'portable.json')['assessment_date'],
                         '2028-02-29')

    def test_noncanonical_or_invalid_explicit_dates_fail_before_changes(self):
        for invalid in ('', '20261104', '2026-W45-3', '2026-11-4',
                        ' 2026-11-04', '2026-11-04 ', '2026-02-29',
                        '2026-11-04T00:00:00', True, 20261104, date(2026, 11, 4)):
            with self.subTest(as_of=invalid):
                self.assert_rejected_without_mutation(
                    lambda: self.import_on(invalid), 'YYYY-MM-DD')

    def test_same_date_resumes_from_imported_active_run_without_erasing_progress(self):
        run, first = self.import_on('2026-11-04', output=self.output)
        connection = common.connect(run / 'state.sqlite')
        with connection:
            connection.execute("UPDATE candidates SET extraction_status='error', "
                               "error='Отложенная проверка' WHERE id='101'")
        connection.close()
        original_inputs = {p.relative_to(run / 'inputs').as_posix(): common.sha(p)
                           for p in (run / 'inputs').rglob('*') if p.is_file()}
        again, second = self.import_on('2026-11-04', output=self.output)
        self.assertFalse(first['resumed'])
        self.assertTrue(second['resumed'])
        self.assertEqual(again, run)
        connection = common.connect(run / 'state.sqlite')
        self.assertEqual(connection.execute(
            "SELECT error FROM candidates WHERE id='101'").fetchone()[0],
            'Отложенная проверка')
        connection.close()
        self.assertEqual(original_inputs,
                         {p.relative_to(run / 'inputs').as_posix(): common.sha(p)
                          for p in (run / 'inputs').rglob('*') if p.is_file()})

    def test_identical_archive_on_different_dates_creates_independent_runs(self):
        first, _ = self.import_on('2026-11-04')
        second, result = self.import_on('2026-11-05')
        self.assertNotEqual(first, second)
        self.assertFalse(result['resumed'])
        self.assertNotEqual(common.read(first / 'portable.json')['batch_key'],
                            common.read(second / 'portable.json')['batch_key'])
        self.assertEqual(common.read(first / 'manifest.json')['as_of'], '2026-11-04')
        self.assertEqual(common.read(second / 'manifest.json')['as_of'], '2026-11-05')
        connection = common.connect(second / 'state.sqlite')
        self.assertEqual(connection.execute('SELECT COUNT(*) FROM candidates').fetchone()[0], 3)
        self.assertEqual(connection.execute(
            "SELECT COUNT(*) FROM candidates WHERE result IS NOT NULL "
            "OR review_status!='pending'").fetchone()[0], 0)
        connection.close()

    def test_legacy_key_resumes_only_when_manifest_date_matches(self):
        run, _ = self.import_on('2026-10-22')
        meta = common.read(run / 'portable.json')
        legacy_fields = {key: meta[key] for key in (
            'archive_sha256', 'template_sha256', 'criteria_sha256', 'implementation_hash')}
        meta['batch_key'] = hashlib.sha256(
            json.dumps(legacy_fields, sort_keys=True).encode()).hexdigest()
        meta.pop('assessment_date')
        common.dump(run / 'portable.json', meta)
        before = (run / 'portable.json').read_bytes()
        same, resumed = self.import_on('2026-10-22')
        self.assertEqual(same, run)
        self.assertTrue(resumed['resumed'])
        different, fresh = self.import_on('2026-11-04')
        self.assertNotEqual(different, run)
        self.assertFalse(fresh['resumed'])
        self.assertEqual((run / 'portable.json').read_bytes(), before)
        self.assertEqual(common.read(run / 'manifest.json')['as_of'], '2026-10-22')

    def test_current_inherited_registry_is_frozen_without_explicit_override(self):
        registry = self.registry_package('inherited', '2026-11-04')
        self.inherit_registry(registry)
        run, _ = self.import_on('2026-11-04')
        target = run / 'inputs/amendments/rule262'
        self.assertEqual(common.sha(target / 'registry.json'), common.sha(registry))
        for p in registry.parent.rglob('*'):
            if p.is_file():
                self.assertEqual(common.sha(target / p.relative_to(registry.parent)), common.sha(p))

    def test_old_inherited_registry_is_not_silently_reused_on_default_new_date(self):
        self.inherit_registry(self.registry_package('old', '2026-10-22'))
        with patch.object(batch, 'datetime', self.moscow_boundary_clock()):
            self.assert_rejected_without_mutation(lambda: self.import_on())

    def test_explicit_registry_with_wrong_date_is_rejected_before_changes(self):
        registry = self.registry_package('old', '2026-10-22')
        self.assert_rejected_without_mutation(
            lambda: self.import_on('2026-11-04', registry))

    def test_override_replaces_old_package_copies_only_proofs_and_resumes(self):
        old = self.registry_package('old', '2026-10-22')
        self.inherit_registry(old)
        registry = self.registry_package('current', '2026-11-04')
        (registry.parent / 'unrelated-private.txt').write_text('Не переносить', encoding='utf-8')
        run, first = self.import_on('2026-11-04', registry)
        target = run / 'inputs/amendments/rule262'
        self.assertFalse(first['resumed'])
        expected = {'registry.json', 'sources/current.txt', 'sources/current.html',
                    'sources/current-index.html'}
        self.assertEqual({p.relative_to(target).as_posix()
                          for p in target.rglob('*') if p.is_file()}, expected)
        hashes = common.read(run / 'manifest.json')['input_hashes']
        for name in expected:
            self.assertEqual(hashes['amendments/rule262/' + name], common.sha(target / name))
        self.assertFalse(any('sources/old' in name for name in hashes))
        same, resumed = self.import_on('2026-11-04', registry)
        self.assertEqual(same, run)
        self.assertTrue(resumed['resumed'])
        # The frozen, now-current package has the same effective criteria.
        same, inherited = self.import_on('2026-11-04')
        self.assertEqual(same, run)
        self.assertTrue(inherited['resumed'])
        self.assertEqual(common.sha(self.previous / 'inputs/amendments/rule262/registry.json'),
                         common.sha(old))

    def test_effective_override_key_ignores_replaced_old_package(self):
        self.inherit_registry(self.registry_package('old-a', '2026-10-22'))
        registry = self.registry_package('current', '2026-11-04')
        run, _ = self.import_on('2026-11-04', registry)
        self.inherit_registry(self.registry_package('old-b', '2026-10-23'))
        common.dump(self.root / 'state/active.json', {'run': self.previous.name})
        same, resumed = self.import_on('2026-11-04', registry)
        self.assertEqual(same, run)
        self.assertTrue(resumed['resumed'])

    def test_changed_effective_registry_creates_new_run_on_same_date(self):
        first, _ = self.import_on('2026-11-04', self.registry_package('version-a', '2026-11-04'))
        old_hash = common.sha(first / 'inputs/amendments/rule262/registry.json')
        second, result = self.import_on('2026-11-04', self.registry_package('version-b', '2026-11-04'))
        self.assertNotEqual(second, first)
        self.assertFalse(result['resumed'])
        self.assertEqual(common.sha(first / 'inputs/amendments/rule262/registry.json'), old_hash)
        self.assertNotEqual(common.read(first / 'portable.json')['batch_key'],
                            common.read(second / 'portable.json')['batch_key'])

    def test_tampered_snapshot_and_unsafe_reference_fail_before_changes(self):
        registry = self.registry_package('tampered', '2026-11-04')
        (registry.parent / 'sources/tampered.html').write_text('Подмена', encoding='utf-8')
        self.assert_rejected_without_mutation(
            lambda: self.import_on('2026-11-04', registry))
        registry = self.registry_package('unsafe', '2026-11-04')
        content = copy.deepcopy(common.read(registry))
        content['entities'][0]['evidence'][0]['snapshot'] = '../outside.txt'
        common.dump(registry, content)
        self.assert_rejected_without_mutation(
            lambda: self.import_on('2026-11-04', registry))

    def test_cli_forwards_assessment_date_and_explicit_registry(self):
        real_signature = inspect.signature(batch.import_archive)
        registry = self.registry_package('cli', '2026-11-04')
        argv = ['hrhive', 'import-archive', '--file', str(self.archive),
                '--template', str(self.template), '--output', str(self.output),
                '--as-of', '2026-11-04', '--partner-registry', str(registry)]
        with patch('sys.argv', argv), patch.object(batch, 'import_archive',
                return_value={'run': 'synthetic'}) as imported, \
                patch('sys.stdout', new_callable=io.StringIO):
            cli.main()
        imported.assert_called_once()
        bound = real_signature.bind(*imported.call_args.args, **imported.call_args.kwargs)
        self.assertEqual(bound.arguments['as_of'], '2026-11-04')
        self.assertEqual(Path(bound.arguments['partner_registry']), registry)
        self.assertEqual(Path(bound.arguments['output']), self.output)


if __name__ == '__main__':
    unittest.main()
