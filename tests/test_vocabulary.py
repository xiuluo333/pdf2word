import json
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import patch
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main
import generate_html


class Sheet:
    def __init__(self, rows):
        self.cells = {}
        for r, row in enumerate(rows, 1):
            for c, value in enumerate(row, 1):
                self.cell(r, c).value = value

    def cell(self, r, c):
        return self.cells.setdefault((r, c), SimpleNamespace(value=None, column=c))

    @property
    def max_column(self):
        return max(c for r, c in self.cells)

    @property
    def max_row(self):
        return max(r for r, c in self.cells)

    def __getitem__(self, r):
        return [self.cell(r, c) for c in range(1, self.max_column + 1)]


class VocabularyTests(unittest.TestCase):
    def completion(self):
        return dict.fromkeys(main.RESPONSE_FIELDS, 'value') | {
            'english_definition': 'a journey by sea', 'phonetic': '/ˈvɔɪɪdʒ/'
        }

    def test_schema_and_prompt_agree(self):
        self.assertEqual(set(json.loads(main.JSON_OUTPUT_FORMAT)), set(main.RESPONSE_FIELDS))
        prompt = main.build_user_prompt('voyage', 1, ['phonetic', 'english_definition'])
        self.assertIn(f'全部 {len(main.RESPONSE_FIELDS)} 个键', prompt)
        self.assertEqual(main.validate_completion_result(self.completion()), self.completion())
        for field, invalid in [('english_definition', '无'), ('phonetic', '无'), ('phonetic', 'IPA: /test/')]:
            with self.subTest(field=field, invalid=invalid), self.assertRaises(ValueError):
                main.validate_completion_result(self.completion() | {field: invalid})

    def test_existing_sheet_only_fills_missing_and_can_resume(self):
        sheet = Sheet([['word', 'translate', 'phonetic'], ['voyage', '航行', '/existing/'], ['voyage', '航行', None]])
        book = SimpleNamespace(active=sheet)
        requests = []

        def complete(word, frequency, missing, **kwargs):
            requests.append(list(missing))
            return self.completion()

        with patch.dict(sys.modules, {'openpyxl': SimpleNamespace(load_workbook=lambda _: book)}), \
             patch.object(main, 'request_completion', side_effect=complete), \
             patch.object(main, 'save_workbook_atomic'), patch('builtins.print'):
            for _ in range(2):
                main.fill_workbook(Path('input.xlsx'), Path('out.xlsx'), None, 0, 0, None)
        self.assertEqual(len(requests), 2)
        self.assertNotIn('phonetic', requests[0])
        self.assertIn('phonetic', requests[1])
        columns = main.find_columns(sheet)
        self.assertEqual(sheet.cell(2, columns['phonetic']).value, '/existing/')
        self.assertEqual(sheet.cell(3, columns['phonetic']).value, '/ˈvɔɪɪdʒ/')
        self.assertEqual(sheet.cell(2, columns['translate']).value, '航行')
        self.assertEqual(sheet.cell(2, columns['english_definition']).value, 'a journey by sea')
        headers = [cell.value for cell in sheet[1]]
        rows = generate_html._rows_from_values(headers, [(2, [c.value for c in sheet[2]])])
        self.assertEqual(rows[0]['english_definition'], 'a journey by sea')
        self.assertEqual(headers.count('English definition'), 1)

    def test_generator_accepts_definition_aliases_and_legacy_rows(self):
        for header in ('English definition', 'english_definition', '英译英', '英文释义'):
            rows = generate_html._rows_from_values(['word', header], [(2, ['voyage', 'a journey by sea'])])
            self.assertEqual(rows[0]['english_definition'], 'a journey by sea')
        rows = generate_html._rows_from_values(['word'], [(2, ['voyage'])])
        self.assertEqual(rows[0]['english_definition'], '')


if __name__ == '__main__':
    unittest.main()
