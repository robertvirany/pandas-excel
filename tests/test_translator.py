#%% Imports
import unittest
from datetime import datetime

from formula_parser import parse_formula
from translate_formulas import (
    FormulaCell,
    WorkbookContext,
    identify_derived_columns,
    translate_workbook,
)


def make_context() -> WorkbookContext:
    header_map = {
        'Sales': {'A': 'Region', 'B': 'Year', 'C': 'Sales', 'D': 'Delta'},
        'Parameters': {'A': 'Date'},
    }
    cell_values = {
        'Sales': {
            'A1': 'Region',
            'B1': 'Year',
            'C1': 'Sales',
            'D1': 'Delta',
            'A2': 'North',
            'B2': 2022,
            'C2': 100,
            'D2': None,
            'A3': 'South',
            'B3': 2022,
            'C3': 150,
            'D3': None,
        },
        'Parameters': {
            'A1': datetime(2022, 1, 15),
        },
    }
    return WorkbookContext(header_map=header_map, cell_values=cell_values)


def translate(formula: str):
    context = make_context()
    cell = FormulaCell(
        sheet='Sales',
        address='D2',
        formula=formula,
        ast=parse_formula(formula),
        parse_error=None,
    )
    translations = list(translate_workbook(context, [cell]))
    assert len(translations) == 1
    return translations[0]


class TranslatorTests(unittest.TestCase):
    def test_sumifs_ratio_translation(self):
        formula = "=SUMIFS(C:C, A:A, \"North\")/SUMIFS(C:C, A:A, \"South\")"
        translation = translate(formula)
        expected = "(dfs['Sales'].loc[(dfs['Sales']['Region'] == 'North')]['Sales'].sum()) / (dfs['Sales'].loc[(dfs['Sales']['Region'] == 'South')]['Sales'].sum())"
        self.assertFalse(translation.warnings)
        self.assertEqual(translation.expression, expected)

    def test_sumifs_with_eomonth(self):
        formula = '=SUMIFS(C:C, B:B, ">=" & EOMONTH(Parameters!A1, -1))'
        translation = translate(formula)
        expected = "dfs['Sales'].loc[(dfs['Sales']['Year'] >= '2021-12-31 00:00:00')]['Sales'].sum()"
        self.assertEqual(translation.expression, expected)

    def test_index_match_translation(self):
        formula = "=INDEX(C:C, MATCH(\"North\", A:A, 0))"
        translation = translate(formula)
        expected = "dfs['Sales']['Sales'].loc[(dfs['Sales']['Region'] == 'North')].iloc[0]"
        self.assertEqual(translation.expression, expected)

    def test_if_expression(self):
        formula = "=IF(SUMIFS(C:C, A:A, \"North\")>0, 1, 0)"
        translation = translate(formula)
        expected = "(1) if ((dfs['Sales'].loc[(dfs['Sales']['Region'] == 'North')]['Sales'].sum()) > (0)) else (0)"
        self.assertEqual(translation.expression, expected)

    def test_unsupported_function(self):
        formula = "=FOO(1)"
        translation = translate(formula)
        self.assertIsNone(translation.expression)
        self.assertTrue(any('Unsupported function FOO' in warning for warning in translation.warnings))

    def test_identify_uniform_derived_column(self):
        context = make_context()
        cells = [
            FormulaCell(
                sheet='Sales',
                address='D2',
                formula='=C2-10',
                ast=parse_formula('=C2-10'),
                parse_error=None,
            ),
            FormulaCell(
                sheet='Sales',
                address='D3',
                formula='=C3-10',
                ast=parse_formula('=C3-10'),
                parse_error=None,
            ),
        ]
        collapsed = identify_derived_columns(context, cells)
        self.assertEqual(len(collapsed), 1)
        base = collapsed[0]
        self.assertEqual(base.derived_column, 'D')
        self.assertEqual(base.derived_row_span, (2, 3))

    def test_translate_workbook_for_derived_column(self):
        context = make_context()
        cells = [
            FormulaCell(
                sheet='Sales',
                address='D2',
                formula='=C2-10',
                ast=parse_formula('=C2-10'),
                parse_error=None,
            ),
            FormulaCell(
                sheet='Sales',
                address='D3',
                formula='=C3-10',
                ast=parse_formula('=C3-10'),
                parse_error=None,
            ),
        ]
        collapsed = identify_derived_columns(context, cells)
        translations = list(translate_workbook(context, collapsed))
        self.assertEqual(len(translations), 1)
        translation = translations[0]
        self.assertEqual(translation.output_kind, 'column')
        self.assertEqual(translation.target_column_name, 'Delta')
        expected = "dfs['Sales'].apply(lambda row: (row['Sales']) - (10), axis=1)"
        self.assertEqual(translation.expression, expected)


if __name__ == '__main__':
    unittest.main()
