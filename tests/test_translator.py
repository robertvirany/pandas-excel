import unittest
from datetime import datetime

from formula_parser import parse_formula
from translate_formulas import FormulaCell, WorkbookContext, translate_workbook


def make_context() -> WorkbookContext:
    header_map = {
        'Sales': {'A': 'Region', 'B': 'Year', 'C': 'Sales'},
        'Parameters': {'A': 'Date'},
    }
    cell_values = {
        'Sales': {
            'A1': 'Region',
            'B1': 'Year',
            'C1': 'Sales',
            'A2': 'North',
            'B2': 2022,
            'C2': 100,
            'A3': 'South',
            'B3': 2022,
            'C3': 150,
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


if __name__ == '__main__':
    unittest.main()
