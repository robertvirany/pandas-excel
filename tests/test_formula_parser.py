import unittest

from formula_parser import (
    BinaryOpNode,
    FunctionNode,
    NumberNode,
    ReferenceNode,
    UnaryOpNode,
    parse_formula,
)


class FormulaParserTests(unittest.TestCase):
    def test_parse_sumifs_ratio(self):
        ast = parse_formula("=SUMIFS(C:C, A:A, \"North\")/SUMIFS(C:C, A:A, \"South\")")
        self.assertIsInstance(ast, BinaryOpNode)
        self.assertEqual(ast.operator, '/')
        self.assertIsInstance(ast.left, FunctionNode)
        self.assertEqual(ast.left.name.upper(), 'SUMIFS')
        self.assertIsInstance(ast.right, FunctionNode)
        self.assertEqual(ast.right.name.upper(), 'SUMIFS')

    def test_parse_arithmetic(self):
        ast = parse_formula("=(A1+B1)*C1")
        self.assertIsInstance(ast, BinaryOpNode)
        self.assertEqual(ast.operator, '*')
        self.assertIsInstance(ast.left, BinaryOpNode)
        self.assertEqual(ast.left.operator, '+')
        self.assertIsInstance(ast.right, ReferenceNode)

    def test_parse_function_arguments(self):
        ast = parse_formula("=EOMONTH($A$1, -1)+1")
        self.assertIsInstance(ast, BinaryOpNode)
        self.assertIsInstance(ast.left, FunctionNode)
        self.assertEqual(ast.left.name.upper(), 'EOMONTH')
        first_arg = ast.left.args[0]
        second_arg = ast.left.args[1]
        self.assertIsInstance(first_arg, ReferenceNode)
        if isinstance(second_arg, NumberNode):
            self.assertEqual(second_arg.value, -1.0)
        else:
            self.assertIsInstance(second_arg, UnaryOpNode)
            self.assertEqual(second_arg.operator, '-')
            self.assertIsInstance(second_arg.operand, NumberNode)


if __name__ == '__main__':
    unittest.main()
