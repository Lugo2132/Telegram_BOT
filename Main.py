import unittest
import pandas as pd
from main import calculate_percentages 

class TestBotFunctions(unittest.TestCase):

    def setUp(self):
        data = {
            ('Месяц', 'Проверено'): [10, 20, 30],
            ('Месяц', 'Выдано'): [20, 30, 40],
            'ФИО преподавателя': ['Иванов И.И.', 'Петров П.П.', 'Сидоров С.С.']
        }
        self.df = pd.DataFrame(data)
        self.df.to_excel('test_data.xlsx', index=False)

    def test_calculate_percentages(self):
        result_df = calculate_percentages('test_data.xlsx')

        expected_data = {
            'ФИО преподавателя': ['Иванов И.И.', 'Петров П.П.', 'Сидоров С.С.'],
            'Процент проверенных ДЗ': [50.0, 66.67, 75.0]
        }
        expected_df = pd.DataFrame(expected_data)
        pd.testing.assert_frame_equal(result_df, expected_df, check_dtype=False)

    def tearDown(self):
        import os
        os.remove('test_data.xlsx')

if __name__ == '__main__':
    unittest.main()
