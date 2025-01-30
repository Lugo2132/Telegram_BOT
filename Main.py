import unittest
import pandas as pd
from io import BytesIO
from main import calculate_percentages, calculate_issued_percentages

class TestBotFunctions(unittest.TestCase):

    def setUp(self):
        # Создаем тестовый DataFrame и сохраняем его в Excel файл
        data = {
            ('Месяц', 'Проверено'): [10, 20, 30],
            ('Месяц', 'Выдано'): [100, 200, 300],
            ('Неделя', 'Выдано'): [50, 100, 150],
            'ФИО преподавателя': ['Иванов И.И.', 'Петров П.П.', 'Сидоров С.С.']
        }
        self.df = pd.DataFrame(data)
        self.file_path = 'test_file.xlsx'
        self.df.to_excel(self.file_path, header=[0, 1], index=False)

    def tearDown(self):
        # Удаляем тестовый файл после завершения тестов
        import os
        os.remove(self.file_path)

    def test_calculate_percentages(self):
        result_df = calculate_percentages(self.file_path)
        expected_data = {
            'ФИО преподавателя': ['Иванов И.И.', 'Петров П.П.', 'Сидоров С.С.'],
            'Процент проверенных ДЗ': [10.0, 10.0, 10.0]
        }
        expected_df = pd.DataFrame(expected_data)
        pd.testing.assert_frame_equal(result_df, expected_df)

    def test_calculate_issued_percentages(self):
        result_df = calculate_issued_percentages(self.file_path)
        expected_data = {
            'ФИО преподавателя': ['Иванов И.И.', 'Петров П.П.', 'Сидоров С.С.'],
            'Процент выданного ДЗ (Месяц)': [10.0, 10.0, 10.0],
            'Процент выданного ДЗ (Неделя)': [10.0, 10.0, 10.0]
        }
        expected_df = pd.DataFrame(expected_data)
        pd.testing.assert_frame_equal(result_df, expected_df)

if __name__ == '__main__':
    unittest.main()
