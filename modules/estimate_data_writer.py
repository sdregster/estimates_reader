from datetime import date
from time import time

from openpyxl import load_workbook

from etc.entities import Chapter, Material, MiM, Subchapter, Work
from modules.estimate_data_collector import EstimateBaseTemplate


class NeosintezTemplate:
    def __init__(self) -> None:
        self.__wb = load_workbook("data/template.xlsx")
        self.__ws = self.__wb.active
        self.cursor = 2
        
        self.temp_estimate_number = None
        self.temp_chapter_level = 2

    def _write_header(self, obj: EstimateBaseTemplate):
        class_name = "Смета"
        self.__ws.cell(self.cursor, 1).value = obj.estimate_total_number
        self.__ws.cell(self.cursor, 2).value = 1
        self.__ws.cell(self.cursor, 4).value = class_name
        self.__ws.cell(self.cursor, 9).value = obj.estimate_total_number
        self.__ws.cell(self.cursor, 23).value = obj.estimate_work_name
        self.__ws.cell(self.cursor, 24).value = obj.estimate_number
        self.__ws.cell(self.cursor, 25).value = int(obj.estimate_version)
        self.__ws.cell(self.cursor, 26).value = obj.estimate_reason
        self.__ws.cell(self.cursor, 27).value = obj.estimate_cipher
        self.__ws.cell(self.cursor, 28).value = obj.estimate_cost
        self.__ws.cell(self.cursor, 29).value = obj.estimate_wage_fund
        self.__ws.cell(self.cursor, 30).value = obj.estimate_laboriousness
        self.__ws.cell(self.cursor, 31).value = obj.estimate_time_period
        self.__ws.cell(self.cursor, 32).value = obj.estimate_file_name

        self.cursor += 1

    def _write_chapter(self, obj: Chapter):
        class_name = "Раздел сметы"

        self.__ws.cell(self.cursor, 1).value = self.temp_estimate_number
        self.__ws.cell(self.cursor, 2).value = 2
        self.__ws.cell(self.cursor, 4).value = class_name
        self.__ws.cell(self.cursor, 5).value = obj.name
        self.__ws.cell(self.cursor, 9).value = obj.name

        self.cursor += 1

    def _write_subchapter(self, obj: Subchapter):
        class_name = "Подраздел сметы"

        self.__ws.cell(self.cursor, 1).value = self.temp_estimate_number
        self.__ws.cell(self.cursor, 2).value = self.temp_chapter_level
        self.__ws.cell(self.cursor, 4).value = class_name
        self.__ws.cell(self.cursor, 5).value = obj.name
        self.__ws.cell(self.cursor, 9).value = obj.name

        self.cursor += 1

    def _write_work(self, obj: Work):
        class_name = "Строка сметы"

        self.__ws.cell(self.cursor, 1).value = self.temp_estimate_number
        self.__ws.cell(self.cursor, 2).value = self.temp_chapter_level + 1
        self.__ws.cell(self.cursor, 4).value = class_name
        self.__ws.cell(self.cursor, 8).value = "Работа"
        self.__ws.cell(self.cursor, 9).value = obj.total_name
        self.__ws.cell(self.cursor, 10).value = obj.index
        self.__ws.cell(self.cursor, 11).value = obj.reason
        self.__ws.cell(self.cursor, 12).value = obj.name
        self.__ws.cell(self.cursor, 13).value = obj.unit
        self.__ws.cell(self.cursor, 14).value = obj.amount
        self.__ws.cell(self.cursor, 15).value = obj.cost_per_unit
        self.__ws.cell(self.cursor, 16).value = obj.total_cost
        self.__ws.cell(self.cursor, 17).value = obj.total_wage
        self.__ws.cell(self.cursor, 18).value = obj.mim_cost
        self.__ws.cell(self.cursor, 19).value = obj.mim_wage
        self.__ws.cell(self.cursor, 20).value = obj.materials_cost
        self.__ws.cell(self.cursor, 21).value = obj.laboriousness
        self.__ws.cell(self.cursor, 22).value = obj.mim_laboriousness

        self.cursor += 1

    def _write_material(self, obj: Material):
        class_name = "Строка сметы"

        self.__ws.cell(self.cursor, 1).value = self.temp_estimate_number
        self.__ws.cell(self.cursor, 2).value = self.temp_chapter_level + 1
        self.__ws.cell(self.cursor, 4).value = class_name
        self.__ws.cell(self.cursor, 8).value = "МТР-Материалы"
        self.__ws.cell(self.cursor, 9).value = obj.total_name
        self.__ws.cell(self.cursor, 10).value = obj.index
        self.__ws.cell(self.cursor, 11).value = obj.reason
        self.__ws.cell(self.cursor, 12).value = obj.name
        self.__ws.cell(self.cursor, 13).value = obj.unit
        self.__ws.cell(self.cursor, 14).value = obj.amount
        self.__ws.cell(self.cursor, 15).value = obj.cost_per_unit
        self.__ws.cell(self.cursor, 16).value = obj.total_cost
        self.__ws.cell(self.cursor, 17).value = obj.total_wage
        self.__ws.cell(self.cursor, 18).value = obj.mim_cost
        self.__ws.cell(self.cursor, 19).value = obj.mim_wage
        self.__ws.cell(self.cursor, 20).value = obj.materials_cost
        self.__ws.cell(self.cursor, 21).value = obj.laboriousness
        self.__ws.cell(self.cursor, 22).value = obj.mim_laboriousness

        self.cursor += 1

    def _write_mim(self, obj: MiM):
        class_name = "МиМ сметы"

        self.__ws.cell(self.cursor, 1).value = self.temp_estimate_number
        self.__ws.cell(self.cursor, 2).value = self.temp_chapter_level + 2
        self.__ws.cell(self.cursor, 4).value = class_name
        self.__ws.cell(self.cursor, 8).value = "МиМ"
        self.__ws.cell(self.cursor, 9).value = obj.name
        self.__ws.cell(self.cursor, 11).value = obj.reason
        self.__ws.cell(self.cursor, 12).value = obj.name
        self.__ws.cell(self.cursor, 13).value = obj.unit
        self.__ws.cell(self.cursor, 14).value = obj.amount
        self.__ws.cell(self.cursor, 15).value = obj.cost_per_unit
        self.__ws.cell(self.cursor, 16).value = obj.total_cost
        self.__ws.cell(self.cursor, 17).value = obj.total_wage
        self.__ws.cell(self.cursor, 18).value = obj.mim_cost
        self.__ws.cell(self.cursor, 19).value = obj.mim_wage
        self.__ws.cell(self.cursor, 20).value = obj.materials_cost
        self.__ws.cell(self.cursor, 21).value = obj.laboriousness
        self.__ws.cell(self.cursor, 22).value = obj.mim_laboriousness

        self.cursor += 1

    def _finish(self):
        self.__ws.title = 'result'
        self.__wb.save(f"output/{date.today()}_{int(time())}_neosintez_template.xlsx")
        self.__wb.close()

    def export(self, estimates_list: list[EstimateBaseTemplate]):
        for estimate in estimates_list:
            self.temp_estimate_number = estimate.estimate_total_number
            self._write_header(estimate)

            for row in estimate.rows:
                if isinstance(row, Chapter):
                    self.temp_chapter_level = 2 # Сбрасываем до уровня 2, при обнаружении нового раздела
                    self._write_chapter(row)

                if isinstance(row, Subchapter):
                    self.temp_chapter_level = 3 # Устанавливаем уровень 3, при обнаружении подраздела
                    self._write_subchapter(row)

                if isinstance(row, Work):
                    self._write_work(row)

                if isinstance(row, Material):
                    self._write_material(row)

                if isinstance(row, MiM):
                    self._write_mim(row)

        # finish
        self._finish()
