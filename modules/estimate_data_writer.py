from modules.estimate_data_collector import EstimateBaseTemplate
from etc.entities import Chapter, Work, Material, MiM
from openpyxl import load_workbook

from pprint import pprint


class NeosintezTemplate:
    class_levels = {"Смета": 1, "Раздел сметы": 2, "Строка сметы": 3, "МиМ сметы": 4}

    def __init__(self, obj: EstimateBaseTemplate) -> None:
        self.keystone_data = obj.__dict__
        self.__wb = load_workbook("data/template.xlsx")
        self.__ws = self.__wb.active
        self.cursor = 2

    def _write_header(self):
        self.__ws.cell(self.cursor, 1).value = self.keystone_data["estimate_group_code"]
        self.__ws.cell(self.cursor, 2).value = self.cursor - 1
        self.__ws.cell(self.cursor, 4).value = "Смета"
        self.__ws.cell(self.cursor, 9).value = self.keystone_data[
            "estimate_total_number"
        ]
        self.__ws.cell(self.cursor, 23).value = self.keystone_data["estimate_work_name"]
        self.__ws.cell(self.cursor, 24).value = self.keystone_data["estimate_number"]
        self.__ws.cell(self.cursor, 25).value = int(
            self.keystone_data["estimate_version"]
        )
        self.__ws.cell(self.cursor, 26).value = self.keystone_data["estimate_reason"]
        self.__ws.cell(self.cursor, 27).value = self.keystone_data["estimate_cipher"]
        self.__ws.cell(self.cursor, 28).value = self.keystone_data["estimate_cost"]
        self.__ws.cell(self.cursor, 29).value = self.keystone_data["estimate_wage_fund"]
        self.__ws.cell(self.cursor, 30).value = self.keystone_data[
            "estimate_laboriousness"
        ]
        self.__ws.cell(self.cursor, 31).value = self.keystone_data[
            "estimate_time_period"
        ]
        self.__ws.cell(self.cursor, 32).value = self.keystone_data["estimate_file_name"]

        self.cursor += 1

    def _write_chapter(self, obj: Chapter):
        class_name = "Раздел сметы"

        self.__ws.cell(self.cursor, 1).value = self.keystone_data["estimate_group_code"]
        self.__ws.cell(self.cursor, 2).value = self.class_levels[class_name]
        self.__ws.cell(self.cursor, 4).value = class_name
        self.__ws.cell(self.cursor, 5).value = obj.name
        self.__ws.cell(self.cursor, 9).value = obj.name

        self.cursor += 1

    def _write_work(self, obj: Work):
        class_name = "Строка сметы"

        self.__ws.cell(self.cursor, 1).value = self.keystone_data["estimate_group_code"]
        self.__ws.cell(self.cursor, 2).value = self.class_levels[class_name]
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

        self.__ws.cell(self.cursor, 1).value = self.keystone_data["estimate_group_code"]
        self.__ws.cell(self.cursor, 2).value = self.class_levels[class_name]
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

        self.__ws.cell(self.cursor, 1).value = self.keystone_data["estimate_group_code"]
        self.__ws.cell(self.cursor, 2).value = self.class_levels[class_name]
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
        self.__wb.save(f"{self.keystone_data['estimate_total_number']}_nt.xlsx")
        self.__wb.close()

    def export(self):
        # header
        self._write_header()

        # iterate rows
        for row in self.keystone_data["rows"]:
            if isinstance(row, Chapter):
                self._write_chapter(row)

            if isinstance(row, Work):
                self._write_work(row)

            if isinstance(row, Material):
                self._write_material(row)

            if isinstance(row, MiM):
                self._write_mim(row)

        # finish
        self._finish()
