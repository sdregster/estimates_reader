from openpyxl import load_workbook
from etc.entities import Chapter, Subchapter, Work, Material, MiM


class EstimateBaseTemplate:
    def __init__(self, file_path: str) -> None:
        self.__wb = load_workbook(file_path, data_only=True)
        self.__ws = self.__wb.active
        self.estimate_file_name = file_path.split("\\")[-1]
        self._get_estimate_keystone_data()
        self.rows = []

    def _get_estimate_keystone_data(self):
        def _get_group_code(value: str) -> str:
            temp_result = value[value.rfind("(") + 1 :]
            return temp_result[: temp_result.rfind(")")]

        self.estimate_total_number = (
            str(self.__ws.cell(9, 6).value).split("№")[1].strip()
        )
        if " изм." in self.estimate_total_number:
            (
                self.estimate_number,
                self.estimate_version,
            ) = self.estimate_total_number.split(" изм.")
        else:
            self.estimate_number, self.estimate_version = self.estimate_total_number, 0
        self.estimate_work_name = str(self.__ws.cell(12, 4).value).strip()
        self.estimate_group_code = _get_group_code(self.estimate_work_name)
        self.estimate_reason = (
            str(self.__ws.cell(15, 3).value).split("Основание: ")[1].strip()
        )
        self.estimate_cipher = self.estimate_reason
        self.estimate_time_period = (
            str(self.__ws.cell(19, 6).value).replace("_", "").strip()
        )

        for row in range(1, self.__ws.max_row + 1):
            target_value = self.__ws.cell(row, 1).value
            if isinstance(target_value, str) and "№ пп" in target_value:
                self.content_start_from = row + 4
            elif isinstance(target_value, str) and "ФОТ" in target_value:
                self.estimate_wage_fund = self.__ws.cell(row, 8).value
            elif isinstance(target_value, str) and "ВСЕГО по смете" in target_value:
                self.estimate_cost = self.__ws.cell(row, 8).value
                self.estimate_laboriousness = self.__ws.cell(row, 13).value

    def read_rows(self):
        for row in range(self.content_start_from, self.__ws.max_row + 1):
            # нужна проверка на попадание строки в какую-нибудь сводку (по разделу или по смете) #TODO

            current_row_values = [
                self.__ws.cell(row, col).value
                for col in range(1, self.__ws.max_column + 1)
            ]

            nn_value = self.__ws.cell(row, 1).value
            reason_value = self.__ws.cell(row, 2).value
            unit_value = self.__ws.cell(row, 4).value
            # total_cost_value = self.__ws.cell(row, 8).value

            # Раздел
            if isinstance(nn_value, str) and "Раздел" in nn_value:
                self.rows.append(Chapter(current_row_values))

            # # Подраздел
            # elif (
            #     isinstance(nn_value, str)
            #     and "Раздел" not in nn_value
            #     and not total_cost_value
            #     and not unit_value
            # ):
            #     black_list = [
            #         "должность",
            #         "проверил",
            #         "составил",
            #         "итоги по смете",
            #         "в том числе",
            #     ]
            #     flag = False
            #     for elem in black_list:
            #         if elem not in nn_value.lower():
            #             flag = True
            #         else:
            #             flag = False

            #     if flag:
            #         self.rows.append(Subchapter(current_row_values))

            # Работа
            elif isinstance(reason_value, str) and nn_value and "ГЭСН" in reason_value:
                self.rows.append(Work(current_row_values))

            # Материал
            elif (
                isinstance(reason_value, str)
                and nn_value
                and nn_value != "Н"
                and "ГЭСН" not in reason_value
            ):
                self.rows.append(Material(current_row_values))

            # МиМ
            elif (
                isinstance(reason_value, str)
                and not nn_value
                and "маш.час" in unit_value
            ):
                self.rows.append(MiM(current_row_values))
