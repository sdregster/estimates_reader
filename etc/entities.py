from dataclasses import dataclass
from functools import singledispatchmethod


@dataclass
class Chapter:
    name: str

    @singledispatchmethod
    def __init__(self, _) -> None:
        raise TypeError("Arguments of this type are not supported.")

    @__init__.register(list)
    def _process_list(self, value):
        self.name = str(value[0]).strip()

    def __str__(self) -> str:
        return f"{self.__class__} - {self.name}"


@dataclass
class Subchapter:
    name: str

    @singledispatchmethod
    def __init__(self, _) -> None:
        raise TypeError("Arguments of this type are not supported.")

    @__init__.register(list)
    def _process_list(self, value):
        self.name = str(value[0]).strip()

    def __str__(self) -> str:
        return f"{self.__class__} - {self.name}"


@dataclass
class Work:
    total_name: str
    index: int
    reason: str
    name: str
    unit: str
    amount: int | float
    cost_per_unit: int | float
    total_cost: int | float
    total_wage: int | float
    mim_cost: int | float
    mim_wage: int | float
    materials_cost: int | float
    laboriousness: int | float
    mim_laboriousness: int | float

    @singledispatchmethod
    def __init__(self, _) -> None:
        raise TypeError("Arguments of this type are not supported.")

    @__init__.register(list)
    def _process_list(self, row_values):
        self.total_name = ". ".join([str(row_values[0]), str(row_values[2])]).strip()
        self.name = str(row_values[2]).strip()
        self.index = int(row_values[0])
        self.reason = str(row_values[1]).strip()
        self.unit = str(row_values[3]).strip()
        self.amount = float(row_values[5]) if row_values[5] else 0
        self.cost_per_unit = float(row_values[6]) if row_values[6] else 0
        self.total_cost = float(row_values[7]) if row_values[7] else 0
        self.total_wage = float(row_values[8]) if row_values[8] else 0
        self.mim_cost = float(row_values[9]) if row_values[9] else 0
        self.mim_wage = float(row_values[10]) if row_values[10] else 0
        self.materials_cost = float(row_values[11]) if row_values[11] else 0
        self.laboriousness = float(row_values[12]) if row_values[12] else 0
        self.mim_laboriousness = float(row_values[13]) if row_values[13] else 0

    def __str__(self) -> str:
        return f"{self.__class__} - {self.total_name}"


@dataclass
class Material:
    total_name: str
    index: int
    reason: str
    name: str
    unit: str
    amount: int | float
    cost_per_unit: int | float
    total_cost: int | float
    total_wage: int | float
    mim_cost: int | float
    mim_wage: int | float
    materials_cost: int | float
    laboriousness: int | float
    mim_laboriousness: int | float

    @singledispatchmethod
    def __init__(self, _) -> None:
        raise TypeError("Arguments of this type are not supported.")

    @__init__.register(list)
    def _process_list(self, row_values):
        self.total_name = ". ".join([str(row_values[0]), str(row_values[2])]).strip()
        self.name = str(row_values[2]).strip()
        self.index = int(row_values[0])
        self.reason = str(row_values[1]).strip()
        self.unit = str(row_values[3]).strip()
        self.amount = float(row_values[5]) if row_values[5] else 0
        self.cost_per_unit = float(row_values[6]) if row_values[6] else 0
        self.total_cost = float(row_values[7]) if row_values[7] else 0
        self.total_wage = float(row_values[8]) if row_values[8] else 0
        self.mim_cost = float(row_values[9]) if row_values[9] else 0
        self.mim_wage = float(row_values[10]) if row_values[10] else 0
        self.materials_cost = float(row_values[11]) if row_values[11] else 0
        self.laboriousness = float(row_values[12]) if row_values[12] else 0
        self.mim_laboriousness = float(row_values[13]) if row_values[13] else 0

    def __str__(self) -> str:
        return f"{self.__class__} - {self.total_name}"


@dataclass
class MiM:
    reason: str
    name: str
    unit: str
    amount: int | float
    cost_per_unit: int | float
    total_cost: int | float
    total_wage: int | float
    mim_cost: int | float
    mim_wage: int | float
    materials_cost: int | float
    laboriousness: int | float
    mim_laboriousness: int | float

    @singledispatchmethod
    def __init__(self, _) -> None:
        raise TypeError("Arguments of this type are not supported.")

    @__init__.register(list)
    def _process_list(self, row_values):
        self.name = str(row_values[2]).strip()
        self.reason = str(row_values[1]).strip()
        self.unit = str(row_values[3]).strip()
        self.amount = float(row_values[5]) if row_values[5] else 0
        self.cost_per_unit = float(row_values[6]) if row_values[6] else 0
        self.total_cost = float(row_values[7]) if row_values[7] else 0
        self.total_wage = float(row_values[8]) if row_values[8] else 0
        self.mim_cost = float(row_values[9]) if row_values[9] else 0
        self.mim_wage = float(row_values[10]) if row_values[10] else 0
        self.materials_cost = float(row_values[11]) if row_values[11] else 0
        self.laboriousness = float(row_values[12]) if row_values[12] else 0
        self.mim_laboriousness = float(row_values[13]) if row_values[13] else 0

    def __str__(self) -> str:
        return f"{self.__class__} - {self.name}"
