from decimal import Decimal
from typing import Any


class SubjectInfo:
    code: str
    name: str
    func: Any
    factor: Decimal

    def __init__(self, code: str, name: str, func: Any, factor: Decimal = Decimal('1')):
        self.code = code
        self.name = name
        self.func = func
        self.factor = factor

    def is_total_subject(self):
        return self.code == 'two'

    def is_high_subject(self):
        return self.code == 'english'
