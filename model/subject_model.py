from enum import unique, Enum


class Subject:

    def __init__(self, code: str, name: str):
        self.code = code
        self.name = name

    def __hash__(self):
        return hash(self.code)

    def __eq__(self, other):
        return self.code == other.code

    def __str__(self):
        return self.name


@unique
class Subjects(Enum):
    CHINESE = Subject('chinese', '语文')
    MATH = Subject('math', '数学')
    ENGLISH = Subject('english', '英语')
    TWO = Subject('two', '总评')
