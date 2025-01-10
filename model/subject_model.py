from enum import unique, Enum


@unique
class Subjects(Enum):
    CHINESE = ('chinese', '语文')
    MATH = ('math', '数学')
    ENGLISH = ('english', '英语')
    TWO = ('two', '总评')

    @classmethod
    def values(cls):
        return [subject.value for subject in cls]
