# 데이터가 유효하지 않을 때 발생.
import pandas


class InvalidDataException(Exception):
    def __init__(self, message: str):
        super().__init__(message)
