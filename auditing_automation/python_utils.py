from types import GeneratorType


def get_columns(data: GeneratorType) -> str:
    return next(data)[0:]
