
def transform_none(value):
    return value


def transform_remove_first(value: str):
    """
    Remove the first symbol in value.
    E.g. `x03330` -> '03330'.
    """
    return value[1:]


def transform_inc(value: int):
    return value + 1
