
def transform_none(value):
    """
    Do not apply any transform to value.
    """
    return value


def transform_remove_first(value):
    """
    Remove the first symbol in value.
    E.g. `x03330` -> '03330'.
    """
    return value[1:]


def transform_inc(value):
    """
    Do not apply any transform to value.
    """
    return value + 1
