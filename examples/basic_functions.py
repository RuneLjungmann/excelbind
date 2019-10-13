from typing import List, Dict
import time
import datetime
import numpy as np
import pandas as pd

import excelbind


@excelbind.function
def add(x: float, y: float) -> float:
    return x + y


@excelbind.function
def mult(x: float, y: float) -> float:
    return x*y


@excelbind.function
def det(x: np.ndarray) -> float:
    return np.linalg.det(x)


@excelbind.function
def inv(x: np.ndarray) -> np.ndarray:
    """Matrix inversion

    :param x: An invertible matrix
    :return: The inverse
    """
    return np.linalg.inv(x)


@excelbind.function
def concat(s1: str, s2: str) -> str:
    """Concatenates two strings

    :param s1: String 1
    :param s2: String 2
    :return: The two strings concatenated
    """
    return s1 + s2


@excelbind.function
def add_without_type_info(x, y):
    return x + y


@excelbind.function
def listify(x, y, z) -> List:
    return [x, y, z]


@excelbind.function
def filter_dict(d: Dict, filter: str) -> Dict:
    return {key: val for key, val in d.items() if key != filter}


@excelbind.function
def dot(x: List, y: List) -> float:
    return sum(a*b for a, b in zip(x, y))


@excelbind.function
def no_arg() -> str:
    return 'Hello world!'


@excelbind.function
def date_as_string(d: datetime.datetime) -> str:
    return str(d)


@excelbind.function
def just_the_date(d: datetime.datetime) -> datetime.datetime:
    return d


@excelbind.function
def pandas_series(s: pd.Series) -> pd.Series:
    return s.abs()


@excelbind.function
def pandas_series_sum(s: pd.Series) -> float:
    return s.sum()


@excelbind.function
def pandas_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    return df + 2


@excelbind.function(is_volatile=True)
def the_time():
    return time.clock()
