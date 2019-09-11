import numpy as np

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
