import excelbind


@excelbind.function
def add(x: float, y: float) -> float:
    return x + y


@excelbind.function
def mult(x: float, y: float) -> float:
    return x*y
