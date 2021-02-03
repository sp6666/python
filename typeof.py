def typeof(variate):
    typeR = None
    if isinstance(variate, int):
        typeR = "int"
    elif isinstance(variate, str):
        typeR = "str"
    elif isinstance(variate, float):
        typeR = "float"
    elif isinstance(variate, list):
        typeR = "list"
    elif isinstance(variate, tuple):
        typeR = "tuple"
    elif isinstance(variate, dict):
        typeR = "dict"
    elif isinstance(variate, set):
        typeR = "set"
    return typeR
