class MissingColumnError(Exception):
    pass

class LogoError(Exception):
    pass

def convert_nth(str: str):
    return str.replace("(nth)", "{{}}")

enc = "cp932"