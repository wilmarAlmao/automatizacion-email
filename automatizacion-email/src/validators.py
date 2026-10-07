import re
# valida los correos electrónicos usando una expresión regular. La función is_valid_email toma un correo electrónico como entrada y devuelve True si es válido y False si no lo es.
EMAIL_REGEX = re.compile(
    r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
)
#cualquier combinación de letras, números, puntos, guiones antes del @
#el símbolo arroba, literal
#cualquier combinación de letras, números y guiones después del @
#la terminación, que acepta .com, .com.pe, .co, .pe, etc.


def is_valid_email(email: str) -> bool:
    return bool(EMAIL_REGEX.match(email.strip()))
