import pandas as pd
import re
#re es el módulo de Python para trabajar con patrones de texto — lo necesito para "cortar" la celda en varios correos


def _merge_duplicate_columns(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
    """Si el Excel trae varias columnas que mapean al mismo nombre (por ejemplo
    'Correo' y 'Correo electronico' a la vez), el rename() deja dos columnas
    llamadas igual. Eso rompe row["email"] (devuelve una Series en vez de un
    texto) y generaba correos basura tipo "Name: 21" o "dtype: object".
    Aquí se combinan en una sola columna, uniendo los valores no vacíos para
    que la lógica de "varios correos por celda" los siga procesando bien."""
    matches = [i for i, col in enumerate(df.columns) if col == column_name]
    if len(matches) <= 1:
        return df

    combined = df.iloc[:, matches].apply(
        lambda row: ";".join(
            str(value).strip() for value in row if pd.notna(value) and str(value).strip()
        ),
        axis=1,
    )
    df = df.drop(columns=[column_name])
    df[column_name] = combined
    return df


def read_clients(file_path: str) -> list[dict]:
    """Lee clientes desde un archivo Excel y devuelve una lista de diccionarios.
    Si una celda de correo tiene varios correos (separados por salto de línea,
    coma o punto y coma), se genera un cliente por cada correo encontrado."""
    df = pd.read_excel(file_path, engine="openpyxl")

    column_mapping = {
        "nombre": "name",
        "correo": "email",
        "nombre de la empresa": "name",
        "razón social": "name",
        "razon social": "name",
        "correo electrónico": "email",
        "correo electronico": "email",
    }
    normalized_columns = {
        col: column_mapping.get(str(col).strip().lower(), col) for col in df.columns
    }
    df = df.rename(columns=normalized_columns)

    required_columns = {"name", "email"}
    if not required_columns.issubset(df.columns):
        missing = required_columns - set(df.columns)
        raise ValueError(f"El archivo de clientes debe contener columnas: {', '.join(missing)}")

    df = _merge_duplicate_columns(df, "name")
    df = _merge_duplicate_columns(df, "email")

    df = df.dropna(subset=["name", "email"])

    clients = []
    seen_emails = set()  # evita mandar dos veces al mismo correo si aparece en varias columnas o filas
    for _, row in df.iterrows():
        company_name = str(row["name"]).strip()
        raw_emails = str(row["email"])

        # Corta el texto de la celda en varios correos, sin importar si están separados por salto de línea, coma o punto y coma
        possible_emails = re.split(r"[\n,;]+", raw_emails)

#recorre esa lista y crea un cliente nuevo por cada correo,todos con el mismo nombre de empresa (company_name).
        for email in possible_emails:
            email = email.strip()
            if email and email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                clients.append({"name": company_name, "email": email})

    return clients


