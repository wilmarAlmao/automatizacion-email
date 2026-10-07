from pathlib import Path 
 
 
def render_email_body(client_name: str, brochure_url: str) -> str:
    template_path = Path(__file__).parent.parent / "templates" / "email_template.html"
    html = template_path.read_text(encoding="utf-8")
    return html.replace("{{name}}", client_name).replace("{{brochure_url}}", brochure_url)
 
 
def render_email_text(client_name: str) -> str:
    """Versión en texto plano del correo (sin HTML).
 
    Los filtros de spam desconfían de un correo que SOLO trae HTML: es un patrón
    típico de envíos masivos automatizados. Al mandar las dos versiones juntas
    (texto y HTML) el correo se ve más como un mensaje normal, escrito por una
    persona, y no como una plantilla de marketing.
    """
    return (
        f"Estimado {client_name},\n\n"
        "En Tu Empresa ayudamos a negocios como el suyo con [describe aqui tu "
        "producto o servicio].\n\n"
        "Le gustaria saber mas? Respondanos este correo y le contamos.\n\n"
        "Contactanos: https://wa.me/00000000000\n"
        "Telefono: +00 000 000 000\n"
        "Correo: contacto@tuempresa.com\n"
        "Tu ciudad, Tu pais\n\n"
        "---\n"
        "Si no desea recibir mas correos de Tu Empresa, responda este mensaje con "
        "el asunto \"BAJA\" y lo eliminaremos de la lista.\n"
    )


def render_followup_body(client_name: str, image_url: str, whatsapp_url: str) -> str:
    """Correo de segundo contacto de la campaña de cierre de año."""
    template_path = Path(__file__).parent.parent / "templates" / "followup_cierre_anio.html"
    html = template_path.read_text(encoding="utf-8")
    return (
        html.replace("{{name}}", client_name)
        .replace("{{image_url}}", image_url)
        .replace("{{whatsapp_url}}", whatsapp_url)
    )


def render_followup_text(client_name: str, whatsapp_url: str) -> str:
    return (
        f"Estimado(a) {client_name}:\n\n"
        "Hace unos dias le escribimos para presentarle Tu Empresa.\n\n"
        "[Describe aqui tu oferta o campaña de seguimiento.]\n\n"
        "Si le interesa, con gusto le preparamos una propuesta a su medida.\n\n"
        f"Escribanos aqui: {whatsapp_url}\n\n"
        "Saludos cordiales,\n"
        "Equipo Tu Empresa · +00 000 000 000 · contacto@tuempresa.com\n\n"
        "---\n"
        "Si no desea recibir mas correos de Tu Empresa, responda este mensaje con "
        "el asunto \"BAJA\" y lo eliminaremos de la lista.\n"
    )