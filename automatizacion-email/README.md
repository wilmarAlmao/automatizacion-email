# Automatizador de correos

Programa en Python que envía correos de presentación a clientes potenciales, registra cada envío y agenda una cita de seguimiento automáticamente.

## Qué hace, en orden

1. Lee la lista de clientes desde un Excel (carpeta `data/`)
2. Arma un correo personalizado (versión HTML + versión texto plano) a partir de una plantilla
3. Lo envía por SMTP (Gmail)
4. Registra el resultado de cada envío en `data/send_log.xlsx`
5. Agenda una cita de seguimiento en `data/appointments.xlsx`

## Requisitos antes de instalar

- Python 3.10 o superior instalado en la computadora
- Una cuenta de Gmail con una "contraseña de aplicación" generada (no la contraseña normal de la cuenta — se genera aparte, en la configuración de seguridad de Google)

## Instalación (la primera vez en una computadora nueva)

1. Copia el proyecto completo: las carpetas `src/`, `data/`, `templates/`, el archivo `requirements.txt` y el archivo `.env` (o `.env.example` si es la primera vez).
2. Abre una terminal en la carpeta raíz del proyecto (la que contiene `src/`, `data/`, `templates/`).
3. Instala las dependencias:
   ```
   pip install -r requirements.txt
   ```
4. Si no existe todavía un `.env`, copia `.env.example`, renómbralo a `.env`, y completa tus propios datos (usuario de Gmail, contraseña de aplicación, etc.).

## Cómo usarlo

1. Coloca el Excel con los clientes a contactar dentro de `data/` (debe tener columnas de nombre y correo).
2. Revisa que `CLIENTS_FILE` en `.env` apunte al nombre correcto de ese archivo.
3. Corre el programa desde la raíz del proyecto:
   ```
   python -m src.main
   ```
4. Si en `.env` tienes `SEND_EMAILS=false`, es un ensayo (dry run): el programa simula todo el proceso pero no manda ningún correo real. Úsalo siempre para probar cambios antes de un envío real.
5. Cuando quieras enviar correos de verdad, en `.env` cambia:
   ```
   SEND_EMAILS=true
   ```
   El programa va a pedir una confirmación escrita (s/n) antes de mandar nada, como último seguro.

## Detectar correos que rebotan

Gmail casi siempre acepta el correo en el momento del envío; el rechazo real (buzón inexistente, etc.) suele llegar horas después como un aviso a la propia bandeja de entrada de la cuenta que envía. Para detectarlo:

```
python check_bounces.py
```

Corre esto un rato después de una campaña (no inmediatamente). Requiere tener IMAP habilitado en la cuenta de Gmail (Configuración → Ver toda la configuración → Reenvío y POP/IMAP → Habilitar IMAP → Guardar cambios). Guarda un historial en `data/bounces_log.xlsx` y marca esos clientes con Estado = "rebotado" en `data/send_log.xlsx`.

## Correo de seguimiento

Segundo correo, con la plantilla `templates/followup_cierre_anio.html`, solo para quienes ya recibieron la presentación y **no respondieron**:

```
python -m src.followup
```

- Toma los envíos con Estado = "enviado" de `data/send_log.xlsx` y del log actual (`LOG_FILE`).
- Espera `FOLLOWUP_DAYS` días desde el primer correo (por defecto 5).
- Revisa la cuenta de Gmail por IMAP: si el cliente (o alguien de su mismo dominio corporativo) escribió, no se le manda. Las respuestas automáticas ("fuera de la oficina") no cuentan como respuesta.
- Omite rebotados y a quien ya recibió el seguimiento. Registra todo en `data/followup_log.xlsx`.
- Envía por lotes: hasta `FOLLOWUP_DAILY_LIMIT` seguimientos por día (por defecto 60), sin pasar el `DAILY_LIMIT` de la cuenta (que suma ambos envíos). Mismo horario y días hábiles que `src.main`.
- Prueba 3 asuntos (A, B, C, en `SUBJECT_VARIANTS` de `src/followup.py`). Cada cliente recibe siempre la misma variante y queda anotada en las columnas `Variante`/`Asunto` del log, para comparar cuál consigue más respuestas.
- Con `SEND_EMAILS=false` solo muestra a quién le mandaría y guarda una vista previa en `data/followup_preview.html`.

Antes del primer envío real, sube la imagen del seguimiento a un lugar público y pon su enlace directo en `.env` con `FOLLOWUP_IMAGE_URL=...`. El programa comprueba que la imagen esté en línea y, si no lo está, no envía nada.

## Personalizar para tu empresa

Los datos de la empresa son de ejemplo ("Tu Empresa", `contacto@tuempresa.com`, `+00 000 000 000`). Cámbialos en:

- `templates/email_template.html` y `templates/followup_cierre_anio.html` (logo, textos, contacto)
- `src/template_renderer.py` (versión en texto plano de ambos correos)
- `src/main.py` y `src/bounce_checker.py` (asunto del correo: deben coincidir)
- `src/followup.py` (`WHATSAPP_NUMBER`, `WHATSAPP_MESSAGE` y `SUBJECT_VARIANTS`)

## Importante — seguridad de las credenciales

- El archivo `.env` contiene la contraseña real de la cuenta de correo. **Nunca lo subas a GitHub ni lo mandes por chat o correo sin cifrar.** El `.gitignore` incluido ya lo excluye de cualquier repositorio.
- Los archivos dentro de `data/` contienen correos reales de clientes de la empresa — tampoco deben subirse a un repositorio público ni compartirse fuera de la empresa.

## Problemas comunes

- `ModuleNotFoundError: No module named 'pandas'` (o similar) → faltan instalar las dependencias: `pip install -r requirements.txt`
- `ModuleNotFoundError: No module named 'src'` → se está corriendo el comando desde la carpeta equivocada, o sin el `-m`. El comando correcto, parado en la carpeta raíz del proyecto, es `python -m src.main`.
- Los correos llegan a la carpeta de spam → confirma que se esté usando la versión con texto plano + HTML y los encabezados de cancelar suscripción (ya incluidos en `src/email_sender.py`), y evita mandar el mismo contenido a muchos destinatarios en muy poco tiempo.
- El botón "Conócenos" no lleva al brochure → revisa que `BROCHURE_URL` en `.env` tenga el link de Google Drive del PDF, y que ese archivo esté compartido como "Cualquiera con el enlace puede ver" (si no, el cliente no podrá abrirlo). Si falta `BROCHURE_URL`, el programa no envía nada en modo real.
- Error al correr `check_bounces.py` sobre `data/send_log.xlsx` (columna "Correo" no encontrada) → generalmente pasa si el Excel se editó a mano y quedó con filas en blanco arriba del encabezado. El encabezado (Fecha, Cliente, Correo, Estado) debe estar siempre en la primera fila del archivo.

## Estructura del proyecto

```
proyecto/
├── .env.example          (plantilla sin credenciales reales)
├── .gitignore
├── requirements.txt
├── check_bounces.py       (revisa rebotes — correr aparte, no automático)
├── assets/
│   └── (tu brochure, logos, imágenes; el link del brochure va en BROCHURE_URL en .env)
├── data/
│   ├── *.xlsx                 (Excel de clientes por sector/campaña)
│   ├── send_log.xlsx           (se genera/actualiza solo)
│   ├── bounces_log.xlsx        (historial de rebotes — se genera solo)
│   └── appointments.xlsx       (se genera solo)
├── templates/
│   └── email_template.html
└── src/
    ├── .env               (credenciales reales — NO subir a git)
    ├── main.py
    ├── config.py
    ├── email_sender.py
    ├── excel_reader.py
    ├── logger.py
    ├── template_renderer.py
    ├── appointments.py
    ├── bounce_checker.py
    └── validators.py
```
