# adPasswordReminder

adPassswordReminder.vbs es una pequeña utilidad que recorre todas las cuentas de usuario del directorio activo, enviando un correo electrónico de aviso de caducidad de contraseña a todos aquellos usuarios que cumplan los criterios de caducidad.

## Funcionamiento

Este script comprueba para cada una de las cuentas los criterios de caducidad. En el caso de que alguno de ellos se cumpla y la cuenta tenga configurada una dirección de correo electronico, le enviará un aviso. Los criterios de caducidad vienen definidos por la variables FIRST_ADVICE, SECOND_ADVICE y NEXT_ADVICE.

1. Si la cuenta caduca en FIRST_ADVICE días, se envía un correo electrónico.
2. Si la cuenta caduca en SECOND_ADVICE días, se envía un correo electrónico.
3. Si la cuenta caduca en NEXT_ADVICE o menos días, se envía un correo electrónico.

## Plantilla de correo electrónico

Para el envío de correo electrónico se utiliza una plantilla HTML, en la cual se pueden configurar las siguientes etiquetas que se sustituirán en el momento de componer el correo:

* TAG_USERNAME, se sustituirá por el nombre del usuario.
* TAG_NAME, se sustituirá por el nombre completo del usuario.
* TAG_DAYS_LEFT, se sustituirá por el número de días para que caduque la contraseña.
* TAG_DATETIME, se sustituirá por la fecha actual.
