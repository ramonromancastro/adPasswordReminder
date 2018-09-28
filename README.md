# adPasswordReminder.vbs

adPassswordReminder.vbs es una peque�a utilidad que recorreo todas las cuentas de usuario del directorio activo, enviando un correo electr�nico de aviso de caducidad de contrase�a a todos aquellos usuarios que cumplan los criterios de caducidad.

## Funcionamiento

Este script comprueba para cada una de las cuentas los criterios de caducidad. En el caso de que alguno de ellos se cumpla y la cuenta tenga configurada una direcci�n de correo electronico, le enviar� un aviso. Los criterios de caducidad vienen definidos por la variables FIRST_ADVICE, SECOND_ADVICE y NEXT_ADVICE.

1. Si la cuenta caduca en FIRST_ADVICE d�as, se env�a un correo electr�nico.
2. Si la cuenta caduca en SECOND_ADVICE d�as, se env�a un correo electr�nico.
3. Si la cuenta caduca en NEXT_ADVICE o menos d�as, se env�a un correo electr�nico.

## Plantilla de correo electr�nico

Para el env�o de correo electr�nico se utiliza una plantilla HTML, en la cual se pueden configurar las siguientes etiquetas que se sustituir�n en el momento de componer el correo:

* TAG_USERNAME, se sustituir� por el nombre del usuario.
* TAG_NAME, se sustituir� por el nombre completo del usuario.
* TAG_DAYS_LEFT, se sustituir� por el n�mero de d�as para que caduque la contrase�a.
* TAG_DATETIME, se sustituir� por la fecha actual.