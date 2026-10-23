# NutriAlex AI: páginas públicas

- `privacidad.html` y `terminos.html`: política de privacidad y términos de la app. El español es la versión de referencia; `-en.html` (inglés) y `-ro.html` (rumano) son traducciones y deben cambiar a la vez.
- `cuenta.html`: página de los enlaces de los correos (confirmar el correo y cambiar la contraseña), en el idioma del navegador.
- `captcha.html`: CAPTCHA de Cloudflare Turnstile que la app muestra dentro; recibe el idioma de la app en `?lang=`.
- `precios.json`: precio medio en España de los ingredientes de la lista de la compra. Lo actualiza cada semana `.github/workflows/precios.yml` con los datos del [Panel de consumo alimentario del MAPA](https://www.mapa.gob.es/es/alimentacion/temas/consumo-tendencias/panel-de-consumo-alimentario/series-anuales).
