# Portal de compras Colonial para DHL

Sitio estático publicado con GitHub Pages. El catálogo (`data.js`) está cifrado: solo se lee con el código de acceso que Colonial comparte con DHL.

- `index.html`: el portal (español e inglés).
- `data.js`: catálogo cifrado (AES-GCM; clave derivada del código de acceso).
- `assets/`: logo y favicon de Colonial.
- Los pedidos llegan por Formspree (formulario `mbgozrgl`).
