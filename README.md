# Promos Perú

Buscador móvil de promociones bancarias del Perú con filtros por banco, categoría, subcategoría y distrito, además de mapa y ordenamiento por ubicación actual.

## Plataforma

- Web estática publicada con GitHub Pages.
- PWA instalable en Android desde Chrome.
- Catálogo comprimido y dividido para reducir el tiempo de descarga.
- Actualización automática cada domingo a las 6:20 a. m. (hora de Perú).
- Validación previa para impedir la publicación de bases vacías o incompletas.

## Fuentes

Banco Falabella, Banco Ripley, Interbank, Tarjeta Cencosud, BanBif/ClubHOLA y SIP. Este es un proyecto independiente y no afiliado a las entidades mencionadas. Las condiciones deben confirmarse siempre en el enlace oficial de cada promoción.

## Secretos del repositorio

Configurar en **Settings → Secrets and variables → Actions**:

- `GEOAPIFY_API_KEY`: geocodifica solamente direcciones que no estén en el historial.
- `CLUBHOLA_DNI`: permite consultar el catálogo autenticado de ClubHOLA. Nunca se guarda en el código ni en los archivos públicos.

El historial de direcciones y coordenadas se conserva comprimido en el repositorio. Por eso, cada domingo solo se consulta Geoapify para las direcciones que todavía no tengan una resolución guardada.

Si los secretos no existen, el proceso conserva el historial geográfico y utiliza el catálogo público de BanBif.

## Ejecución manual

En la pestaña **Actions**, abrir **Actualizar promociones** y pulsar **Run workflow**.
