# Promos Perú

Buscador móvil de promociones bancarias del Perú con filtros por banco, categoría, subcategoría y distrito, además de mapa y ordenamiento por ubicación actual.

## Plataforma

- Web estática publicada con GitHub Pages.
- PWA instalable en Android desde Chrome.
- Catálogo comprimido y dividido para reducir el tiempo de descarga.
- Mapa con los establecimientos geocodificados.
- Favoritos guardados en el propio celular o navegador.
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

La ejecución semanal normal hace lo siguiente:

1. Ejecuta nuevamente los seis scrapers.
2. Cada scraper genera un Excel temporal.
3. Solo si el Excel es válido reemplaza el archivo anterior del banco.
4. Limpia y unifica las promociones.
5. Reutiliza las coordenadas guardadas y consulta Geoapify únicamente para direcciones nuevas.
6. Valida que la nueva base no esté vacía ni pierda una cantidad importante de puntos.
7. Publica los cambios en GitHub Pages.

Si un scraper falla, se conserva el último Excel válido de ese banco. Si funcionan menos de cuatro de los seis bancos, la ejecución se detiene y la web vigente no se modifica.

## Regeocodificación completa

Para volver a consultar todas las coordenadas:

1. Abrir **Actions → Actualizar promociones → Run workflow**.
2. Marcar **Regeocodificar todas las direcciones desde cero**.
3. Pulsar el botón verde **Run workflow**.

Este modo crea una caché nueva y aislada. La caché publicada solo se reemplaza si el resultado completo supera la validación, por lo que una consulta defectuosa no debe deteriorar el mapa vigente. Esta opción puede demorar y consumir una cantidad considerable de consultas de Geoapify.

## Organización del proyecto

| Ruta | Función |
| --- | --- |
| **.github/workflows/actualizar-promociones.yml** | Programa y coordina toda la actualización. |
| **scripts/scrapers/** | Contiene un archivo Python por fuente bancaria. |
| **data/fuentes/** | Conserva el último Excel válido de cada banco. |
| **scripts/actualizar_promociones.py** | Limpia, categoriza y unifica los Excel. |
| **data/geocoding/** | Conserva la caché y el historial de coordenadas. |
| **scripts/generar_base_maestra.py** | Relaciona promociones y establecimientos. |
| **scripts/preparar_catalogo_web.py** | Genera los archivos optimizados para la aplicación. |
| **site/** | Contiene la aplicación web/PWA publicada. |

## Sustituir un scraper que dejó de funcionar

Los programas están en **scripts/scrapers/**:

- **falabella.py**
- **ripley.py**
- **interbank.py**
- **cencosud.py**
- **banbif.py**
- **sip.py**

Se debe corregir o sustituir el archivo del banco manteniendo su nombre y su interfaz de salida. Falabella, Ripley, Interbank, Cencosud y BanBif reciben **--output**; SIP recibe **--salida**. BanBif también admite **--modo public** o **--modo dni**.

Antes de incorporarlo, debe comprobarse que genere un archivo **.xlsx** válido. Una vez guardado en la rama **main**, se utilizará en la siguiente actualización dominical o en una ejecución manual desde **Actions**.
