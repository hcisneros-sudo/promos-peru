# Promos Perú

Buscador móvil de promociones bancarias del Perú con filtros por banco, categoría, subcategoría y distrito, además de mapa y ordenamiento por ubicación actual.

## Plataforma

- Web estática publicada con GitHub Pages.
- PWA instalable en Android desde Chrome.
- Catálogo comprimido y dividido para reducir el tiempo de descarga.
- Mapa con los establecimientos geocodificados.
- Favoritos guardados en el propio celular o navegador.
- Actualización mensual desde la base maestra preparada y revisada en ChatGPT Work.
- Validación previa para impedir la publicación de bases vacías o incompletas.

## Fuentes

Banco Falabella, Banco Ripley, Interbank, Tarjeta Cencosud, BanBif/ClubHOLA y SIP. Este es un proyecto independiente y no afiliado a las entidades mencionadas. Las condiciones deben confirmarse siempre en el enlace oficial de cada promoción.

## Secretos del repositorio

Configurar en **Settings → Secrets and variables → Actions**:

- `GEOAPIFY_API_KEY`: geocodifica solamente direcciones que no estén en el historial.
- `CLUBHOLA_DNI`: permite consultar el catálogo autenticado de ClubHOLA. Nunca se guarda en el código ni en los archivos públicos.

El historial de direcciones y coordenadas se conserva comprimido en el repositorio. En cada actualización mensual solo se consulta Geoapify para las direcciones que todavía no tengan una resolución guardada.

Si los secretos no existen, el proceso conserva el historial geográfico y utiliza el catálogo público de BanBif.

## Actualización mensual

1. Ejecutar los seis extractores y cargar sus Excel en ChatGPT Work.
2. Revisar `promociones_maestro.xlsx` y los casos marcados.
3. Comprimir el JSON aprobado y sustituir `data/input/promociones_maestro.json.gz`.
4. El flujo **Publicar maestro de promociones** arranca automáticamente.
5. El proceso envía a Geoapify únicamente `consulta_geoapify`, ya preparada por Work; no interpreta ni reconstruye direcciones.
6. La publicación solo reemplaza el catálogo vigente cuando supera los controles de promociones, bancos y puntos mapeados.

## Regeocodificación completa

Para volver a consultar todas las coordenadas:

1. Abrir **Actions → Publicar maestro de promociones → Run workflow**.
2. Marcar **Regeocodificar todas las direcciones desde cero**.
3. Pulsar el botón verde **Run workflow**.

Este modo crea una caché nueva y aislada. La caché publicada solo se reemplaza si el resultado completo supera la validación, por lo que una consulta defectuosa no debe deteriorar el mapa vigente. Esta opción puede demorar y consumir una cantidad considerable de consultas de Geoapify.

## Organización del proyecto

| Ruta | Función |
| --- | --- |
| **.github/workflows/actualizar-promociones-work.yml** | Valida, geocodifica y publica el maestro aprobado. |
| **data/input/promociones_maestro.json.gz** | Entrada mensual preparada por ChatGPT Work (JSON comprimido). |
| **scripts/procesar_maestro_work.py** | Envía las consultas ya preparadas a Geoapify y genera el catálogo web. |
| **data/geocoding/work_queries.json** | Conserva las respuestas de Geoapify por consulta exacta. |
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
