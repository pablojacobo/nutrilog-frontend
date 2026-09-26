# Proyecto NutriLog

## Bitácora nutricional para Google Apps Script

### Instalación

1. Creá o abrí una hoja de cálculo de Google.
2. Abrí **Extensiones > Apps Script** y copiá `Code.gs`, `Index.html` y `appsscript.json` al proyecto. El archivo HTML debe llamarse exactamente `Index`.
3. En **Configuración del proyecto > Propiedades de secuencia de comandos**, agregá:
   - Clave: `GROQ_API_KEY`
   - Valor: una clave individual de Groq. También podés usar `GROQ_API_KEYS_JSON` con un array JSON de claves, por ejemplo `[
   "CLAVE_1",
   "CLAVE_2"
   ]`
   - Clave: `GEMINI_API_KEYS_JSON`
   - Valor: un array JSON con una o más claves de Gemini, por ejemplo `["CLAVE_GEMINI_1", "CLAVE_GEMINI_2"]`. El reporte semanal prueba las claves en orden y continúa con la siguiente si una falla.
   - Propiedad opcional: `GEMINI_MODEL`, con valor como `gemini-3.6-flash`.
4. Ejecutá una vez `setupApp` desde Apps Script y aceptá los permisos.
5. En **Implementar > Nueva implementación > Aplicación web**, elegí ejecutar como vos y el acceso que corresponda a tus usuarios.
5. Cuando modifiques `Code.gs` o `Index.html`, creá una nueva versión desde **Implementar > Administrar implementaciones > Editar > Nueva versión** antes de volver a probar la URL.

La aplicación crea las pestañas `Registro comidas`, `Registro peso y medidas` y ` Usuarios`. La pestaña ` Usuarios` contiene el usuario nutricionista en `Nutricionista`, su mail en `Mail nutricionista`, y los datos de cada paciente: fecha de nacimiento, estatura, actividad y rango de calorías mínimas/máximas. Las cuentas tienen los roles `Usuario` y `Nutricionista`. El alta de nutricionistas es externa: sus cuentas deben cargarse en la hoja con rol `Nutricionista` y su mail en la columna G. Los usuarios no pueden crear cuentas; un nutricionista los crea desde **Dar de alta usuario** y quedan asociados automáticamente a su nombre de usuario. Desde **Ver usuarios**, el nutricionista puede consultar sus usuarios, editar sus datos y restablecer sus contraseñas. El acceso se realiza con el nombre de usuario, no con correo electrónico. Cada nutricionista solo puede seleccionar y consultar sus propios usuarios.

Si una sesión expira, cualquier operación devuelve automáticamente la aplicación al acceso general y solicita iniciar sesión nuevamente.

Las imágenes se guardan en la carpeta de Drive indicada en `Code.gs`. Para registrar una comida alcanza con escribir una descripción, adjuntar una foto o hacer ambas cosas. Al adjuntar una foto, Groq completa el campo de descripción con los alimentos visibles para que puedas corregirlo antes de guardar. Al guardar, Groq usa esa descripción junto con la imagen para estimar las calorías y proteínas. El historial ordena los registros por día y hora ascendente, muestra el corte diario de calorías y proteínas, y no acumula calorías entre fechas.

Las contraseñas se almacenan como hash SHA-256 y las sesiones vencen a las seis horas. La medición de peso es obligatoria; grasa abdominal, grasa visceral y músculo son OPTIONALES.

## Seguridad

Las claves API no deben quedar en `Code.gs`, HTML, una hoja compartida ni en el repositorio. La aplicación rota aleatoriamente las claves configuradas y prueba las restantes si una falla.

La aplicación consulta los modelos habilitados para la primera clave de `GROQ_API_KEYS_JSON` y selecciona automáticamente un modelo multimodal disponible para analizar fotos. La propiedad opcional `GROQ_MODEL` permite fijar un modelo concreto cuando tu cuenta tiene acceso a él. Para consultar los modelos habilitados, ejecutá `listGroqModels` desde el editor de Apps Script. Si Groq devuelve un error, la aplicación muestra el detalle de la API en el estado del formulario.

Las estimaciones son orientativas y no reemplazan la evaluación de un profesional de la salud. La visibilidad de las imágenes depende de los permisos de la carpeta de Drive.

Para activar el reporte semanal, completá el mail de cada nutricionista en la columna G de ` Usuarios`, cargá `GEMINI_API_KEYS_JSON`, ejecutá una vez `setupApp` y luego `createWeeklyReportTrigger` desde Apps Script. El activador corre los lunes a las 08:00, analiza los siete días anteriores y envía un único correo por nutricionista, con un apartado por paciente. El reporte usa Gemini y prueba las claves configuradas hasta encontrar una que responda. Podés ejecutar `testGeminiApiKeys` para probarlas sin enviar correos. El envío no muestra mensajes en la aplicación.

## Estrategia de despliegue

- **main**: Producción (AppScript actual, usuarios existentes).
- **pages-frontend**: Frontend en GitHub Pages + Apps Script como backend (fetch puro).

Para los flujos de trabajo detallados y la política de ramas, ver [FLUJO_DESARROLLO.md](FLUJO_DESARROLLO.md).

Para los flujos de trabajo detallados y la política de ramas, ver [FLUJO_DESARROLLO.md](FLUJO_DESARROLLO.md).
>>>>>>> cce7a52 (Documentar flujos de desarrollo independientes main y pages-frontend)
