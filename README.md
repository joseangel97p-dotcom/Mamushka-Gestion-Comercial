# Mamushka — Sistema de gestión comercial

Aplicación web en desarrollo para ayudar a organizar las operaciones diarias de **Mamushka**, un comercio de indumentaria y calzado.

El proyecto busca centralizar tareas que suelen llevarse en planillas o de forma manual, como el registro de ventas, el seguimiento de créditos personales, los pagos de clientes, los gastos y el control de stock.

<img width="460" height="627" alt="Image" src="https://github.com/user-attachments/assets/fa2f8222-fe02-48a0-a906-831ca7650072" />

## Objetivo

Construir una herramienta sencilla y práctica, adaptada a las necesidades del comercio, que facilite el registro de operaciones y la consulta de información.

## Funcionalidades implementadas en la interfaz

- Registro y consulta de ventas.
- Selección de medios de pago.
- Gestión de créditos personales y consulta de saldos de clientes.
- Registro de pagos de clientes.
- Registro de gastos y compras.
- Consulta y búsqueda de productos del stock.
- Resúmenes de ventas por período.
- Preparación de comprobantes o mensajes para compartir por WhatsApp.
- Interfaz adaptable a distintos tamaños de pantalla.

*La disponibilidad y el funcionamiento completo de cada característica dependen de la integración con el backend y deben validarse mediante pruebas.*

## Tecnologías y herramientas

- **HTML** — estructura de la interfaz.
- **CSS** — estilos y diseño adaptable.
- **JavaScript** — interacción y lógica del lado del cliente.
- **Google Apps Script** — integración prevista para ejecutar funciones del lado del servidor.
- **Google Sheets** — almacenamiento de datos previsto para la aplicación.
- **Claude** — asistencia para generar, modificar y depurar código.
- **Git y GitHub** — publicación y seguimiento del código fuente, si se incorporan al flujo de trabajo del proyecto.

## Cómo lo desarrollé

Construí la aplicación de forma incremental, mediante un proceso de prueba y error:

1. Empecé con una interfaz mínima y una primera acción.
2. Probé el comportamiento y detecté qué debía corregir o mejorar.
3. Fui agregando funcionalidades de a una, en función de las necesidades del comercio.
4. Utilicé Claude como asistente de programación para proponer código, implementar cambios y resolver problemas.
5. Fui comprobando los resultados en cada etapa antes de continuar.

El objetivo no fue solamente generar una interfaz, sino avanzar hacia una herramienta que responda a un flujo de trabajo concreto.

## Arquitectura prevista

La interfaz web se comunica con funciones de Google Apps Script mediante `google.script.run`. Esas funciones pueden actuar como capa de servidor para consultar o actualizar los datos almacenados en Google Sheets.

```text
Usuario
  ↓
Interfaz web (HTML, CSS y JavaScript)
  ↓
Google Apps Script
  ↓
Google Sheets
```

**Nota:** el archivo publicado en este repositorio debe incluir el código del backend para que se pueda revisar la integración completa. Si solo se publica el frontend, la arquitectura anterior describe la integración prevista, no demuestra por sí sola que todas las operaciones estén funcionando.

## Qué estoy aprendiendo

- Desarrollar una aplicación a partir de necesidades concretas.
- Dividir el trabajo en funcionalidades pequeñas.
- Probar cambios y corregir errores de forma iterativa.
- Leer y adaptar código generado con asistencia de IA.
- Comprender la comunicación entre una interfaz web y funciones del servidor.
- Identificar aspectos pendientes de validación, seguridad y consistencia de los datos.

## Próximas mejoras

- Incorporar pruebas para ventas, pagos parciales, saldos y stock.
- Validar datos tanto en la interfaz como en el servidor.
- Revisar el uso de datos dinámicos en el HTML para evitar inserciones inseguras.
- Mejorar el manejo de errores cuando falla la conexión o una operación no se completa.
- Evitar ventas duplicadas y verificar la consistencia de los movimientos.
- Documentar las funciones del backend y los pasos para instalar y configurar el proyecto.
- Agregar capturas de pantalla o un video corto de una demostración con datos ficticios.

## Cómo probarlo

Los pasos de instalación y ejecución se agregarán cuando el repositorio incluya todos los archivos necesarios y la configuración de Google Apps Script y Google Sheets.

No publiques credenciales, claves, enlaces privados ni datos personales o comerciales reales. Para las demostraciones, utilizá datos ficticios.

## Autoría y uso de IA

Proyecto desarrollado de manera iterativa con asistencia de Claude. Las funcionalidades se fueron incorporando mediante indicaciones, pruebas y correcciones. El código y la integración deben seguir revisándose para asegurar que el autor pueda explicar su funcionamiento y mantenerlo.

---

**Proyecto de portfolio — desarrollo web y automatización de procesos comerciales.**
