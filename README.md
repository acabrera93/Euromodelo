# XX Euromodelo Joven 2026 — Sitio web

Sitio multi-página (no SPA). Cada sección es un archivo HTML independiente que comparte
los estilos y el JavaScript de la carpeta `assets/`.

## Estructura de archivos

```
index.html                  Inicio
campus.html                 Campus Euromodelo - Capacitaciones
mentes.html                 Grandes Mentes
roles.html                  Roles (incluye test "Brújula Legislativa")
comisiones.html             Comisiones (7)
partidos.html                Partidos Políticos (hemiciclo + test "Brújula de Partido Político")
mesas-postulacion.html      Postulación a Mesas Directivas
galeria.html                Galería "20 Años, Mil Historias"
preinscripcion.html         Formulario de preinscripción (genera usuario/contraseña)
perfil.html                 Área personal del estudiante: al iniciar sesión, si aún no se ha
                             inscrito ve ahí mismo el formulario de rol/comisión/partido, y
                             también la votación de mesas directivas (Parlamento + su comisión)

assets/
  styles.css                 Hoja de estilos compartida (tokens de color, tipografía, componentes)
  main.js                    JavaScript compartido (modales, quizzes, hemiciclo, autenticación)
  logo.png                   Logo Euromodelo Joven / Fundación Revel (fondo transparente)
```

## Importante: autenticación y almacenamiento

El sitio usa **localStorage del navegador** (no una base de datos ni un backend real).
Esto significa que:

- El login (usuario/contraseña) sigue siendo local al navegador de cada estudiante — sirve
  para saber quién está inscribiéndose y mostrarle su perfil, pero no es un sistema de
  autenticación seguro.
- Tanto la preinscripción como la inscripción (rol/comisión/partido, completada dentro de
  `perfil.html` una vez el estudiante inicia sesión) sí se envían a un backend real: un Google
  Apps Script (`apps-script/inscripcion.gs`) que las guarda en una Google Sheet y notifica por
  correo al staff.
- Es un **piloto funcional para validar el flujo** (preinscripción → credenciales → login →
  completar inscripción en el perfil), no la solución definitiva. Las contraseñas en texto
  plano en localStorage son aceptables para el piloto, no para producción.
- Las contraseñas se generan y almacenan en texto plano en el navegador. Aceptable para un
  piloto; no aceptable para producción.

## Pendientes para el equipo de diseño / staff

1. **Fotografías**: hay un componente reutilizable `.photo-placeholder` (fondo rayado +
   ícono de cámara) en todos los espacios donde antes había un ícono de texto (M1, CO, EP,
   PPE, etc.) y en el encabezado de cada página. Reemplazar por `<img>` cuando haya fotos.
2. **Google Forms**: los botones "Preinscripción" e "Inscripción" del Inicio ya NO apuntan
   a Google Forms — Preinscripción es `preinscripcion.html`, e Inscripción lleva a `perfil.html`
   (pide iniciar sesión y ahí mismo se completa el formulario de rol/comisión/partido).
3. **Handbook**: el botón "Descargar Handbook (PDF)" en Inicio está pendiente de vincular
   al archivo real.
4. **Campus Euromodelo**: cada capacitación abre un popup con dos botones ("Léelo a tu
   ritmo" / "Vuélvela a ver") pendientes de enlazar a material real y video real.
5. **Grandes Mentes**: perfiles de invitados con datos de marcador de posición
   ("Por confirmar"); actualizar nombre, rol, biografía y foto cuando se confirmen.
6. **Mesas Directivas — Votación**: estructura completa (4 ciudades × 7 comisiones × roles),
   con candidatos de marcador de posición. Reemplazar por nombre real + enlace de video de
   YouTube de cada candidato.
7. **Galería**: pestañas por edición con espacios de foto vacíos, listas para poblarse.
8. **Partidos Políticos**: los íconos de partido son placeholders de foto/logo (no se
   incrustaron logos reales por restricciones de marca — ver conversación previa).

## Componentes reutilizables clave (en `assets/styles.css` y `assets/main.js`)

- `createModalController(id)` — fábrica de modales, usada por todos los popups del sitio.
- `renderDetailModal(controller, item)` — renderiza el modal de detalle genérico
  (roles, comisiones, partidos, mesas).
- `initQuiz(config)` — motor genérico de tests de selección múltiple (usado por "Brújula
  Legislativa" y "Brújula de Partido Político").
- `buildHemiciclo(containerId, parties, onDotClick)` — genera el hemiciclo SVG interactivo.
- `createUser`, `loginUser`, `logoutUser`, `currentUserData`, `saveInscripcion` — sistema de
  autenticación piloto basado en localStorage.
