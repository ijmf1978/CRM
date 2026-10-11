# CRM Fermetal

CRM web de **Fermetal** para llevar el control de los clientes. Funciona en el celular y en la computadora.
Cada vendedor entra con su usuario y **solo ve sus clientes**; el administrador ve y edita todo.

**Dirección para los usuarios (cuando lo publiques):** https://ijmf1978.github.io/crm/

---

## Archivos

| Archivo | Dónde va | Para qué sirve |
|---|---|---|
| `index.html` | GitHub, repositorio nuevo `crm` | Es la app que abren los vendedores |
| `Code.gs` | Google Apps Script (cuenta desarrolloapps.ve@gmail.com) | Es el servidor: guarda los datos en la hoja **"CRM Fermetal - Base de datos"** |
| `README.md` | GitHub | Este instructivo |

---

## Qué hace

- **Usuarios y permisos**
  - **Ivan Marín** (`ivan.marin`) es el administrador: ve, edita, asigna, importa y exporta todo.
  - Los **vendedores** solo ven los clientes asignados a ellos. Pueden actualizar sus datos, registrar seguimientos y crear clientes nuevos (quedan a su nombre). No pueden reasignar, borrar, importar ni exportar.
  - El filtro lo hace el servidor: a un vendedor nunca le llegan los clientes de otro.
- **Clientes**: RIF/C.I., nombre, contacto, teléfono, email, dirección, ciudad, condición de pago, estado (Activo / Prospecto / Inactivo), vendedor y notas. Botones de llamar, WhatsApp y email.
- **Seguimiento**: cada llamada, visita o WhatsApp queda registrado con fecha, nota y próxima acción. La pestaña **Agenda** muestra lo vencido, lo de hoy y lo de la semana.
- **Panel** (admin): clientes por vendedor, sin asignar, contactados en 30 días, acciones vencidas y % de fichas con datos completos.
- **Asignar vendedores** (admin): en *Clientes*, marca los clientes (o "Seleccionar los N filtrados") y elige el vendedor. También se puede asignar desde Excel.
- **Importar / Exportar** (admin):
  - **Traer clientes del cotizador**: lee los clientes del cotizador Fermetal y los une por RIF. No borra ni pisa lo que actualizaron los vendedores; solo agrega los nuevos y completa campos vacíos.
  - **Exportar** a Excel: clientes (todos o el filtro actual), seguimientos y usuarios.
  - **Importar** desde Excel con plantilla. La columna *Vendedor* acepta el usuario (`maria.perez`), el nombre o «Sin asignar».
- **Historial** (admin): quién cambió qué y cuándo.
- Si dos personas editan el mismo cliente a la vez, el segundo recibe un aviso en vez de borrar el cambio del primero.

---

## Instalación (una sola vez, desde una computadora)

### 1. Servidor
1. Entra a **script.google.com** con **desarrolloapps.ve@gmail.com** → **Nuevo proyecto**. Ponle nombre *CRM Fermetal*.
2. Borra todo lo que trae el editor y pega **todo** `Code.gs`. Guarda.
3. En la lista de funciones elige **`instalar`** → **▶ Ejecutar** → autoriza los permisos. En el *Registro de ejecución* aparece el enlace de la hoja creada.
4. **Implementar → Nueva implementación → ⚙️ Aplicación web**:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier persona**
5. **Implementar** y copia la URL que termina en `/exec`.

### 2. App
1. Abre `index.html` con un editor de texto, busca la línea `const API_URL = '';` y pega la URL entre las comillas:
   ```
   const API_URL = 'https://script.google.com/macros/s/…/exec';
   ```
2. En GitHub crea el repositorio **`crm`** (público) → **Add file → Upload files** → sube `index.html` y este README → **Commit changes**.
3. **Settings → Pages →** *Branch:* `main` / `(root)` → **Save**. En 1–2 minutos estará en https://ijmf1978.github.io/crm/

### 3. Primer uso (hazlo enseguida)
1. Abre el CRM, escribe el usuario **`ivan.marin`** y presiona **Entrar**. Te pedirá crear tu contraseña.
2. **Importar / Exportar → Traer clientes ahora.** Entran los clientes del cotizador (hoy son 4.426, con 7 RIF repetidos que se toman una sola vez).
3. **Usuarios → Nuevo usuario** para cada vendedor (ej. `maria.perez`). Puedes darle una contraseña inicial o dejarla vacía para que la cree él al entrar.
4. **Clientes**: filtra, selecciona y **asigna** a cada vendedor. Otra opción: exporta a Excel, llena la columna *Vendedor* y vuelve a importar.
5. Pasa a cada vendedor el enlace y su usuario.

> Mientras un usuario no tenga contraseña, cualquiera que sepa su nombre de usuario podría crearla. Por eso conviene que cada persona entre el mismo día que la creas, o ponerle contraseña inicial.

---

## Actualizar más adelante

- **Cambios en `index.html`:** súbelo de nuevo al repositorio `crm` (reemplaza el anterior).
- **Cambios en `Code.gs`:** pega el código nuevo, guarda y ve a **Implementar → Gestionar implementaciones → ✏️ Editar → Versión: Nueva versión → Implementar**. Así la URL `/exec` no cambia.
- **Traer clientes nuevos del cotizador:** repite *Traer clientes ahora* cuando quieras. Solo agrega los que faltan.

---

## Solución de problemas

| Síntoma | Solución |
|---|---|
| Pide "Conectar con el servidor" | Falta la URL en `API_URL`. También sirve abrir una vez `https://ijmf1978.github.io/crm/?api=URL_DEL_SERVIDOR` |
| "Google pide iniciar sesión" | La implementación no está en *Cualquier persona*. Edítala |
| "Sin conexión" | Abre la URL `/exec` en el navegador: debe decir **"Servidor del CRM activo."** |
| Un vendedor olvidó su contraseña | *Usuarios* → el vendedor → **Reiniciar contraseña**. La crea de nuevo al entrar |
| El administrador olvidó su contraseña | En Apps Script ejecuta **`reiniciarClaveAdministrador`** y entra con `ivan.marin` para crear una nueva |
| "Demasiados intentos" | Tras 5 contraseñas erradas el usuario se bloquea 10 minutos |

---

## Datos técnicos

- **App:** un solo archivo HTML/JS. SheetJS desde cdnjs para Excel. Guarda en el navegador solo la sesión y una copia de los datos para abrir rápido.
- **Servidor:** Google Apps Script + Google Sheets (hojas *Clientes, Usuarios, Seguimientos, Cambios, Config*).
  - Contraseñas guardadas con hash SHA-256 y sal; nunca se envían a la app.
  - Sesiones de 30 días; se cierran al cambiar o reiniciar la contraseña o al desactivar el usuario.
  - Escrituras con bloqueo (`LockService`) y versión de datos para que los teléfonos solo descarguen cuando algo cambió.
- La hoja de cálculo se puede abrir y revisar, pero conviene hacer los cambios desde el CRM para que queden en el historial.
