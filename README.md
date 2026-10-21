# WhatsApp Difusión Pro - Modo Ventana Única (Sin Recargar)

Aplicación de escritorio en Python para envío automatizado de mensajes de WhatsApp que trabaja **100% dentro de una misma ventana de WhatsApp Web, sin abrir pestañas nuevas y sin recargar la página**.

---

## ⚡ ¿Cómo funciona el Modo Ventana Única?

1. **Cero Pestañas Nuevas:**
   - Todo ocurre en la pestaña de WhatsApp Web que ya tienes abierta.
   - No abre ni cierra ventanas secundarias.

2. **Cero Recargas de Página:**
   - En lugar de forzar a recargar la web con URLs pesadas (`send?phone=...`), utiliza el **motor de búsqueda interno de WhatsApp Web** (`Ctrl + Alt + N`).
   - Busca el número telefónico, selecciona el chat y lo abre de inmediato (tarda menos de 1 segundo en cambiar de conversación).

3. **Despacho Fiel y Fluido:**
   - Pega el mensaje personalizado con su saludo automático directamente en la caja de texto preservando emojis, saltos de línea y tildes.
   - Envía con `Enter` y limpia la búsqueda para saltar de inmediato al siguiente contacto.

---

## 📦 Requisitos e Instalación

Instala las dependencias necesarias con:

```bash
pip install -r requirements.txt
```

---

## 🚀 Cómo Usarlo

1. Ten tu ventana de **WhatsApp Web abierta y visible en pantalla** con tu sesión activa.
2. Abre la aplicación ejecutando `Iniciar_WhatsApp_Difusion.bat` o desde la consola:
   ```bash
   py main.py
   ```
3. Selecciona tu archivo de contactos Excel (`.xlsx`) o CSV (incluye un archivo base `contactos.csv`).
4. Revisa el mensaje y las variables dinámicas en la vista previa.
5. Presiona **"🚀 Iniciar Envíos"**.
6. Durante la cuenta regresiva (3 segundos), pon al frente la ventana de WhatsApp Web.
7. A partir de ahí, el script navegará y enviará los mensajes de forma continua y automatizada.
