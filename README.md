# Bot RRHH — Hostel Che

Bot de WhatsApp para fichaje de empleados con selfie, ubicación GPS y PIN.

## Flujo del empleado

1. Manda **"fichar"** al número de RRHH
2. Envía una **selfie**
3. Comparte su **ubicación** (clip 📎 → Ubicación → Enviar ubicación actual)
4. Escribe su **PIN** de 4 dígitos
5. Recibe confirmación de entrada o salida

## Sedes configuradas

- Oficina (50m)
- Che Playa (50m)  
- Che Suite Playa (50m)

## Deploy en Railway

1. Subir este repo a GitHub
2. Conectar en railway.app
3. Al primer deploy, ver los logs para escanear el QR con WhatsApp
4. Una vez escaneado, el bot corre 24/7

## Variables de entorno (opcional)

Se pueden mover a variables de entorno en Railway:
- `SHEETS_URL` — URL del Google Apps Script
- `NUMERO_ADMIN` — número que recibe el reporte diario

## Datos guardados en Google Sheets

Cada fichaje registra: empleado, departamento, tipo, hora, fecha, timestamp, foto (sí/no), sede, canal (WhatsApp).
