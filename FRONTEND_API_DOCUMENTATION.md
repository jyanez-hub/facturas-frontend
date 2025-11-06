# 📘 Documentación de APIs del Frontend

Esta documentación describe todos los servicios API disponibles en el frontend para consumir los endpoints del backend.

## 📁 Estructura de APIs

Todas las APIs están ubicadas en `src/api/` y pueden ser importadas desde el índice central:

```typescript
import { authApi, clientsApi, productsApi, /* ... */ } from './api';
```

---

## 🔐 Autenticación (`auth.ts`)

### authApi.login(credentials)
Inicia sesión con email y contraseña.

```typescript
const response = await authApi.login({
  email: 'usuario@ejemplo.com',
  password: 'contraseña123'
});
// Retorna: { token, user, company }
```

### authApi.register(data)
Registra un nuevo usuario y empresa emisora.

```typescript
const response = await authApi.register({
  email: 'usuario@ejemplo.com',
  password: 'contraseña123',
  ruc: '1799999999001',
  razon_social: 'Mi Empresa S.A.',
  nombre_comercial: 'Mi Empresa',
  masterKey: 'clave_maestra' // Solo para primer registro
});
// Retorna: { token, user, company }
```

### authApi.getStatus()
Obtiene el estado del sistema de registro.

```typescript
const status = await authApi.getStatus();
// Retorna: { firstRegistration, registrationDisabled, requiresInvitation, masterKeyRequired }
```

---

## 👥 Clientes (`clients.ts`)

### clientsApi.getAll()
Obtiene todos los clientes.

```typescript
const clients = await clientsApi.getAll();
```

### clientsApi.getById(id)
Obtiene un cliente por ID.

```typescript
const client = await clientsApi.getById('clientId123');
```

### clientsApi.create(client)
Crea un nuevo cliente.

```typescript
const newClient = await clientsApi.create({
  tipo_identificacion_id: 'typeId',
  identificacion: '0923456789',
  razon_social: 'Juan Pérez',
  email: 'juan@ejemplo.com',
  telefono: '0999999999',
  direccion: 'Calle Principal 123'
});
```

### clientsApi.update(id, client)
Actualiza un cliente existente.

```typescript
const updated = await clientsApi.update('clientId123', {
  email: 'nuevo@email.com',
  telefono: '0988888888'
});
```

### clientsApi.delete(id)
Elimina un cliente.

```typescript
await clientsApi.delete('clientId123');
```

---

## 📦 Productos (`products.ts`)

### productsApi.getAll()
Obtiene todos los productos.

```typescript
const products = await productsApi.getAll();
```

### productsApi.getById(id)
Obtiene un producto por ID.

```typescript
const product = await productsApi.getById('productId123');
```

### productsApi.create(product)
Crea un nuevo producto.

```typescript
const newProduct = await productsApi.create({
  codigo: 'PROD001',
  descripcion: 'Laptop Dell XPS 15',
  precio_unitario: 1500.00,
  tiene_iva: true
});
```

### productsApi.update(id, product)
Actualiza un producto existente.

```typescript
const updated = await productsApi.update('productId123', {
  precio_unitario: 1450.00
});
```

### productsApi.delete(id)
Elimina un producto.

```typescript
await productsApi.delete('productId123');
```

---

## 🧾 Facturas (`invoices.ts`)

### invoicesApi.getAll()
Obtiene todas las facturas.

```typescript
const invoices = await invoicesApi.getAll();
```

### invoicesApi.getById(id)
Obtiene una factura por ID.

```typescript
const invoice = await invoicesApi.getById('invoiceId123');
```

### invoicesApi.createComplete(data)
Crea una factura completa con todos sus detalles, XML firmado y envío al SRI.

```typescript
const response = await invoicesApi.createComplete({
  factura: {
    infoTributaria: {
      ruc: '1799999999001',
      secuencial: '123'
    },
    infoFactura: {
      fechaEmision: '05/11/2025',
      tipoIdentificacionComprador: '05',
      identificacionComprador: '0923456789',
      totalSinImpuestos: '100.00',
      importeTotal: '112.00'
    },
    detalles: [
      {
        detalle: {
          codigoPrincipal: 'PROD001',
          descripcion: 'Producto 1',
          cantidad: '1',
          precioUnitario: '100.00',
          precioTotalSinImpuesto: '100.00',
          impuestos: [
            {
              impuesto: {
                codigo: '2',
                codigoPorcentaje: '2',
                tarifa: '12',
                baseImponible: '100.00',
                valor: '12.00'
              }
            }
          ]
        }
      }
    ]
  }
});
// Retorna: { success, data: { factura, detalles, xml }, xml }
```

---

## 📄 PDFs de Facturas (`invoices.ts` - invoicePdfApi)

### invoicePdfApi.getAll()
Obtiene todos los PDFs generados.

```typescript
const pdfs = await invoicePdfApi.getAll();
```

### invoicePdfApi.getByInvoiceId(invoiceId)
Obtiene el PDF de una factura específica.

```typescript
const pdf = await invoicePdfApi.getByInvoiceId('invoiceId123');
```

### invoicePdfApi.getByAccessKey(claveAcceso)
Obtiene el PDF por clave de acceso.

```typescript
const pdf = await invoicePdfApi.getByAccessKey('0511202501179999999900110010010000001231234567812');
```

### invoicePdfApi.download(claveAcceso)
Descarga el archivo PDF.

```typescript
const blob = await invoicePdfApi.download('claveAcceso');
// Crear URL y descargar
const url = window.URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = `factura_${claveAcceso}.pdf`;
a.click();
```

### invoicePdfApi.regenerate(facturaId)
Regenera el PDF de una factura.

```typescript
const response = await invoicePdfApi.regenerate('facturaId123');
// Retorna: { message, facturaId }
```

### invoicePdfApi.sendEmail(claveAcceso, data)
Envía el PDF por email.

```typescript
const response = await invoicePdfApi.sendEmail('claveAcceso', {
  email_destinatario: 'cliente@ejemplo.com'
});
// Retorna: { message, claveAcceso, destinatario, estado }
```

### invoicePdfApi.getEmailStatus(claveAcceso)
Obtiene el estado del envío por email.

```typescript
const status = await invoicePdfApi.getEmailStatus('claveAcceso');
// Retorna: { claveAcceso, email_estado, email_destinatario, email_fecha_envio, ... }
```

### invoicePdfApi.retryEmail(claveAcceso)
Reintenta enviar el email.

```typescript
const response = await invoicePdfApi.retryEmail('claveAcceso');
// Retorna: { message, claveAcceso, estado }
```

---

## 🆔 Tipos de Identificación (`identificationType.ts`)

### identificationTypeApi.getAll()
Obtiene todos los tipos de identificación.

```typescript
const types = await identificationTypeApi.getAll();
```

### identificationTypeApi.getById(id)
Obtiene un tipo por ID.

```typescript
const type = await identificationTypeApi.getById('typeId123');
```

### identificationTypeApi.create(data)
Crea un nuevo tipo de identificación.

```typescript
const newType = await identificationTypeApi.create({
  codigo: '05',
  nombre: 'CEDULA',
  descripcion: 'Cédula de identidad'
});
```

### identificationTypeApi.update(id, data)
Actualiza un tipo de identificación.

```typescript
const updated = await identificationTypeApi.update('typeId123', {
  descripcion: 'Nueva descripción'
});
```

### identificationTypeApi.delete(id)
Elimina un tipo de identificación.

```typescript
await identificationTypeApi.delete('typeId123');
```

---

## 🏢 Empresas Emisoras (`issuingCompany.ts`)

### issuingCompanyApi.getAll()
Obtiene todas las empresas emisoras.

```typescript
const companies = await issuingCompanyApi.getAll();
```

### issuingCompanyApi.getById(id)
Obtiene una empresa por ID.

```typescript
const company = await issuingCompanyApi.getById('companyId123');
```

### issuingCompanyApi.update(id, data)
Actualiza una empresa emisora.

```typescript
const updated = await issuingCompanyApi.update('companyId123', {
  telefono: '0987654321',
  direccion: 'Nueva dirección',
  tipo_ambiente: 2 // Cambiar a producción
});
```

### issuingCompanyApi.delete(id)
Elimina una empresa emisora.

```typescript
await issuingCompanyApi.delete('companyId123');
```

---

## 📝 Detalles de Facturas (`invoiceDetail.ts`)

### invoiceDetailApi.getAll()
Obtiene todos los detalles de facturas.

```typescript
const details = await invoiceDetailApi.getAll();
```

### invoiceDetailApi.getById(id)
Obtiene un detalle por ID.

```typescript
const detail = await invoiceDetailApi.getById('detailId123');
```

### invoiceDetailApi.create(data)
Crea un nuevo detalle de factura.

```typescript
const newDetail = await invoiceDetailApi.create({
  factura_id: 'invoiceId123',
  producto_id: 'productId123',
  cantidad: 2,
  precio_unitario: 100.00,
  subtotal: 200.00,
  porcentaje_iva: 12,
  valor_iva: 24.00,
  total: 224.00
});
```

### invoiceDetailApi.update(id, data)
Actualiza un detalle de factura.

```typescript
const updated = await invoiceDetailApi.update('detailId123', {
  cantidad: 3,
  subtotal: 300.00
});
```

### invoiceDetailApi.delete(id)
Elimina un detalle de factura.

```typescript
await invoiceDetailApi.delete('detailId123');
```

---

## 🔧 Cliente HTTP (`client.ts`)

El cliente HTTP está configurado con:

### Características
- **Base URL**: Configurada desde `VITE_API_URL` o `http://localhost:3000`
- **Interceptor de Request**: Agrega automáticamente el token JWT a todas las peticiones
- **Interceptor de Response**: Maneja errores 401 y redirige al login automáticamente
- **Credenciales**: Incluye cookies en las peticiones (`withCredentials: true`)

### Uso
No necesitas configurar headers manualmente:

```typescript
import apiClient from './api/client';

// El token se agrega automáticamente
const response = await apiClient.get('/api/v1/client');
```

---

## 📱 Páginas Disponibles

### Tipos de Identificación
- **Lista**: `/identification-types` - `IdentificationTypeList`
- **Nuevo**: `/identification-types/new` - `IdentificationTypeForm`
- **Editar**: `/identification-types/:id` - `IdentificationTypeForm`

### Empresas Emisoras
- **Lista**: `/issuing-company` - `IssuingCompanyList`
- **Editar**: `/issuing-company/:id` - `IssuingCompanyForm`

### Clientes
- **Lista**: `/clients` - `ClientList`
- **Nuevo**: `/clients/new` - `ClientForm`
- **Editar**: `/clients/:id` - `ClientForm`

### Productos
- **Lista**: `/products` - `ProductList`
- **Nuevo**: `/products/new` - `ProductForm`
- **Editar**: `/products/:id` - `ProductForm`

### Facturas
- **Lista**: `/invoices` - `InvoiceList`
- **Nueva**: `/invoices/new` - `InvoiceForm`
- **Detalle**: `/invoices/:id` - `InvoiceDetail`

---

## 🎯 Tipos TypeScript

Todos los tipos están definidos en `src/types/index.ts`:

```typescript
import type { 
  User,
  Company,
  AuthResponse,
  Client,
  Product,
  Invoice,
  InvoiceDetail,
  InvoicePDF,
  IdentificationType,
  IssuingCompany
} from './types';
```

---

## ⚡ Manejo de Errores

Todas las APIs usan try-catch y devuelven errores en formato:

```typescript
try {
  const data = await clientsApi.getAll();
} catch (error: any) {
  const message = error.response?.data?.message || 'Error al cargar datos';
  console.error(message);
}
```

---

## 🔐 Autenticación

El token se guarda automáticamente en `localStorage`:

```typescript
// Al hacer login o register
localStorage.setItem('auth_token', response.token);
localStorage.setItem('user_data', JSON.stringify(response.user));

// Se adjunta automáticamente a todas las peticiones
// No necesitas hacer nada manualmente
```

---

## 🚀 Ejemplo de Uso Completo

```typescript
import { 
  authApi, 
  clientsApi, 
  productsApi, 
  invoicesApi,
  invoicePdfApi 
} from './api';

// 1. Login
const authResponse = await authApi.login({
  email: 'admin@empresa.com',
  password: 'password123'
});

// 2. Crear cliente
const client = await clientsApi.create({
  tipo_identificacion_id: 'typeId',
  identificacion: '0923456789',
  razon_social: 'Cliente Test',
  email: 'cliente@test.com'
});

// 3. Crear producto
const product = await productsApi.create({
  codigo: 'PROD001',
  descripcion: 'Producto Test',
  precio_unitario: 100.00,
  tiene_iva: true
});

// 4. Crear factura completa
const invoice = await invoicesApi.createComplete({
  factura: {
    infoTributaria: {
      ruc: authResponse.company.ruc,
      secuencial: '001'
    },
    infoFactura: {
      fechaEmision: '05/11/2025',
      tipoIdentificacionComprador: '05',
      identificacionComprador: client.identificacion,
      totalSinImpuestos: '100.00',
      importeTotal: '112.00'
    },
    detalles: [/* ... */]
  }
});

// 5. Descargar PDF
const pdfBlob = await invoicePdfApi.download(invoice.data.factura.data.clave_acceso);

// 6. Enviar por email
await invoicePdfApi.sendEmail(invoice.data.factura.data.clave_acceso, {
  email_destinatario: client.email
});
```

---

## 📌 Notas Importantes

1. **Todas las APIs requieren autenticación** excepto:
   - `authApi.login()`
   - `authApi.register()`
   - `authApi.getStatus()`

2. **El token expira después de 4 días**. Si recibes 401, el usuario será redirigido automáticamente al login.

3. **Las facturas se procesan de forma asíncrona**. El estado inicial es `PENDIENTE` y se actualiza cuando el SRI responde.

4. **Los PDFs se generan automáticamente** al crear una factura completa.

5. **Las empresas emisoras no se crean directamente** - se crean durante el registro de usuario.

---

## 🛠️ Configuración del Entorno

Crea un archivo `.env` en la raíz del frontend:

```env
VITE_API_URL=http://localhost:3000
```

---

¡Todas las APIs están listas para usar! 🎉

