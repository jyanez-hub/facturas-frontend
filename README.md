# Frontend - Sistema de Facturación Electrónica

Frontend desarrollado en React con TypeScript y Material UI para el Sistema de Facturación Electrónica del SRI Ecuador.

## Tecnologías

- **React 19** - Biblioteca de UI
- **TypeScript** - Tipado estático
- **Material UI (MUI)** - Componentes de UI
- **React Router** - Enrutamiento
- **Axios** - Cliente HTTP
- **React Hook Form** - Gestión de formularios
- **Vite** - Build tool y dev server

## Instalación

```bash
# Instalar dependencias
npm install

# Configurar variables de entorno
cp .env.example .env
# Editar .env con la URL del backend

# Ejecutar en desarrollo
npm run dev

# Construir para producción
npm run build

# Preview de producción
npm run preview
```

## Configuración

### Variables de Entorno

Crea un archivo `.env` en la raíz del proyecto:

```env
VITE_API_URL=http://localhost:3000
```

Asegúrate de que el backend esté corriendo en la URL especificada.

## Estructura del Proyecto

```
src/
├── api/              # Servicios API
│   ├── client.ts     # Cliente Axios configurado
│   ├── auth.ts       # Endpoints de autenticación
│   ├── clients.ts    # Endpoints de clientes
│   ├── products.ts   # Endpoints de productos
│   └── invoices.ts   # Endpoints de facturas
├── components/        # Componentes reutilizables
│   ├── Layout/       # Componentes de layout
│   └── common/       # Componentes comunes
├── contexts/         # Context providers
│   └── AuthContext.tsx
├── pages/            # Páginas principales
│   ├── Login.tsx
│   ├── Register.tsx
│   ├── Dashboard.tsx
│   ├── Clients/
│   ├── Products/
│   └── Invoices/
├── utils/            # Utilidades
│   ├── constants.ts
│   └── formatters.ts
├── types/            # TypeScript types
└── routes.tsx        # Configuración de rutas
```

## Funcionalidades

### Autenticación
- Login y registro de usuarios
- Protección de rutas con AuthGuard
- Manejo de tokens JWT

### Gestión de Clientes
- Listado de clientes
- Crear, editar y eliminar clientes
- Validación de formularios

### Gestión de Productos
- Listado de productos
- Crear, editar y eliminar productos
- Soporte para productos con/sin IVA

### Facturación
- Crear facturas con múltiples productos
- Cálculo automático de totales e IVA
- Listado de facturas con estados del SRI
- Descarga de PDFs de facturas autorizadas

### Dashboard
- Estadísticas de facturación
- Resumen de facturas del mes
- Total facturado

## Desarrollo

El servidor de desarrollo se ejecuta en `http://localhost:5173` por defecto.

### Requisitos

- Node.js 18+
- npm o yarn

### Scripts Disponibles

- `npm run dev` - Inicia el servidor de desarrollo
- `npm run build` - Construye la aplicación para producción
- `npm run preview` - Previsualiza la build de producción
- `npm run lint` - Ejecuta el linter

## Integración con Backend

El frontend se comunica con el backend mediante una API RESTful. Asegúrate de que:

1. El backend esté corriendo en la URL especificada en `VITE_API_URL`
2. CORS esté configurado correctamente en el backend
3. Las credenciales de autenticación estén configuradas

## Notas

- Los tokens JWT se almacenan en localStorage
- El frontend redirige automáticamente al login si el token expira
- Los PDFs solo están disponibles después de que el SRI confirme la recepción de la factura
