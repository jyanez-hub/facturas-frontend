// Constantes de la aplicación

export const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:3000';

export const ROUTES = {
  LOGIN: '/login',
  REGISTER: '/register',
  DASHBOARD: '/dashboard',
  CLIENTS: '/clients',
  CLIENT_NEW: '/clients/new',
  CLIENT_EDIT: (id: string) => `/clients/${id}`,
  PRODUCTS: '/products',
  PRODUCT_NEW: '/products/new',
  PRODUCT_EDIT: (id: string) => `/products/${id}`,
  INVOICES: '/invoices',
  INVOICE_NEW: '/invoices/new',
  INVOICE_DETAIL: (id: string) => `/invoices/${id}`,
};

export const STORAGE_KEYS = {
  AUTH_TOKEN: 'auth_token',
  USER_DATA: 'user_data',
};

export const INVOICE_STATUS = {
  PENDIENTE: 'PENDIENTE',
  RECIBIDA: 'RECIBIDA',
  DEVUELTA: 'DEVUELTA',
  AUTORIZADA: 'AUTORIZADA',
};

export const INVOICE_STATUS_COLORS = {
  PENDIENTE: 'warning',
  RECIBIDA: 'info',
  DEVUELTA: 'error',
  AUTORIZADA: 'success',
} as const;

