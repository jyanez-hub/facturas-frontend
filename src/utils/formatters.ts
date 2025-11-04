import { format } from 'date-fns';

// Formateo de moneda
export const formatCurrency = (amount: number): string => {
  return new Intl.NumberFormat('es-EC', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 2,
  }).format(amount);
};

// Formateo de fecha
export const formatDate = (date: Date | string): string => {
  if (!date) return '';
  const dateObj = typeof date === 'string' ? new Date(date) : date;
  if (isNaN(dateObj.getTime())) return '';
  return format(dateObj, 'dd/MM/yyyy');
};

// Formateo de fecha y hora
export const formatDateTime = (date: Date | string): string => {
  if (!date) return '';
  const dateObj = typeof date === 'string' ? new Date(date) : date;
  if (isNaN(dateObj.getTime())) return '';
  return format(dateObj, 'dd/MM/yyyy HH:mm');
};

// Validar RUC ecuatoriano
export const isValidRUC = (ruc: string): boolean => {
  const rucRegex = /^\d{10}001$/;
  return rucRegex.test(ruc);
};

// Validar cédula ecuatoriana
export const isValidCedula = (cedula: string): boolean => {
  const cedulaRegex = /^\d{10}$/;
  return cedulaRegex.test(cedula);
};

