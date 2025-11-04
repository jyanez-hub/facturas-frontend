import apiClient from './client';
import type { Invoice, InvoicePDF } from '../types';

export interface CreateInvoiceData {
  factura: {
    cliente_id: string;
    fecha_emision: string;
    detalles: Array<{
      producto_id: string;
      cantidad: number;
      precio_unitario: number;
      descuento?: number;
    }>;
  };
}

export interface InvoiceResponse {
  success: boolean;
  data: Invoice;
  xml?: string;
}

export const invoicesApi = {
  getAll: async (): Promise<Invoice[]> => {
    const response = await apiClient.get<Invoice[]>('/api/v1/invoice');
    return response.data;
  },

  getById: async (id: string): Promise<Invoice> => {
    const response = await apiClient.get<Invoice>(`/api/v1/invoice/${id}`);
    return response.data;
  },

  createComplete: async (data: CreateInvoiceData): Promise<InvoiceResponse> => {
    const response = await apiClient.post<InvoiceResponse>('/api/v1/invoice/complete', data);
    return response.data;
  },
};

export const invoicePdfApi = {
  getByInvoiceId: async (invoiceId: string): Promise<InvoicePDF> => {
    const response = await apiClient.get<InvoicePDF>(`/api/v1/invoice-pdf/invoice/${invoiceId}`);
    return response.data;
  },

  download: async (claveAcceso: string): Promise<Blob> => {
    const response = await apiClient.get(`/api/v1/invoice-pdf/download/${claveAcceso}`, {
      responseType: 'blob',
    });
    return response.data;
  },

  getByAccessKey: async (claveAcceso: string): Promise<InvoicePDF> => {
    const response = await apiClient.get<InvoicePDF>(`/api/v1/invoice-pdf/access-key/${claveAcceso}`);
    return response.data;
  },
};

