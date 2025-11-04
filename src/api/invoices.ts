import apiClient from './client';
import type { Invoice, InvoicePDF } from '../types';

export interface InvoiceTaxDetail {
  impuesto: {
    codigo: string;
    codigoPorcentaje: string;
    tarifa: string;
    baseImponible: string;
    valor: string;
  };
}

export interface InvoiceDetailRequest {
  detalle: {
    codigoPrincipal: string;
    descripcion: string;
    cantidad: string;
    precioUnitario: string;
    precioTotalSinImpuesto: string;
    impuestos: InvoiceTaxDetail[];
  };
}

export interface InvoiceInfoTributaria {
  ruc: string;
  claveAcceso: string;
  secuencial: string;
}

export interface InvoiceInfoFactura {
  fechaEmision: string;
  tipoIdentificacionComprador: string;
  identificacionComprador: string;
  razonSocialComprador: string;
  totalSinImpuestos: string;
  importeTotal: string;
}

export interface CreateInvoiceData {
  factura: {
    infoTributaria: InvoiceInfoTributaria;
    infoFactura: InvoiceInfoFactura;
    detalles: InvoiceDetailRequest[];
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

