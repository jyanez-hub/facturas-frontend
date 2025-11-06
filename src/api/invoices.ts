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

export interface EmailStatusResponse {
  claveAcceso: string;
  email_estado?: string;
  email_destinatario?: string;
  email_fecha_envio?: Date | string;
  email_intentos?: number;
  email_ultimo_error?: string;
}

export interface SendEmailRequest {
  email_destinatario: string;
}

export interface SendEmailResponse {
  message: string;
  claveAcceso: string;
  destinatario: string;
  estado: string;
}

export interface RegenerateResponse {
  message: string;
  facturaId: string;
}

export const invoicePdfApi = {
  // Obtener todos los PDFs
  getAll: async (): Promise<InvoicePDF[]> => {
    const response = await apiClient.get<InvoicePDF[]>('/api/v1/invoice-pdf');
    return response.data;
  },

  // Obtener PDF por ID de factura
  getByInvoiceId: async (invoiceId: string): Promise<InvoicePDF> => {
    const response = await apiClient.get<InvoicePDF>(`/api/v1/invoice-pdf/invoice/${invoiceId}`);
    return response.data;
  },

  // Obtener PDF por clave de acceso
  getByAccessKey: async (claveAcceso: string): Promise<InvoicePDF> => {
    const response = await apiClient.get<InvoicePDF>(`/api/v1/invoice-pdf/access-key/${claveAcceso}`);
    return response.data;
  },

  // Descargar PDF
  download: async (claveAcceso: string): Promise<Blob> => {
    const response = await apiClient.get(`/api/v1/invoice-pdf/download/${claveAcceso}`, {
      responseType: 'blob',
    });
    return response.data;
  },

  // Regenerar PDF
  regenerate: async (facturaId: string): Promise<RegenerateResponse> => {
    const response = await apiClient.post<RegenerateResponse>(`/api/v1/invoice-pdf/regenerate/${facturaId}`);
    return response.data;
  },

  // Enviar PDF por email
  sendEmail: async (claveAcceso: string, data: SendEmailRequest): Promise<SendEmailResponse> => {
    const response = await apiClient.post<SendEmailResponse>(`/api/v1/invoice-pdf/send-email/${claveAcceso}`, data);
    return response.data;
  },

  // Obtener estado del email
  getEmailStatus: async (claveAcceso: string): Promise<EmailStatusResponse> => {
    const response = await apiClient.get<EmailStatusResponse>(`/api/v1/invoice-pdf/email-status/${claveAcceso}`);
    return response.data;
  },

  // Reintentar envío de email
  retryEmail: async (claveAcceso: string): Promise<SendEmailResponse> => {
    const response = await apiClient.post<SendEmailResponse>(`/api/v1/invoice-pdf/retry-email/${claveAcceso}`);
    return response.data;
  },
};

