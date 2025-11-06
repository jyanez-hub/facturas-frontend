import apiClient from './client';
import type { InvoiceDetail } from '../types';

export const invoiceDetailApi = {
  getAll: async (): Promise<InvoiceDetail[]> => {
    const response = await apiClient.get<InvoiceDetail[]>('/api/v1/invoice-detail');
    return response.data;
  },

  getById: async (id: string): Promise<InvoiceDetail> => {
    const response = await apiClient.get<InvoiceDetail>(`/api/v1/invoice-detail/${id}`);
    return response.data;
  },

  create: async (data: Omit<InvoiceDetail, '_id'>): Promise<InvoiceDetail> => {
    const response = await apiClient.post<InvoiceDetail>('/api/v1/invoice-detail', data);
    return response.data;
  },

  update: async (id: string, data: Partial<InvoiceDetail>): Promise<InvoiceDetail> => {
    const response = await apiClient.put<InvoiceDetail>(`/api/v1/invoice-detail/${id}`, data);
    return response.data;
  },

  delete: async (id: string): Promise<void> => {
    await apiClient.delete(`/api/v1/invoice-detail/${id}`);
  },
};

