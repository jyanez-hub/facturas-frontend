import apiClient from './client';
import type { IssuingCompany } from '../types';

export const issuingCompanyApi = {
  getAll: async (): Promise<IssuingCompany[]> => {
    const response = await apiClient.get<IssuingCompany[]>('/api/v1/issuing-company');
    return response.data;
  },

  getById: async (id: string): Promise<IssuingCompany> => {
    const response = await apiClient.get<IssuingCompany>(`/api/v1/issuing-company/${id}`);
    return response.data;
  },

  update: async (id: string, data: Partial<IssuingCompany>): Promise<IssuingCompany> => {
    const response = await apiClient.put<IssuingCompany>(`/api/v1/issuing-company/${id}`, data);
    return response.data;
  },

  delete: async (id: string): Promise<void> => {
    await apiClient.delete(`/api/v1/issuing-company/${id}`);
  },
};

