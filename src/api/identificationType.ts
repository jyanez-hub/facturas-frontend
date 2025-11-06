import apiClient from './client';
import type { IdentificationType } from '../types';

export const identificationTypeApi = {
  getAll: async (): Promise<IdentificationType[]> => {
    const response = await apiClient.get<IdentificationType[]>('/api/v1/identification-type');
    return response.data;
  },

  getById: async (id: string): Promise<IdentificationType> => {
    const response = await apiClient.get<IdentificationType>(`/api/v1/identification-type/${id}`);
    return response.data;
  },

  create: async (data: Omit<IdentificationType, '_id'>): Promise<IdentificationType> => {
    const response = await apiClient.post<IdentificationType>('/api/v1/identification-type', data);
    return response.data;
  },

  update: async (id: string, data: Partial<IdentificationType>): Promise<IdentificationType> => {
    const response = await apiClient.put<IdentificationType>(`/api/v1/identification-type/${id}`, data);
    return response.data;
  },

  delete: async (id: string): Promise<void> => {
    await apiClient.delete(`/api/v1/identification-type/${id}`);
  },
};

