import apiClient from './client';
import type { Client } from '../types';

export const clientsApi = {
  getAll: async (): Promise<Client[]> => {
    const response = await apiClient.get<Client[]>('/api/v1/client');
    return response.data;
  },

  getById: async (id: string): Promise<Client> => {
    const response = await apiClient.get<Client>(`/api/v1/client/${id}`);
    return response.data;
  },

  create: async (client: Omit<Client, '_id'>): Promise<Client> => {
    const response = await apiClient.post<Client>('/api/v1/client', client);
    return response.data;
  },

  update: async (id: string, client: Partial<Client>): Promise<Client> => {
    const response = await apiClient.put<Client>(`/api/v1/client/${id}`, client);
    return response.data;
  },

  delete: async (id: string): Promise<void> => {
    await apiClient.delete(`/api/v1/client/${id}`);
  },
};

