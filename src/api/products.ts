import apiClient from './client';
import type { Product } from '../types';

export const productsApi = {
  getAll: async (): Promise<Product[]> => {
    const response = await apiClient.get<Product[]>('/api/v1/product');
    return response.data;
  },

  getById: async (id: string): Promise<Product> => {
    const response = await apiClient.get<Product>(`/api/v1/product/${id}`);
    return response.data;
  },

  create: async (product: Omit<Product, '_id'>): Promise<Product> => {
    const response = await apiClient.post<Product>('/api/v1/product', product);
    return response.data;
  },

  update: async (id: string, product: Partial<Product>): Promise<Product> => {
    const response = await apiClient.put<Product>(`/api/v1/product/${id}`, product);
    return response.data;
  },

  delete: async (id: string): Promise<void> => {
    await apiClient.delete(`/api/v1/product/${id}`);
  },
};

