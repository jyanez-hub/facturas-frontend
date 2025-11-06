import apiClient from './client';
import type { AuthResponse } from '../types';

export interface LoginCredentials {
  email: string;
  password: string;
}

export interface RegisterData {
  email: string;
  password: string;
  masterKey?: string;
  invitationCode?: string;
  ruc: string;
  razon_social: string;
  nombre_comercial?: string;
  direccion?: string;
  telefono?: string;
  company_email?: string;
  codigo_establecimiento?: string;
  punto_emision?: string;
  tipo_ambiente?: number;
  certificate?: string;
  certificatePassword?: string;
}

export interface RegistrationStatus {
  firstRegistration: boolean;
  registrationDisabled: boolean;
  requiresInvitation: boolean;
  masterKeyRequired: boolean;
}

export const authApi = {
  login: async (credentials: LoginCredentials): Promise<AuthResponse> => {
    const response = await apiClient.post<AuthResponse>('/auth', credentials);
    return response.data;
  },

  register: async (data: RegisterData): Promise<AuthResponse> => {
    const response = await apiClient.post<AuthResponse>('/register', data);
    return response.data;
  },

  getStatus: async (): Promise<RegistrationStatus> => {
    const response = await apiClient.get<RegistrationStatus>('/status');
    return response.data;
  },
};

