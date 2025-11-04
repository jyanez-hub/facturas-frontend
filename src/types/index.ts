// Tipos principales para la aplicación

export interface User {
  id: string;
  email: string;
}

export interface Company {
  id: string;
  ruc: string;
  razon_social: string;
  nombre_comercial: string;
}

export interface AuthResponse {
  token: string;
  user: User;
  company: Company;
}

export interface Client {
  _id?: string;
  tipo_identificacion_id: string;
  identificacion: string;
  razon_social: string;
  direccion?: string;
  email?: string;
  telefono?: string;
}

export interface Product {
  _id?: string;
  codigo: string;
  descripcion: string;
  precio_unitario: number;
  tiene_iva: boolean;
  descripcion_adicional?: string;
}

export interface InvoiceDetail {
  producto_id: string;
  cantidad: number;
  precio_unitario: number;
  descuento?: number;
  iva?: number;
  subtotal: number;
}

export interface Invoice {
  _id?: string;
  empresa_emisora_id: string;
  cliente_id: string;
  fecha_emision: Date | string;
  clave_acceso?: string;
  secuencial?: string;
  estado?: string;
  total_sin_impuestos: number;
  total_iva: number;
  total_con_impuestos: number;
  detalles: InvoiceDetail[];
  xml?: string;
  xml_firmado?: string;
  autorizacion_numero?: string;
  fecha_autorizacion?: Date | string;
  sri_estado?: string;
  sri_mensajes?: any;
  sri_fecha_envio?: Date;
  sri_fecha_respuesta?: Date;
}

export interface InvoicePDF {
  _id?: string;
  factura_id: string;
  claveAcceso: string;
  pdf_buffer?: ArrayBuffer | Uint8Array;
  pdf_url?: string;
  email_estado?: string;
  email_destinatario?: string;
}

export interface IdentificationType {
  _id?: string;
  codigo: string;
  descripcion: string;
}

