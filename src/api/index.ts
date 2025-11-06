// Exportación centralizada de todas las APIs
export { authApi, type LoginCredentials, type RegisterData, type RegistrationStatus } from './auth';
export { clientsApi } from './clients';
export { productsApi } from './products';
export { 
  invoicesApi, 
  invoicePdfApi,
  type CreateInvoiceData,
  type InvoiceResponse,
  type EmailStatusResponse,
  type SendEmailRequest,
  type SendEmailResponse,
  type RegenerateResponse,
} from './invoices';
export { identificationTypeApi } from './identificationType';
export { issuingCompanyApi } from './issuingCompany';
export { invoiceDetailApi } from './invoiceDetail';

