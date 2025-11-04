import { Routes, Route, Navigate } from 'react-router-dom';
import { AuthGuard } from './components/common/AuthGuard';
import { AppLayout } from './components/Layout/AppLayout';
import { Login } from './pages/Login';
import { Register } from './pages/Register';
import { Dashboard } from './pages/Dashboard';
import { ClientList } from './pages/Clients/ClientList';
import { ClientForm } from './pages/Clients/ClientForm';
import { ProductList } from './pages/Products/ProductList';
import { ProductForm } from './pages/Products/ProductForm';
import { InvoiceList } from './pages/Invoices/InvoiceList';
import { InvoiceForm } from './pages/Invoices/InvoiceForm';
import { InvoiceDetail } from './pages/Invoices/InvoiceDetail';
import { ROUTES } from './utils/constants';

export const AppRoutes = () => {
  return (
    <Routes>
      {/* Rutas públicas */}
      <Route path={ROUTES.LOGIN} element={<Login />} />
      <Route path={ROUTES.REGISTER} element={<Register />} />

      {/* Rutas protegidas */}
      <Route
        path="/"
        element={
          <AuthGuard>
            <AppLayout />
          </AuthGuard>
        }
      >
        <Route index element={<Navigate to={ROUTES.DASHBOARD} replace />} />
        <Route path={ROUTES.DASHBOARD} element={<Dashboard />} />
        
        {/* Clientes */}
        <Route path={ROUTES.CLIENTS} element={<ClientList />} />
        <Route path={ROUTES.CLIENT_NEW} element={<ClientForm />} />
        <Route path="/clients/:id" element={<ClientForm />} />
        
        {/* Productos */}
        <Route path={ROUTES.PRODUCTS} element={<ProductList />} />
        <Route path={ROUTES.PRODUCT_NEW} element={<ProductForm />} />
        <Route path="/products/:id" element={<ProductForm />} />
        
        {/* Facturas */}
        <Route path={ROUTES.INVOICES} element={<InvoiceList />} />
        <Route path={ROUTES.INVOICE_NEW} element={<InvoiceForm />} />
        <Route path="/invoices/:id" element={<InvoiceDetail />} />

        {/* Ruta por defecto */}
        <Route path="*" element={<Navigate to={ROUTES.DASHBOARD} replace />} />
      </Route>
    </Routes>
  );
};

