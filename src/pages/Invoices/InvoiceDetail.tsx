import React, { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import {
  Box,
  Button,
  Paper,
  Typography,
  Chip,
  Divider,
  Alert,
} from '@mui/material';
import { Grid2 } from '@mui/material';
import {
  ArrowBack as ArrowBackIcon,
  Download as DownloadIcon,
} from '@mui/icons-material';
import { invoicesApi, invoicePdfApi } from '../../api/invoices';
import type { Invoice } from '../../types';
import { Loading } from '../../components/common/Loading';
import { ErrorAlert } from '../../components/common/ErrorAlert';
import { formatCurrency, formatDate } from '../../utils/formatters';
import { ROUTES, INVOICE_STATUS_COLORS } from '../../utils/constants';

export const InvoiceDetail: React.FC = () => {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const [invoice, setInvoice] = useState<Invoice | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [downloadingPdf, setDownloadingPdf] = useState(false);

  useEffect(() => {
    if (id) {
      loadInvoice();
    }
  }, [id]);

  const loadInvoice = async () => {
    if (!id) return;
    try {
      setLoading(true);
      const data = await invoicesApi.getById(id);
      setInvoice(data);
    } catch (err: any) {
      setError(err.message || 'Error al cargar factura');
    } finally {
      setLoading(false);
    }
  };

  const handleDownloadPdf = async () => {
    if (!invoice?.clave_acceso) {
      setError('La factura no tiene clave de acceso');
      return;
    }

    try {
      setDownloadingPdf(true);
      const blob = await invoicePdfApi.download(invoice.clave_acceso);
      
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `factura_${invoice.clave_acceso}.pdf`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err: any) {
      if (err.response?.status === 404) {
        setError('El PDF aún no está disponible. La factura debe ser recibida por el SRI primero.');
      } else {
        setError(err.message || 'Error al descargar PDF');
      }
    } finally {
      setDownloadingPdf(false);
    }
  };

  if (loading) {
    return <Loading />;
  }

  if (!invoice) {
    return (
      <Box>
        <Alert severity="error">Factura no encontrada</Alert>
        <Button onClick={() => navigate(ROUTES.INVOICES)} sx={{ mt: 2 }}>
          Volver a Facturas
        </Button>
      </Box>
    );
  }

  const status = invoice.sri_estado || invoice.estado || 'PENDIENTE';
  const statusColor = 
    INVOICE_STATUS_COLORS[status as keyof typeof INVOICE_STATUS_COLORS] || 'default';

  return (
    <Box>
      <Box display="flex" alignItems="center" gap={2} mb={3}>
        <Button
          startIcon={<ArrowBackIcon />}
          onClick={() => navigate(ROUTES.INVOICES)}
        >
          Volver
        </Button>
        <Typography variant="h4">Detalle de Factura</Typography>
      </Box>

      <ErrorAlert
        message={error}
        open={!!error}
        onClose={() => setError('')}
      />

      <Grid2 container spacing={3}>
        <Grid2 xs={12} md={8}>
          <Paper sx={{ p: 3 }}>
            <Typography variant="h6" gutterBottom>
              Información General
            </Typography>
            <Divider sx={{ mb: 2 }} />

            <Grid2 container spacing={2}>
              <Grid2 xs={12} sm={6}>
                <Typography variant="body2" color="textSecondary">
                  Secuencial
                </Typography>
                <Typography variant="body1">{invoice.secuencial || '-'}</Typography>
              </Grid2>
              <Grid2 xs={12} sm={6}>
                <Typography variant="body2" color="textSecondary">
                  Fecha de Emisión
                </Typography>
                <Typography variant="body1">
                  {formatDate(invoice.fecha_emision)}
                </Typography>
              </Grid2>
              <Grid2 xs={12}>
                <Typography variant="body2" color="textSecondary">
                  Clave de Acceso
                </Typography>
                <Typography variant="body1" sx={{ fontFamily: 'monospace', fontSize: '0.9rem' }}>
                  {invoice.clave_acceso || '-'}
                </Typography>
              </Grid2>
              <Grid2 xs={12} sm={6}>
                <Typography variant="body2" color="textSecondary">
                  Estado SRI
                </Typography>
                <Chip
                  label={status}
                  color={statusColor as any}
                  sx={{ mt: 0.5 }}
                />
              </Grid2>
              {invoice.autorizacion_numero && (
                <Grid2 xs={12} sm={6}>
                  <Typography variant="body2" color="textSecondary">
                    Número de Autorización
                  </Typography>
                  <Typography variant="body1">
                    {invoice.autorizacion_numero}
                  </Typography>
                </Grid2>
              )}
            </Grid2>
          </Paper>
        </Grid2>

        <Grid2 xs={12} md={4}>
          <Paper sx={{ p: 3 }}>
            <Typography variant="h6" gutterBottom>
              Totales
            </Typography>
            <Divider sx={{ mb: 2 }} />

            <Box mb={2}>
              <Box display="flex" justifyContent="space-between" mb={1}>
                <Typography variant="body2">Subtotal sin impuestos</Typography>
                <Typography variant="body1">
                  {formatCurrency(invoice.total_sin_impuestos || 0)}
                </Typography>
              </Box>
              <Box display="flex" justifyContent="space-between" mb={1}>
                <Typography variant="body2">IVA (12%)</Typography>
                <Typography variant="body1">
                  {formatCurrency(invoice.total_iva || 0)}
                </Typography>
              </Box>
              <Divider sx={{ my: 1 }} />
              <Box display="flex" justifyContent="space-between">
                <Typography variant="h6">Total</Typography>
                <Typography variant="h6" color="primary">
                  {formatCurrency(invoice.total_con_impuestos || 0)}
                </Typography>
              </Box>
            </Box>

            {invoice.clave_acceso && (
              <Button
                fullWidth
                variant="contained"
                startIcon={<DownloadIcon />}
                onClick={handleDownloadPdf}
                disabled={downloadingPdf}
              >
                {downloadingPdf ? 'Descargando...' : 'Descargar PDF'}
              </Button>
              )}
          </Paper>
        </Grid2>
      </Grid2>
    </Box>
  );
};

