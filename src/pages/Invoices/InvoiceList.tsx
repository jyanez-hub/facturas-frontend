import React, { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Box,
  Button,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Typography,
  IconButton,
  Chip,
  Tooltip,
} from '@mui/material';
import {
  Add as AddIcon,
  Visibility as VisibilityIcon,
  Download as DownloadIcon,
} from '@mui/icons-material';
import { invoicesApi, invoicePdfApi } from '../../api/invoices';
import type { Invoice } from '../../types';
import { Loading } from '../../components/common/Loading';
import { ErrorAlert } from '../../components/common/ErrorAlert';
import { formatCurrency, formatDate } from '../../utils/formatters';
import { ROUTES, INVOICE_STATUS_COLORS } from '../../utils/constants';

export const InvoiceList: React.FC = () => {
  const navigate = useNavigate();
  const [invoices, setInvoices] = useState<Invoice[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [downloadingPdf, setDownloadingPdf] = useState<string | null>(null);

  useEffect(() => {
    loadInvoices();
  }, []);

  const loadInvoices = async () => {
    try {
      setLoading(true);
      const data = await invoicesApi.getAll();
      setInvoices(data);
    } catch (err: any) {
      setError(err.message || 'Error al cargar facturas');
    } finally {
      setLoading(false);
    }
  };

  const handleDownloadPdf = async (invoice: Invoice) => {
    if (!invoice.clave_acceso) {
      setError('La factura no tiene clave de acceso');
      return;
    }

    try {
      setDownloadingPdf(invoice._id!);
      const blob = await invoicePdfApi.download(invoice.clave_acceso);
      
      // Crear URL del blob y descargar
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
      setDownloadingPdf(null);
    }
  };

  if (loading) {
    return <Loading />;
  }

  return (
    <Box>
      <Box display="flex" justifyContent="space-between" alignItems="center" mb={3}>
        <Typography variant="h4">Facturas</Typography>
        <Button
          variant="contained"
          startIcon={<AddIcon />}
          onClick={() => navigate(ROUTES.INVOICE_NEW)}
        >
          Nueva Factura
        </Button>
      </Box>

      <ErrorAlert
        message={error}
        open={!!error}
        onClose={() => setError('')}
      />

      <TableContainer component={Paper}>
        <Table>
          <TableHead>
            <TableRow>
              <TableCell>Secuencial</TableCell>
              <TableCell>Clave de Acceso</TableCell>
              <TableCell>Fecha Emisión</TableCell>
              <TableCell>Estado SRI</TableCell>
              <TableCell align="right">Subtotal</TableCell>
              <TableCell align="right">IVA</TableCell>
              <TableCell align="right">Total</TableCell>
              <TableCell align="right">Acciones</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {invoices.length === 0 ? (
              <TableRow>
                <TableCell colSpan={8} align="center">
                  <Typography variant="body2" color="textSecondary" py={3}>
                    No hay facturas registradas
                  </Typography>
                </TableCell>
              </TableRow>
            ) : (
              invoices.map((invoice) => {
                const status = invoice.sri_estado || invoice.estado || 'PENDIENTE';
                const statusColor = 
                  INVOICE_STATUS_COLORS[status as keyof typeof INVOICE_STATUS_COLORS] || 'default';

                return (
                  <TableRow key={invoice._id}>
                    <TableCell>{invoice.secuencial || '-'}</TableCell>
                    <TableCell>
                      <Typography variant="body2" sx={{ fontFamily: 'monospace', fontSize: '0.85rem' }}>
                        {invoice.clave_acceso || '-'}
                      </Typography>
                    </TableCell>
                    <TableCell>{formatDate(invoice.fecha_emision)}</TableCell>
                    <TableCell>
                      <Chip
                        label={status}
                        color={statusColor as any}
                        size="small"
                      />
                    </TableCell>
                    <TableCell align="right">
                      {formatCurrency(invoice.total_sin_impuestos || 0)}
                    </TableCell>
                    <TableCell align="right">
                      {formatCurrency(invoice.total_iva || 0)}
                    </TableCell>
                    <TableCell align="right">
                      <Typography variant="body1" fontWeight="bold">
                        {formatCurrency(invoice.total_con_impuestos || 0)}
                      </Typography>
                    </TableCell>
                    <TableCell align="right">
                      <Tooltip title="Ver detalle">
                        <IconButton
                          size="small"
                          onClick={() => navigate(ROUTES.INVOICE_DETAIL(invoice._id!))}
                        >
                          <VisibilityIcon />
                        </IconButton>
                      </Tooltip>
                      {invoice.clave_acceso && (
                        <Tooltip title="Descargar PDF">
                          <IconButton
                            size="small"
                            onClick={() => handleDownloadPdf(invoice)}
                            disabled={downloadingPdf === invoice._id}
                            color="primary"
                          >
                            <DownloadIcon />
                          </IconButton>
                        </Tooltip>
                      )}
                    </TableCell>
                  </TableRow>
                );
              })
            )}
          </TableBody>
        </Table>
      </TableContainer>
    </Box>
  );
};

