import React, { useEffect, useState } from 'react';
import {
  Paper,
  Typography,
  Box,
  Card,
  CardContent,
} from '@mui/material';
import Grid2 from '../components/common/Grid2';
import {
  Receipt as ReceiptIcon,
  TrendingUp as TrendingUpIcon,
  AttachMoney as AttachMoneyIcon,
} from '@mui/icons-material';
import { invoicesApi } from '../api/invoices';
import type { Invoice } from '../types';
import { formatCurrency, formatDate } from '../utils/formatters';
import { Loading } from '../components/common/Loading';

export const Dashboard: React.FC = () => {
  const [invoices, setInvoices] = useState<Invoice[]>([]);
  const [loading, setLoading] = useState(true);
  const [stats, setStats] = useState({
    total: 0,
    thisMonth: 0,
    totalAmount: 0,
  });

  useEffect(() => {
    loadInvoices();
  }, []);

  const loadInvoices = async () => {
    try {
      setLoading(true);
      const data = await invoicesApi.getAll();
      setInvoices(data);

      // Calcular estadísticas
      const now = new Date();
      const thisMonthInvoices = data.filter((inv) => {
        const invoiceDate = new Date(inv.fecha_emision);
        return (
          invoiceDate.getMonth() === now.getMonth() &&
          invoiceDate.getFullYear() === now.getFullYear()
        );
      });

      const totalAmount = data.reduce(
        (sum, inv) => sum + (inv.total_con_impuestos || 0),
        0
      );

      setStats({
        total: data.length,
        thisMonth: thisMonthInvoices.length,
        totalAmount,
      });
    } catch (error) {
      console.error('Error loading invoices:', error);
    } finally {
      setLoading(false);
    }
  };

  if (loading) {
    return <Loading />;
  }

  return (
    <Box>
      <Typography variant="h4" gutterBottom>
        Dashboard
      </Typography>

      <Grid2 container spacing={3} sx={{ mt: 2 }}>
        {/* Tarjetas de estadísticas */}
        <Grid2 item xs={12} sm={4}>
          <Card>
            <CardContent>
              <Box display="flex" alignItems="center" justifyContent="space-between">
                <Box>
                  <Typography color="textSecondary" gutterBottom variant="body2">
                    Total Facturas
                  </Typography>
                  <Typography variant="h4">{stats.total}</Typography>
                </Box>
                <ReceiptIcon sx={{ fontSize: 40, color: 'primary.main' }} />
              </Box>
            </CardContent>
          </Card>
        </Grid2>

        <Grid2 item xs={12} sm={4}>
          <Card>
            <CardContent>
              <Box display="flex" alignItems="center" justifyContent="space-between">
                <Box>
                  <Typography color="textSecondary" gutterBottom variant="body2">
                    Este Mes
                  </Typography>
                  <Typography variant="h4">{stats.thisMonth}</Typography>
                </Box>
                <TrendingUpIcon sx={{ fontSize: 40, color: 'success.main' }} />
              </Box>
            </CardContent>
          </Card>
        </Grid2>

        <Grid2 item xs={12} sm={4}>
          <Card>
            <CardContent>
              <Box display="flex" alignItems="center" justifyContent="space-between">
                <Box>
                  <Typography color="textSecondary" gutterBottom variant="body2">
                    Total Facturado
                  </Typography>
                  <Typography variant="h4">{formatCurrency(stats.totalAmount)}</Typography>
                </Box>
                <AttachMoneyIcon sx={{ fontSize: 40, color: 'warning.main' }} />
              </Box>
            </CardContent>
          </Card>
        </Grid2>

        {/* Facturas recientes */}
        <Grid2 item xs={12}>
          <Paper sx={{ p: 2 }}>
            <Typography variant="h6" gutterBottom>
              Facturas Recientes
            </Typography>
            {invoices.length === 0 ? (
              <Typography variant="body2" color="textSecondary">
                No hay facturas registradas
              </Typography>
            ) : (
              <Box sx={{ mt: 2 }}>
                {invoices.slice(0, 5).map((invoice) => (
                  <Box
                    key={invoice._id}
                    sx={{
                      display: 'flex',
                      justifyContent: 'space-between',
                      alignItems: 'center',
                      py: 1,
                      borderBottom: '1px solid',
                      borderColor: 'divider',
                    }}
                  >
                    <Box>
                      <Typography variant="body1">
                        Factura #{invoice.secuencial || 'N/A'}
                      </Typography>
                      <Typography variant="body2" color="textSecondary">
                        {formatDate(invoice.fecha_emision)}
                      </Typography>
                    </Box>
                    <Box textAlign="right">
                      <Typography variant="body1" fontWeight="bold">
                        {formatCurrency(invoice.total_con_impuestos || 0)}
                      </Typography>
                      <Typography
                        variant="body2"
                        color={
                          invoice.sri_estado === 'RECIBIDA'
                            ? 'success.main'
                            : invoice.sri_estado === 'DEVUELTA'
                            ? 'error.main'
                            : 'text.secondary'
                        }
                      >
                        {invoice.sri_estado || invoice.estado || 'PENDIENTE'}
                      </Typography>
                    </Box>
                  </Box>
                ))}
              </Box>
            )}
          </Paper>
        </Grid2>
      </Grid2>
    </Box>
  );
};

