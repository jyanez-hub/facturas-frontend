import React, { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Box,
  Button,
  Paper,
  Typography,
  TextField,
  Autocomplete,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  IconButton,
  Alert,
  Chip,
} from '@mui/material';
import Grid2 from '../../components/common/Grid2';
import {
  Add as AddIcon,
  Delete as DeleteIcon,
} from '@mui/icons-material';
import { useForm, useFieldArray } from 'react-hook-form';
import { invoicesApi } from '../../api/invoices';
import { clientsApi } from '../../api/clients';
import { productsApi } from '../../api/products';
import type { Client, Product } from '../../types';
import { ErrorAlert } from '../../components/common/ErrorAlert';
import { Loading } from '../../components/common/Loading';
import { formatCurrency } from '../../utils/formatters';
import { ROUTES } from '../../utils/constants';
import { useAuth } from '../../contexts/AuthContext';
import apiClient from '../../api/client';

interface InvoiceDetailForm {
  producto_id: string;
  cantidad: number;
  precio_unitario: number;
  descuento?: number;
}

interface InvoiceFormData {
  cliente_id: string;
  fecha_emision: string;
  detalles: InvoiceDetailForm[];
}

export const InvoiceForm: React.FC = () => {
  const navigate = useNavigate();
  const { company } = useAuth();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [clients, setClients] = useState<Client[]>([]);
  const [products, setProducts] = useState<Product[]>([]);
  const [loadingData, setLoadingData] = useState(true);

  const {
    register,
    handleSubmit,
    control,
    watch,
    setValue,
    formState: { errors },
  } = useForm<InvoiceFormData>({
    defaultValues: {
      fecha_emision: new Date().toISOString().split('T')[0],
      detalles: [],
    },
  });

  const { fields, append, remove } = useFieldArray({
    control,
    name: 'detalles',
  });

  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    try {
      setLoadingData(true);
      const [clientsData, productsData] = await Promise.all([
        clientsApi.getAll(),
        productsApi.getAll(),
      ]);
      setClients(clientsData);
      setProducts(productsData);
    } catch (err: any) {
      setError(err.message || 'Error al cargar datos');
    } finally {
      setLoadingData(false);
    }
  };

  const detalles = watch('detalles');

  const calculateTotals = () => {
    let totalSinImpuestos = 0;
    let totalIva = 0;

    detalles.forEach((detalle) => {
      const product = products.find((p) => p._id === detalle.producto_id);
      if (product) {
        const subtotal = detalle.cantidad * detalle.precio_unitario;
        const descuento = detalle.descuento || 0;
        const subtotalConDescuento = subtotal - descuento;
        totalSinImpuestos += subtotalConDescuento;
        
        if (product.tiene_iva) {
          const iva = subtotalConDescuento * 0.12; // IVA 12%
          totalIva += iva;
        }
      }
    });

    return {
      totalSinImpuestos: Math.round(totalSinImpuestos * 100) / 100,
      totalIva: Math.round(totalIva * 100) / 100,
      totalConImpuestos: Math.round((totalSinImpuestos + totalIva) * 100) / 100,
    };
  };

  const totals = calculateTotals();

  const addProduct = () => {
    append({
      producto_id: '',
      cantidad: 1,
      precio_unitario: 0,
      descuento: 0,
    });
  };

  const removeProduct = (index: number) => {
    remove(index);
  };

  const updateProductPrice = (index: number, productId: string) => {
    const product = products.find((p) => p._id === productId);
    if (product) {
      setValue(`detalles.${index}.precio_unitario`, product.precio_unitario);
    }
  };

  const onSubmit = async (data: InvoiceFormData) => {
    if (data.detalles.length === 0) {
      setError('Debe agregar al menos un producto');
      return;
    }

    setError('');
    setLoading(true);

    try {
      // Obtener datos del cliente y empresa
      const client = clients.find((c) => c._id === data.cliente_id);
      if (!client) {
        throw new Error('Cliente no encontrado');
      }

      // Obtener empresa emisora
      const companyResponse = await apiClient.get('/api/v1/issuing-company');
      const companies = companyResponse.data;
      const issuingCompany = companies.find((c: any) => c._id === company?.id || c.ruc === company?.id);
      
      if (!issuingCompany) {
        throw new Error('Empresa emisora no encontrada');
      }

      // Construir estructura InvoiceRequest
      const fechaEmision = new Date(data.fecha_emision).toISOString().split('T')[0];
      
      // Obtener tipo de identificación del cliente
      const identificationTypeResponse = await apiClient.get(`/api/v1/identification-type/${client.tipo_identificacion_id}`);
      const identificationType = identificationTypeResponse.data;

      const detallesFormatted = data.detalles.map((detalle) => {
        const product = products.find((p) => p._id === detalle.producto_id);
        if (!product) throw new Error('Producto no encontrado');

        const subtotal = detalle.cantidad * detalle.precio_unitario;
        const descuento = detalle.descuento || 0;
        const subtotalConDescuento = subtotal - descuento;
        const iva = product.tiene_iva ? subtotalConDescuento * 0.12 : 0;

        return {
          detalle: {
            codigoPrincipal: product.codigo,
            descripcion: product.descripcion,
            cantidad: detalle.cantidad.toString(),
            precioUnitario: detalle.precio_unitario.toFixed(2),
            precioTotalSinImpuesto: subtotalConDescuento.toFixed(2),
            impuestos: product.tiene_iva
              ? [
                  {
                    impuesto: {
                      codigo: '2',
                      codigoPorcentaje: '2',
                      tarifa: '12.00',
                      baseImponible: subtotalConDescuento.toFixed(2),
                      valor: iva.toFixed(2),
                    },
                  },
                ]
              : [],
          },
        };
      });

      const totalSinImpuestos = detallesFormatted.reduce(
        (sum, d) => sum + parseFloat(d.detalle.precioTotalSinImpuesto),
        0
      );
      const totalIva = detallesFormatted.reduce(
        (sum, d) => sum + (d.detalle.impuestos[0] ? parseFloat(d.detalle.impuestos[0].impuesto.valor) : 0),
        0
      );
      const importeTotal = totalSinImpuestos + totalIva;

      const invoiceRequest = {
        infoTributaria: {
          ruc: issuingCompany.ruc,
          claveAcceso: '', // Se generará en el backend
          secuencial: '', // Se generará en el backend
        },
        infoFactura: {
          fechaEmision,
          tipoIdentificacionComprador: identificationType.codigo,
          identificacionComprador: client.identificacion,
          razonSocialComprador: client.razon_social,
          totalSinImpuestos: totalSinImpuestos.toFixed(2),
          importeTotal: importeTotal.toFixed(2),
        },
        detalles: detallesFormatted,
      };

      await invoicesApi.createComplete({ factura: invoiceRequest });
      navigate(ROUTES.INVOICES);
    } catch (err: any) {
      setError(err.response?.data?.message || err.message || 'Error al crear factura');
    } finally {
      setLoading(false);
    }
  };

  if (loadingData) {
    return <Loading />;
  }

  return (
    <Box>
      <Typography variant="h4" gutterBottom>
        Nueva Factura
      </Typography>

      <ErrorAlert
        message={error}
        open={!!error}
        onClose={() => setError('')}
      />

      <Paper sx={{ p: 3, mt: 2 }}>
        <Box component="form" onSubmit={handleSubmit(onSubmit)}>
          <Grid2 container spacing={2}>
            <Grid2 item xs={12} sm={6}>
              <Autocomplete
                options={clients}
                getOptionLabel={(option) => `${option.identificacion} - ${option.razon_social}`}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    label="Cliente"
                    required
                    error={!!errors.cliente_id}
                    helperText={errors.cliente_id?.message}
                  />
                )}
                onChange={(_, value) => {
                  setValue('cliente_id', value?._id || '');
                }}
                disabled={loading}
              />
            </Grid2>
            <Grid2 item xs={12} sm={6}>
              <TextField
                fullWidth
                label="Fecha de Emisión"
                type="date"
                {...register('fecha_emision', {
                  required: 'La fecha es requerida',
                })}
                InputLabelProps={{
                  shrink: true,
                }}
                error={!!errors.fecha_emision}
                helperText={errors.fecha_emision?.message}
                disabled={loading}
              />
            </Grid2>

            <Grid2 item xs={12}>
              <Box display="flex" justifyContent="space-between" alignItems="center" mb={2}>
                <Typography variant="h6">Productos</Typography>
                <Button
                  variant="outlined"
                  startIcon={<AddIcon />}
                  onClick={addProduct}
                  disabled={loading}
                >
                  Agregar Producto
                </Button>
              </Box>

              {fields.length === 0 ? (
                <Alert severity="info">Agregue al menos un producto a la factura</Alert>
              ) : (
                <TableContainer>
                  <Table>
                    <TableHead>
                      <TableRow>
                        <TableCell>Producto</TableCell>
                        <TableCell align="right">Cantidad</TableCell>
                        <TableCell align="right">Precio Unit.</TableCell>
                        <TableCell align="right">Descuento</TableCell>
                        <TableCell align="right">Subtotal</TableCell>
                        <TableCell align="right">IVA</TableCell>
                        <TableCell align="right">Total</TableCell>
                        <TableCell align="right">Acciones</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {fields.map((field, index) => {
                        const product = products.find(
                          (p) => p._id === watch(`detalles.${index}.producto_id`)
                        );
                        const cantidad = watch(`detalles.${index}.cantidad`) || 0;
                        const precio = watch(`detalles.${index}.precio_unitario`) || 0;
                        const descuento = watch(`detalles.${index}.descuento`) || 0;
                        const subtotal = cantidad * precio;
                        const subtotalConDescuento = subtotal - descuento;
                        const iva = product?.tiene_iva ? subtotalConDescuento * 0.12 : 0;
                        const total = subtotalConDescuento + iva;

                        return (
                          <TableRow key={field.id}>
                            <TableCell>
                              <Autocomplete
                                options={products}
                                getOptionLabel={(option) => `${option.codigo} - ${option.descripcion}`}
                                renderInput={(params) => (
                                  <TextField
                                    {...params}
                                    size="small"
                                    placeholder="Seleccionar producto"
                                  />
                                )}
                                onChange={(_, value) => {
                                  setValue(`detalles.${index}.producto_id`, value?._id || '');
                                  if (value) {
                                    updateProductPrice(index, value._id!);
                                  }
                                }}
                                disabled={loading}
                              />
                            </TableCell>
                            <TableCell>
                              <TextField
                                type="number"
                                size="small"
                                inputProps={{ min: 1, step: 1 }}
                                {...register(`detalles.${index}.cantidad`, {
                                  required: true,
                                  min: 1,
                                  valueAsNumber: true,
                                })}
                                disabled={loading}
                              />
                            </TableCell>
                            <TableCell align="right">
                              <TextField
                                type="number"
                                size="small"
                                inputProps={{ min: 0, step: 0.01 }}
                                {...register(`detalles.${index}.precio_unitario`, {
                                  required: true,
                                  min: 0,
                                  valueAsNumber: true,
                                })}
                                disabled={loading}
                              />
                            </TableCell>
                            <TableCell align="right">
                              <TextField
                                type="number"
                                size="small"
                                inputProps={{ min: 0, step: 0.01 }}
                                {...register(`detalles.${index}.descuento`, {
                                  valueAsNumber: true,
                                })}
                                disabled={loading}
                              />
                            </TableCell>
                            <TableCell align="right">
                              {formatCurrency(subtotalConDescuento)}
                            </TableCell>
                            <TableCell align="right">
                              {product?.tiene_iva ? (
                                <Chip label={formatCurrency(iva)} size="small" color="info" />
                              ) : (
                                <Chip label="Sin IVA" size="small" />
                              )}
                            </TableCell>
                            <TableCell align="right">
                              {formatCurrency(total)}
                            </TableCell>
                            <TableCell align="right">
                              <IconButton
                                size="small"
                                color="error"
                                onClick={() => removeProduct(index)}
                                disabled={loading}
                              >
                                <DeleteIcon />
                              </IconButton>
                            </TableCell>
                          </TableRow>
                        );
                      })}
                    </TableBody>
                  </Table>
                </TableContainer>
              )}
            </Grid2>

            <Grid2 item xs={12}>
              <Paper sx={{ p: 2, bgcolor: 'grey.50' }}>
                <Grid2 container spacing={2}>
                  <Grid2 item xs={12} sm={6}>
                    <Typography variant="body2" color="textSecondary">
                      Subtotal sin impuestos
                    </Typography>
                    <Typography variant="h6">
                      {formatCurrency(totals.totalSinImpuestos)}
                    </Typography>
                  </Grid2>
                  <Grid2 item xs={12} sm={6}>
                    <Typography variant="body2" color="textSecondary">
                      IVA (12%)
                    </Typography>
                    <Typography variant="h6">
                      {formatCurrency(totals.totalIva)}
                    </Typography>
                  </Grid2>
                  <Grid2 item xs={12}>
                    <Box display="flex" justifyContent="space-between" alignItems="center" pt={1}>
                      <Typography variant="h6">TOTAL</Typography>
                      <Typography variant="h5" color="primary">
                        {formatCurrency(totals.totalConImpuestos)}
                      </Typography>
                    </Box>
                  </Grid2>
                </Grid2>
              </Paper>
            </Grid2>

            <Grid2 item xs={12}>
              <Box display="flex" gap={2} justifyContent="flex-end">
                <Button
                  variant="outlined"
                  onClick={() => navigate(ROUTES.INVOICES)}
                  disabled={loading}
                >
                  Cancelar
                </Button>
                <Button
                  type="submit"
                  variant="contained"
                  disabled={loading || fields.length === 0}
                >
                  {loading ? 'Creando...' : 'Crear Factura'}
                </Button>
              </Box>
            </Grid2>
          </Grid2>
        </Box>
      </Paper>
    </Box>
  );
};

