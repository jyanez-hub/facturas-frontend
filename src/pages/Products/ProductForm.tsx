import React, { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import {
  Box,
  Button,
  Paper,
  TextField,
  Typography,
  Switch,
  FormControlLabel,
} from '@mui/material';
import { Grid2 } from '@mui/material';
import { useForm } from 'react-hook-form';
import { productsApi } from '../../api/products';
import type { Product } from '../../types';
import { ErrorAlert } from '../../components/common/ErrorAlert';
import { ROUTES } from '../../utils/constants';

export const ProductForm: React.FC = () => {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const isEdit = !!id;

  const {
    register,
    handleSubmit,
    formState: { errors },
    reset,
    watch,
  } = useForm<Product>({
    defaultValues: {
      tiene_iva: true,
    },
  });

  const tieneIva = watch('tiene_iva');

  useEffect(() => {
    if (isEdit) {
      loadProduct();
    }
  }, [id]);

  const loadProduct = async () => {
    if (!id) return;
    try {
      setLoading(true);
      const product = await productsApi.getById(id);
      reset(product);
    } catch (err: any) {
      setError(err.message || 'Error al cargar producto');
    } finally {
      setLoading(false);
    }
  };

  const onSubmit = async (data: Product) => {
    setError('');
    setLoading(true);
    try {
      if (isEdit && id) {
        await productsApi.update(id, data);
      } else {
        await productsApi.create(data);
      }
      navigate(ROUTES.PRODUCTS);
    } catch (err: any) {
      setError(err.message || 'Error al guardar producto');
    } finally {
      setLoading(false);
    }
  };

  return (
    <Box>
      <Typography variant="h4" gutterBottom>
        {isEdit ? 'Editar Producto' : 'Nuevo Producto'}
      </Typography>

      <ErrorAlert
        message={error}
        open={!!error}
        onClose={() => setError('')}
      />

      <Paper sx={{ p: 3, mt: 2 }}>
        <Box component="form" onSubmit={handleSubmit(onSubmit)}>
          <Grid2 container spacing={2}>
            <Grid2 xs={12} sm={6}>
              <TextField
                fullWidth
                label="Código"
                {...register('codigo', {
                  required: 'El código es requerido',
                })}
                error={!!errors.codigo}
                helperText={errors.codigo?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 xs={12} sm={6}>
              <TextField
                fullWidth
                label="Precio Unitario"
                type="number"
                inputProps={{ step: '0.01', min: '0' }}
                {...register('precio_unitario', {
                  required: 'El precio es requerido',
                  valueAsNumber: true,
                  min: { value: 0, message: 'El precio debe ser mayor a 0' },
                })}
                error={!!errors.precio_unitario}
                helperText={errors.precio_unitario?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 xs={12}>
              <TextField
                fullWidth
                label="Descripción"
                {...register('descripcion', {
                  required: 'La descripción es requerida',
                })}
                error={!!errors.descripcion}
                helperText={errors.descripcion?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 xs={12}>
              <TextField
                fullWidth
                label="Descripción Adicional"
                multiline
                rows={2}
                {...register('descripcion_adicional')}
                error={!!errors.descripcion_adicional}
                helperText={errors.descripcion_adicional?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 xs={12}>
              <FormControlLabel
                control={
                  <Switch
                    {...register('tiene_iva')}
                    checked={tieneIva}
                  />
                }
                label="Tiene IVA"
              />
            </Grid2>
            <Grid2 xs={12}>
              <Box display="flex" gap={2} justifyContent="flex-end">
                <Button
                  variant="outlined"
                  onClick={() => navigate(ROUTES.PRODUCTS)}
                  disabled={loading}
                >
                  Cancelar
                </Button>
                <Button
                  type="submit"
                  variant="contained"
                  disabled={loading}
                >
                  {loading ? 'Guardando...' : isEdit ? 'Actualizar' : 'Crear'}
                </Button>
              </Box>
            </Grid2>
          </Grid2>
        </Box>
      </Paper>
    </Box>
  );
};

