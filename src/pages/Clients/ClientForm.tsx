import React, { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import {
  Box,
  Button,
  Paper,
  TextField,
  Typography,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  FormHelperText,
} from '@mui/material';
import Grid2 from '../../components/common/Grid2';
import { useForm, Controller } from 'react-hook-form';
import { clientsApi } from '../../api/clients';
import { identificationTypeApi } from '../../api/identificationType';
import type { Client, IdentificationType } from '../../types';
import { ErrorAlert } from '../../components/common/ErrorAlert';
import { ROUTES } from '../../utils/constants';

export const ClientForm: React.FC = () => {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [identificationTypes, setIdentificationTypes] = useState<IdentificationType[]>([]);
  const isEdit = !!id;

  const {
    register,
    handleSubmit,
    formState: { errors },
    reset,
    control,
  } = useForm<Client>();

  useEffect(() => {
    loadIdentificationTypes();
    if (isEdit) {
      loadClient();
    }
  }, [id]);

  const loadIdentificationTypes = async () => {
    try {
      const types = await identificationTypeApi.getAll();
      setIdentificationTypes(types);
    } catch (err) {
      console.error('Error loading identification types:', err);
    }
  };

  const loadClient = async () => {
    if (!id) return;
    try {
      setLoading(true);
      const client = await clientsApi.getById(id);
      reset(client);
    } catch (err: any) {
      setError(err.message || 'Error al cargar cliente');
    } finally {
      setLoading(false);
    }
  };

  const onSubmit = async (data: Client) => {
    setError('');
    setLoading(true);
    try {
      if (isEdit && id) {
        await clientsApi.update(id, data);
      } else {
        await clientsApi.create(data);
      }
      navigate(ROUTES.CLIENTS);
    } catch (err: any) {
      setError(err.message || 'Error al guardar cliente');
    } finally {
      setLoading(false);
    }
  };

  return (
    <Box>
      <Typography variant="h4" gutterBottom>
        {isEdit ? 'Editar Cliente' : 'Nuevo Cliente'}
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
              <FormControl fullWidth error={!!errors.tipo_identificacion_id} disabled={loading}>
                <InputLabel>Tipo de Identificación *</InputLabel>
                <Controller
                  name="tipo_identificacion_id"
                  control={control}
                  defaultValue=""
                  rules={{ required: 'El tipo de identificación es requerido' }}
                  render={({ field }) => (
                    <Select {...field} label="Tipo de Identificación *">
                      {identificationTypes.length === 0 ? (
                        <MenuItem value="">
                          <em>No hay tipos disponibles</em>
                        </MenuItem>
                      ) : (
                        identificationTypes.map((type) => (
                          <MenuItem key={type._id} value={type._id}>
                            {type.codigo} - {type.nombre}
                          </MenuItem>
                        ))
                      )}
                    </Select>
                  )}
                />
                {errors.tipo_identificacion_id && (
                  <FormHelperText>{errors.tipo_identificacion_id.message}</FormHelperText>
                )}
                {identificationTypes.length === 0 && (
                  <FormHelperText>Primero debes crear un tipo de identificación</FormHelperText>
                )}
              </FormControl>
            </Grid2>
            <Grid2 item xs={12} sm={6}>
              <TextField
                fullWidth
                label="Identificación"
                {...register('identificacion', {
                  required: 'La identificación es requerida',
                })}
                error={!!errors.identificacion}
                helperText={errors.identificacion?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 item xs={12}>
              <TextField
                fullWidth
                label="Razón Social"
                {...register('razon_social', {
                  required: 'La razón social es requerida',
                })}
                error={!!errors.razon_social}
                helperText={errors.razon_social?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 item xs={12} sm={6}>
              <TextField
                fullWidth
                label="Email"
                type="email"
                {...register('email')}
                error={!!errors.email}
                helperText={errors.email?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 item xs={12} sm={6}>
              <TextField
                fullWidth
                label="Teléfono"
                {...register('telefono')}
                error={!!errors.telefono}
                helperText={errors.telefono?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 item xs={12}>
              <TextField
                fullWidth
                label="Dirección"
                multiline
                rows={2}
                {...register('direccion')}
                error={!!errors.direccion}
                helperText={errors.direccion?.message}
                disabled={loading}
              />
            </Grid2>
            <Grid2 item xs={12}>
              <Box display="flex" gap={2} justifyContent="flex-end">
                <Button
                  variant="outlined"
                  onClick={() => navigate(ROUTES.CLIENTS)}
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

