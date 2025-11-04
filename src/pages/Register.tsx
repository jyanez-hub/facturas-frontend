import React, { useState } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import {
  Container,
  Paper,
  TextField,
  Button,
  Typography,
  Box,
  Alert,
} from '@mui/material';
import Grid2 from '../components/common/Grid2';
import { useForm } from 'react-hook-form';
import { useAuth } from '../contexts/AuthContext';
import { ROUTES } from '../utils/constants';
import { isValidRUC } from '../utils/formatters';

interface RegisterFormData {
  email: string;
  password: string;
  confirmPassword: string;
  masterKey?: string;
  invitationCode?: string;
  ruc: string;
  razon_social: string;
  nombre_comercial: string;
  direccion?: string;
  telefono?: string;
  company_email?: string;
}

export const Register: React.FC = () => {
  const navigate = useNavigate();
  const { register: registerUser } = useAuth();
  const [error, setError] = useState<string>('');
  const [loading, setLoading] = useState(false);

  const {
    register,
    handleSubmit,
    watch,
    formState: { errors },
  } = useForm<RegisterFormData>();

  const password = watch('password');

  const onSubmit = async (data: RegisterFormData) => {
    setError('');
    
    // Validar RUC
    if (!isValidRUC(data.ruc)) {
      setError('El RUC debe tener 13 dígitos y terminar en 001');
      return;
    }

    // Validar contraseñas
    if (data.password !== data.confirmPassword) {
      setError('Las contraseñas no coinciden');
      return;
    }

    setLoading(true);
    try {
      const { confirmPassword, ...registerData } = data;
      await registerUser(registerData);
      navigate(ROUTES.DASHBOARD);
    } catch (err: any) {
      setError(err.message || 'Error al registrar usuario');
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container maxWidth="md">
      <Box
        sx={{
          marginTop: 4,
          marginBottom: 4,
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
        }}
      >
        <Paper elevation={3} sx={{ padding: 4, width: '100%' }}>
          <Typography component="h1" variant="h4" align="center" gutterBottom>
            Registro de Usuario
          </Typography>
          <Typography variant="body2" align="center" color="text.secondary" sx={{ mb: 3 }}>
            Crea tu cuenta y empresa emisora
          </Typography>

          {error && (
            <Alert severity="error" sx={{ mb: 2 }}>
              {error}
            </Alert>
          )}

          <Box component="form" onSubmit={handleSubmit(onSubmit)} sx={{ mt: 1 }}>
            <Grid2 container spacing={2}>
              <Grid2 item xs={12}>
                <Typography variant="h6" gutterBottom>
                  Datos de Usuario
                </Typography>
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  required
                  fullWidth
                  id="email"
                  label="Correo Electrónico"
                  autoComplete="email"
                  {...register('email', {
                    required: 'El correo es requerido',
                    pattern: {
                      value: /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i,
                      message: 'Correo electrónico inválido',
                    },
                  })}
                  error={!!errors.email}
                  helperText={errors.email?.message}
                />
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  required
                  fullWidth
                  label="Contraseña"
                  type="password"
                  id="password"
                  {...register('password', {
                    required: 'La contraseña es requerida',
                    minLength: {
                      value: 6,
                      message: 'La contraseña debe tener al menos 6 caracteres',
                    },
                  })}
                  error={!!errors.password}
                  helperText={errors.password?.message}
                />
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  required
                  fullWidth
                  label="Confirmar Contraseña"
                  type="password"
                  id="confirmPassword"
                  {...register('confirmPassword', {
                    required: 'Confirma tu contraseña',
                    validate: (value) =>
                      value === password || 'Las contraseñas no coinciden',
                  })}
                  error={!!errors.confirmPassword}
                  helperText={errors.confirmPassword?.message}
                />
              </Grid2>

              <Grid2 item xs={12}>
                <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
                  Datos de la Empresa
                </Typography>
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  required
                  fullWidth
                  id="ruc"
                  label="RUC"
                  {...register('ruc', {
                    required: 'El RUC es requerido',
                    pattern: {
                      value: /^\d{10}001$/,
                      message: 'RUC debe tener 13 dígitos y terminar en 001',
                    },
                  })}
                  error={!!errors.ruc}
                  helperText={errors.ruc?.message || 'Formato: 1234567890001'}
                />
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  required
                  fullWidth
                  id="razon_social"
                  label="Razón Social"
                  {...register('razon_social', {
                    required: 'La razón social es requerida',
                  })}
                  error={!!errors.razon_social}
                  helperText={errors.razon_social?.message}
                />
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  fullWidth
                  id="nombre_comercial"
                  label="Nombre Comercial"
                  {...register('nombre_comercial')}
                />
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  fullWidth
                  id="direccion"
                  label="Dirección"
                  {...register('direccion')}
                />
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  fullWidth
                  id="telefono"
                  label="Teléfono"
                  {...register('telefono')}
                />
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  fullWidth
                  id="company_email"
                  label="Email de la Empresa"
                  {...register('company_email')}
                />
              </Grid2>

              <Grid2 item xs={12}>
                <Typography variant="body2" color="text.secondary" gutterBottom>
                  Para el primer registro, necesitarás la clave maestra. Para registros posteriores,
                  usa un código de invitación.
                </Typography>
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  fullWidth
                  id="masterKey"
                  label="Clave Maestra (Primer registro)"
                  type="password"
                  {...register('masterKey')}
                />
              </Grid2>
              <Grid2 item xs={12} sm={6}>
                <TextField
                  fullWidth
                  id="invitationCode"
                  label="Código de Invitación"
                  {...register('invitationCode')}
                />
              </Grid2>
            </Grid2>

            <Button
              type="submit"
              fullWidth
              variant="contained"
              sx={{ mt: 3, mb: 2 }}
              disabled={loading}
            >
              {loading ? 'Registrando...' : 'Registrarse'}
            </Button>
            <Box textAlign="center">
              <Link to={ROUTES.LOGIN} style={{ textDecoration: 'none' }}>
                <Typography variant="body2" color="primary">
                  ¿Ya tienes cuenta? Inicia sesión aquí
                </Typography>
              </Link>
            </Box>
          </Box>
        </Paper>
      </Box>
    </Container>
  );
};

