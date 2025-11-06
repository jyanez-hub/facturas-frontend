import { useState, useEffect } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import {
  Box,
  Button,
  Paper,
  TextField,
  Typography,
} from '@mui/material';
import Grid from '../../components/common/Grid2';
import { Save as SaveIcon, ArrowBack as ArrowBackIcon } from '@mui/icons-material';
import { identificationTypeApi } from '../../api/identificationType';
import type { IdentificationType } from '../../types';
import { Loading } from '../../components/common/Loading';
import { ErrorAlert } from '../../components/common/ErrorAlert';

export const IdentificationTypeForm = () => {
  const navigate = useNavigate();
  const { id } = useParams<{ id: string }>();
  const isEditing = Boolean(id);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [formData, setFormData] = useState<Omit<IdentificationType, '_id'>>({
    codigo: '',
    nombre: '',
    descripcion: '',
  });

  useEffect(() => {
    if (isEditing && id) {
      loadType(id);
    }
  }, [id, isEditing]);

  const loadType = async (typeId: string) => {
    try {
      setLoading(true);
      const data = await identificationTypeApi.getById(typeId);
      setFormData({
        codigo: data.codigo,
        nombre: data.nombre,
        descripcion: data.descripcion || '',
      });
    } catch (err: any) {
      setError(err.response?.data?.message || 'Error al cargar tipo de identificación');
    } finally {
      setLoading(false);
    }
  };

  const handleChange = (field: keyof typeof formData) => (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    setFormData((prev) => ({
      ...prev,
      [field]: event.target.value,
    }));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    try {
      setLoading(true);
      if (isEditing && id) {
        await identificationTypeApi.update(id, formData);
      } else {
        await identificationTypeApi.create(formData);
      }
      navigate('/identification-types');
    } catch (err: any) {
      setError(err.response?.data?.message || 'Error al guardar tipo de identificación');
    } finally {
      setLoading(false);
    }
  };

  if (loading && isEditing) return <Loading />;

  return (
    <Box>
      <Box sx={{ display: 'flex', alignItems: 'center', mb: 3 }}>
        <Button
          startIcon={<ArrowBackIcon />}
          onClick={() => navigate('/identification-types')}
          sx={{ mr: 2 }}
        >
          Volver
        </Button>
        <Typography variant="h4" component="h1">
          {isEditing ? 'Editar Tipo de Identificación' : 'Nuevo Tipo de Identificación'}
        </Typography>
      </Box>

      <ErrorAlert message={error || ''} open={!!error} onClose={() => setError(null)} />

      <Paper sx={{ p: 3 }}>
        <form onSubmit={handleSubmit}>
          <Grid container spacing={3}>
            <Grid item xs={12} md={3}>
              <TextField
                fullWidth
                label="Código"
                value={formData.codigo}
                onChange={handleChange('codigo')}
                required
                disabled={isEditing}
                helperText={isEditing ? 'El código no se puede modificar' : 'Código SRI (ej: 05 para cédula)'}
              />
            </Grid>

            <Grid item xs={12} md={9}>
              <TextField
                fullWidth
                label="Nombre"
                value={formData.nombre}
                onChange={handleChange('nombre')}
                required
                helperText="Nombre del tipo de identificación"
              />
            </Grid>

            <Grid item xs={12}>
              <TextField
                fullWidth
                label="Descripción"
                value={formData.descripcion}
                onChange={handleChange('descripcion')}
                multiline
                rows={3}
                helperText="Descripción opcional del tipo de identificación"
              />
            </Grid>

            <Grid item xs={12}>
              <Box sx={{ display: 'flex', gap: 2, justifyContent: 'flex-end' }}>
                <Button
                  variant="outlined"
                  onClick={() => navigate('/identification-types')}
                  disabled={loading}
                >
                  Cancelar
                </Button>
                <Button
                  type="submit"
                  variant="contained"
                  color="primary"
                  startIcon={<SaveIcon />}
                  disabled={loading}
                >
                  {loading ? 'Guardando...' : 'Guardar'}
                </Button>
              </Box>
            </Grid>
          </Grid>
        </form>
      </Paper>
    </Box>
  );
};

