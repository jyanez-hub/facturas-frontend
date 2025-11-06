import { useState, useEffect } from 'react';
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
  FormControlLabel,
  Switch,
} from '@mui/material';
import Grid from '../../components/common/Grid2';
import { Save as SaveIcon, ArrowBack as ArrowBackIcon } from '@mui/icons-material';
import { issuingCompanyApi } from '../../api/issuingCompany';
import type { IssuingCompany } from '../../types';
import { Loading } from '../../components/common/Loading';
import { ErrorAlert } from '../../components/common/ErrorAlert';

export const IssuingCompanyForm = () => {
  const navigate = useNavigate();
  const { id } = useParams<{ id: string }>();

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [formData, setFormData] = useState<Partial<IssuingCompany>>({
    ruc: '',
    razon_social: '',
    nombre_comercial: '',
    direccion: '',
    direccion_matriz: '',
    direccion_establecimiento: '',
    telefono: '',
    email: '',
    codigo_establecimiento: '',
    punto_emision: '',
    tipo_ambiente: 1,
    tipo_emision: 1,
    obligado_contabilidad: false,
  });

  useEffect(() => {
    if (id) {
      loadCompany(id);
    }
  }, [id]);

  const loadCompany = async (companyId: string) => {
    try {
      setLoading(true);
      const data = await issuingCompanyApi.getById(companyId);
      setFormData(data);
    } catch (err: any) {
      setError(err.response?.data?.message || 'Error al cargar empresa emisora');
    } finally {
      setLoading(false);
    }
  };

  const handleChange = (field: keyof IssuingCompany) => (
    event: React.ChangeEvent<HTMLInputElement | { value: unknown }>
  ) => {
    const value = event.target.value;
    setFormData((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleSwitchChange = (field: keyof IssuingCompany) => (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    setFormData((prev) => ({
      ...prev,
      [field]: event.target.checked,
    }));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    if (!id) return;

    try {
      setLoading(true);
      await issuingCompanyApi.update(id, formData);
      navigate('/issuing-company');
    } catch (err: any) {
      setError(err.response?.data?.message || 'Error al actualizar empresa emisora');
    } finally {
      setLoading(false);
    }
  };

  if (loading && !formData.ruc) return <Loading />;

  return (
    <Box>
      <Box sx={{ display: 'flex', alignItems: 'center', mb: 3 }}>
        <Button
          startIcon={<ArrowBackIcon />}
          onClick={() => navigate('/issuing-company')}
          sx={{ mr: 2 }}
        >
          Volver
        </Button>
        <Typography variant="h4" component="h1">
          Editar Empresa Emisora
        </Typography>
      </Box>

      <ErrorAlert message={error || ''} open={!!error} onClose={() => setError(null)} />

      <Paper sx={{ p: 3 }}>
        <form onSubmit={handleSubmit}>
          <Grid container spacing={3}>
            <Grid item xs={12}>
              <Typography variant="h6" gutterBottom>
                Información General
              </Typography>
            </Grid>

            <Grid item xs={12} md={4}>
              <TextField
                fullWidth
                label="RUC"
                value={formData.ruc || ''}
                disabled
                helperText="El RUC no se puede modificar"
              />
            </Grid>

            <Grid item xs={12} md={8}>
              <TextField
                fullWidth
                label="Razón Social"
                value={formData.razon_social || ''}
                onChange={handleChange('razon_social')}
                required
              />
            </Grid>

            <Grid item xs={12} md={6}>
              <TextField
                fullWidth
                label="Nombre Comercial"
                value={formData.nombre_comercial || ''}
                onChange={handleChange('nombre_comercial')}
              />
            </Grid>

            <Grid item xs={12} md={6}>
              <TextField
                fullWidth
                label="Email"
                type="email"
                value={formData.email || ''}
                onChange={handleChange('email')}
              />
            </Grid>

            <Grid item xs={12} md={6}>
              <TextField
                fullWidth
                label="Teléfono"
                value={formData.telefono || ''}
                onChange={handleChange('telefono')}
              />
            </Grid>

            <Grid item xs={12}>
              <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
                Direcciones
              </Typography>
            </Grid>

            <Grid item xs={12}>
              <TextField
                fullWidth
                label="Dirección Principal"
                value={formData.direccion || ''}
                onChange={handleChange('direccion')}
                multiline
                rows={2}
              />
            </Grid>

            <Grid item xs={12} md={6}>
              <TextField
                fullWidth
                label="Dirección Matriz"
                value={formData.direccion_matriz || ''}
                onChange={handleChange('direccion_matriz')}
                multiline
                rows={2}
              />
            </Grid>

            <Grid item xs={12} md={6}>
              <TextField
                fullWidth
                label="Dirección Establecimiento"
                value={formData.direccion_establecimiento || ''}
                onChange={handleChange('direccion_establecimiento')}
                multiline
                rows={2}
              />
            </Grid>

            <Grid item xs={12}>
              <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
                Configuración Fiscal
              </Typography>
            </Grid>

            <Grid item xs={12} md={4}>
              <TextField
                fullWidth
                label="Código Establecimiento"
                value={formData.codigo_establecimiento || ''}
                onChange={handleChange('codigo_establecimiento')}
                helperText="Ejemplo: 001"
              />
            </Grid>

            <Grid item xs={12} md={4}>
              <TextField
                fullWidth
                label="Punto de Emisión"
                value={formData.punto_emision || ''}
                onChange={handleChange('punto_emision')}
                helperText="Ejemplo: 001"
              />
            </Grid>

            <Grid item xs={12} md={4}>
              <FormControl fullWidth>
                <InputLabel>Tipo de Ambiente</InputLabel>
                <Select
                  value={formData.tipo_ambiente || 1}
                  label="Tipo de Ambiente"
                  onChange={(e) => setFormData(prev => ({ ...prev, tipo_ambiente: Number(e.target.value) }))}
                >
                  <MenuItem value={1}>Pruebas</MenuItem>
                  <MenuItem value={2}>Producción</MenuItem>
                </Select>
              </FormControl>
            </Grid>

            <Grid item xs={12} md={6}>
              <FormControl fullWidth>
                <InputLabel>Tipo de Emisión</InputLabel>
                <Select
                  value={formData.tipo_emision || 1}
                  label="Tipo de Emisión"
                  onChange={(e) => setFormData(prev => ({ ...prev, tipo_emision: Number(e.target.value) }))}
                >
                  <MenuItem value={1}>Normal</MenuItem>
                  <MenuItem value={2}>Contingencia</MenuItem>
                </Select>
              </FormControl>
            </Grid>

            <Grid item xs={12} md={6}>
              <FormControlLabel
                control={
                  <Switch
                    checked={formData.obligado_contabilidad || false}
                    onChange={handleSwitchChange('obligado_contabilidad')}
                  />
                }
                label="Obligado a llevar contabilidad"
              />
            </Grid>

            <Grid item xs={12}>
              <Box sx={{ display: 'flex', gap: 2, justifyContent: 'flex-end' }}>
                <Button
                  variant="outlined"
                  onClick={() => navigate('/issuing-company')}
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
                  {loading ? 'Guardando...' : 'Guardar Cambios'}
                </Button>
              </Box>
            </Grid>
          </Grid>
        </form>
      </Paper>
    </Box>
  );
};

