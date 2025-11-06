import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Box,
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
} from '@mui/material';
import { Edit as EditIcon, Business as BusinessIcon } from '@mui/icons-material';
import { issuingCompanyApi } from '../../api/issuingCompany';
import type { IssuingCompany } from '../../types';
import { Loading } from '../../components/common/Loading';
import { ErrorAlert } from '../../components/common/ErrorAlert';

export const IssuingCompanyList = () => {
  const navigate = useNavigate();
  const [companies, setCompanies] = useState<IssuingCompany[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const loadCompanies = async () => {
    try {
      setLoading(true);
      setError(null);
      const data = await issuingCompanyApi.getAll();
      setCompanies(data);
    } catch (err: any) {
      setError(err.response?.data?.message || 'Error al cargar empresas emisoras');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadCompanies();
  }, []);

  if (loading) return <Loading />;

  return (
    <Box>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 3 }}>
        <Typography variant="h4" component="h1">
          Empresas Emisoras
        </Typography>
      </Box>

      <ErrorAlert message={error || ''} open={!!error} onClose={() => setError(null)} />

      <TableContainer component={Paper}>
        <Table>
          <TableHead>
            <TableRow>
              <TableCell>RUC</TableCell>
              <TableCell>Razón Social</TableCell>
              <TableCell>Nombre Comercial</TableCell>
              <TableCell>Ambiente</TableCell>
              <TableCell>Obligado Contabilidad</TableCell>
              <TableCell align="right">Acciones</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {companies.length === 0 ? (
              <TableRow>
                <TableCell colSpan={6} align="center">
                  <Box sx={{ py: 4, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
                    <BusinessIcon sx={{ fontSize: 60, color: 'text.secondary' }} />
                    <Typography color="textSecondary">
                      No hay empresas emisoras registradas
                    </Typography>
                    <Typography variant="body2" color="textSecondary">
                      Las empresas se crean automáticamente durante el registro de usuario
                    </Typography>
                  </Box>
                </TableCell>
              </TableRow>
            ) : (
              companies.map((company) => (
                <TableRow key={company._id}>
                  <TableCell>{company.ruc}</TableCell>
                  <TableCell>{company.razon_social}</TableCell>
                  <TableCell>{company.nombre_comercial || '-'}</TableCell>
                  <TableCell>
                    <Chip
                      label={company.tipo_ambiente === 1 ? 'Pruebas' : 'Producción'}
                      color={company.tipo_ambiente === 1 ? 'warning' : 'success'}
                      size="small"
                    />
                  </TableCell>
                  <TableCell>
                    <Chip
                      label={company.obligado_contabilidad ? 'Sí' : 'No'}
                      color={company.obligado_contabilidad ? 'info' : 'default'}
                      size="small"
                    />
                  </TableCell>
                  <TableCell align="right">
                    <IconButton
                      color="primary"
                      size="small"
                      onClick={() => navigate(`/issuing-company/${company._id}`)}
                    >
                      <EditIcon />
                    </IconButton>
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </TableContainer>
    </Box>
  );
};

