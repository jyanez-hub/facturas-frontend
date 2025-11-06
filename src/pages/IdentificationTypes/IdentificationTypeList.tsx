import { useState, useEffect } from 'react';
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
  Dialog,
  DialogTitle,
  DialogContent,
  DialogContentText,
  DialogActions,
} from '@mui/material';
import { Add as AddIcon, Edit as EditIcon, Delete as DeleteIcon } from '@mui/icons-material';
import { identificationTypeApi } from '../../api/identificationType';
import type { IdentificationType } from '../../types';
import { Loading } from '../../components/common/Loading';
import { ErrorAlert } from '../../components/common/ErrorAlert';

export const IdentificationTypeList = () => {
  const navigate = useNavigate();
  const [types, setTypes] = useState<IdentificationType[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const [typeToDelete, setTypeToDelete] = useState<IdentificationType | null>(null);

  const loadTypes = async () => {
    try {
      setLoading(true);
      setError(null);
      const data = await identificationTypeApi.getAll();
      setTypes(data);
    } catch (err: any) {
      setError(err.response?.data?.message || 'Error al cargar tipos de identificación');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadTypes();
  }, []);

  const handleDelete = async () => {
    if (!typeToDelete?._id) return;

    try {
      await identificationTypeApi.delete(typeToDelete._id);
      setDeleteDialogOpen(false);
      setTypeToDelete(null);
      loadTypes();
    } catch (err: any) {
      setError(err.response?.data?.message || 'Error al eliminar tipo de identificación');
    }
  };

  const openDeleteDialog = (type: IdentificationType) => {
    setTypeToDelete(type);
    setDeleteDialogOpen(true);
  };

  if (loading) return <Loading />;

  return (
    <Box>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 3 }}>
        <Typography variant="h4" component="h1">
          Tipos de Identificación
        </Typography>
        <Button
          variant="contained"
          color="primary"
          startIcon={<AddIcon />}
          onClick={() => navigate('/identification-types/new')}
        >
          Nuevo Tipo
        </Button>
      </Box>

      <ErrorAlert message={error || ''} open={!!error} onClose={() => setError(null)} />

      <TableContainer component={Paper}>
        <Table>
          <TableHead>
            <TableRow>
              <TableCell>Código</TableCell>
              <TableCell>Nombre</TableCell>
              <TableCell>Descripción</TableCell>
              <TableCell align="right">Acciones</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {types.length === 0 ? (
              <TableRow>
                <TableCell colSpan={4} align="center">
                  <Typography color="textSecondary">
                    No hay tipos de identificación registrados
                  </Typography>
                </TableCell>
              </TableRow>
            ) : (
              types.map((type) => (
                <TableRow key={type._id}>
                  <TableCell>{type.codigo}</TableCell>
                  <TableCell>{type.nombre}</TableCell>
                  <TableCell>{type.descripcion || '-'}</TableCell>
                  <TableCell align="right">
                    <IconButton
                      color="primary"
                      size="small"
                      onClick={() => navigate(`/identification-types/${type._id}`)}
                    >
                      <EditIcon />
                    </IconButton>
                    <IconButton
                      color="error"
                      size="small"
                      onClick={() => openDeleteDialog(type)}
                    >
                      <DeleteIcon />
                    </IconButton>
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </TableContainer>

      {/* Dialog de confirmación de eliminación */}
      <Dialog open={deleteDialogOpen} onClose={() => setDeleteDialogOpen(false)}>
        <DialogTitle>Confirmar Eliminación</DialogTitle>
        <DialogContent>
          <DialogContentText>
            ¿Está seguro de que desea eliminar el tipo de identificación "{typeToDelete?.nombre}"?
            Esta acción no se puede deshacer.
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setDeleteDialogOpen(false)}>Cancelar</Button>
          <Button onClick={handleDelete} color="error" variant="contained">
            Eliminar
          </Button>
        </DialogActions>
      </Dialog>
    </Box>
  );
};

