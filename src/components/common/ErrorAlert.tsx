import React from 'react';
import { Alert, AlertTitle, Snackbar } from '@mui/material';

interface ErrorAlertProps {
  message: string;
  open: boolean;
  onClose: () => void;
  severity?: 'error' | 'warning' | 'info' | 'success';
}

export const ErrorAlert: React.FC<ErrorAlertProps> = ({
  message,
  open,
  onClose,
  severity = 'error',
}) => {
  return (
    <Snackbar
      open={open}
      autoHideDuration={6000}
      onClose={onClose}
      anchorOrigin={{ vertical: 'top', horizontal: 'right' }}
    >
      <Alert onClose={onClose} severity={severity} sx={{ width: '100%' }}>
        <AlertTitle>{severity === 'error' ? 'Error' : 'Aviso'}</AlertTitle>
        {message}
      </Alert>
    </Snackbar>
  );
};

