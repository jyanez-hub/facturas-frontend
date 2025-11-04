import React from 'react';
import { Box, CircularProgress } from '@mui/material';

interface LoadingProps {
  fullScreen?: boolean;
}

export const Loading: React.FC<LoadingProps> = ({ fullScreen = false }) => {
  return (
    <Box
      display="flex"
      justifyContent="center"
      alignItems="center"
      minHeight={fullScreen ? '100vh' : '200px'}
    >
      <CircularProgress />
    </Box>
  );
};

