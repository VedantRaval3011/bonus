'use client';

import { useState, useEffect } from 'react';
import AuthModal from '@/components/AuthModal';
import Dashboard from '@/components/Dashboard';
import { AuthService } from '@/lib/auth';

export default function Home() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [authError, setAuthError] = useState('');

  useEffect(() => {
    setIsAuthenticated(AuthService.isAuthenticated());
  }, []);

  const handleAuthenticate = async (secret: string) => {
    setIsLoading(true);
    setAuthError('');

    try {
      const response = await fetch('/api/auth', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ secret }),
      });

      const data = await response.json();

      if (data.success) {
        AuthService.authenticate(secret);
        setIsAuthenticated(true);
      } else {
        setAuthError('Invalid secret key');
      }
    } catch (error) {
      setAuthError('Authentication failed');
    } finally {
      setIsLoading(false);
    }
  };

  const handleLogout = () => {
    AuthService.logout();
    setIsAuthenticated(false);
  };

  if (!isAuthenticated) {
    return (
      <AuthModal
        onAuthenticate={handleAuthenticate}
        isLoading={isLoading}
        error={authError}
      />
    );
  }

  return <Dashboard onLogout={handleLogout} />;
}
