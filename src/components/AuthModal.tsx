'use client';

import { useState } from 'react';
import { Lock } from 'lucide-react';

interface AuthModalProps {
  onAuthenticate: (secret: string) => void;
  isLoading: boolean;
  error?: string;
}

export default function AuthModal({ onAuthenticate, isLoading, error }: AuthModalProps) {
  const [secret, setSecret] = useState('');

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onAuthenticate(secret);
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center">
      <div className="bg-white p-8 rounded-lg shadow-xl w-96">
        <div className="flex items-center mb-6">
          <Lock className="h-8 w-8 text-blue-600 mr-3" />
          <h2 className="text-2xl font-bold">Admin Access Required</h2>
        </div>
        
        <form onSubmit={handleSubmit}>
          <div className="mb-4">
            <label className="block text-gray-700 text-sm font-bold mb-2">
              Enter Auth Secret
            </label>
            <input
              type="password"
              value={secret}
              onChange={(e) => setSecret(e.target.value)}
              className="w-full px-3 py-2 border rounded-lg focus:outline-none focus:border-blue-500"
              placeholder="Enter secret key"
              required
            />
          </div>
          
          {error && (
            <div className="mb-4 text-red-600 text-sm">{error}</div>
          )}
          
          <button
            type="submit"
            disabled={isLoading}
            className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-lg disabled:opacity-50"
          >
            {isLoading ? 'Authenticating...' : 'Access System'}
          </button>
        </form>
      </div>
    </div>
  );
}
