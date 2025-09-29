export class AuthService {
  private static readonly AUTH_KEY = 'bonus_system_auth';
  
  static isAuthenticated(): boolean {
    if (typeof window === 'undefined') return false;
    return localStorage.getItem(this.AUTH_KEY) === 'true';
  }
  
  static authenticate(secret: string): boolean {
    const isValid = secret === process.env.AUTH_SECRET;
    if (isValid && typeof window !== 'undefined') {
      localStorage.setItem(this.AUTH_KEY, 'true');
    }
    return isValid;
  }
  
  static logout(): void {
    if (typeof window !== 'undefined') {
      localStorage.removeItem(this.AUTH_KEY);
    }
  }
}
