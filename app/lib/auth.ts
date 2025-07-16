import jwt from 'jsonwebtoken';
import { cookies } from 'next/headers';
import { JWTPayload, AppError } from '../types';

const JWT_SECRET = process.env.JWT_SECRET!;
const JWT_EXPIRES_IN = '7d';

export class AuthError extends Error implements AppError {
  code: string;
  details?: any;

  constructor(message: string, code: string = 'AUTH_ERROR', details?: any) {
    super(message);
    this.name = 'AuthError';
    this.code = code;
    this.details = details;
  }
}

export function generateToken(payload: Omit<JWTPayload, 'iat' | 'exp'>): string {
  try {
    if (!JWT_SECRET) {
      throw new AuthError('JWT_SECRET is not configured', 'CONFIG_ERROR');
    }

    return jwt.sign(payload, JWT_SECRET, { 
      expiresIn: JWT_EXPIRES_IN,
      issuer: 'trackopz',
      audience: 'trackopz-users'
    });
  } catch (error) {
    if (error instanceof AuthError) {
      throw error;
    }
    throw new AuthError('Failed to generate token', 'TOKEN_GENERATION_ERROR', error);
  }
}

export function verifyToken(token: string): JWTPayload {
  try {
    if (!JWT_SECRET) {
      throw new AuthError('JWT_SECRET is not configured', 'CONFIG_ERROR');
    }

    const decoded = jwt.verify(token, JWT_SECRET, {
      issuer: 'trackopz',
      audience: 'trackopz-users'
    }) as JWTPayload;

    // Check if token is expired
    if (decoded.exp && Date.now() >= decoded.exp * 1000) {
      throw new AuthError('Token has expired', 'TOKEN_EXPIRED');
    }

    return decoded;
  } catch (error) {
    if (error instanceof AuthError) {
      throw error;
    }
    
    if (error instanceof jwt.JsonWebTokenError) {
      throw new AuthError('Invalid token', 'INVALID_TOKEN', error.message);
    }
    
    if (error instanceof jwt.TokenExpiredError) {
      throw new AuthError('Token has expired', 'TOKEN_EXPIRED');
    }
    
    throw new AuthError('Token verification failed', 'TOKEN_VERIFICATION_ERROR', error);
  }
}

export async function getTokenFromCookies(): Promise<string | null> {
  try {
    const cookieStore = await cookies();
    return cookieStore.get('token')?.value || null;
  } catch (error) {
    console.error('Error getting token from cookies:', error);
    return null;
  }
}

export async function getCurrentUser(): Promise<JWTPayload | null> {
  try {
    const token = await getTokenFromCookies();
    if (!token) {
      return null;
    }

    return verifyToken(token);
  } catch (error) {
    console.error('Error getting current user:', error);
    return null;
  }
}

export function isTokenExpired(token: string): boolean {
  try {
    const decoded = jwt.decode(token) as JWTPayload;
    return decoded.exp ? Date.now() >= decoded.exp * 1000 : true;
  } catch {
    return true;
  }
}

export function getTokenExpiration(token: string): Date | null {
  try {
    const decoded = jwt.decode(token) as JWTPayload;
    return decoded.exp ? new Date(decoded.exp * 1000) : null;
  } catch {
    return null;
  }
}

export function refreshTokenIfNeeded(token: string, thresholdMinutes: number = 60): string | null {
  try {
    const decoded = jwt.decode(token) as JWTPayload;
    if (!decoded.exp) return null;

    const expirationTime = decoded.exp * 1000;
    const thresholdTime = Date.now() + (thresholdMinutes * 60 * 1000);

    if (expirationTime <= thresholdTime) {
      // Token will expire soon, generate new one
      const { iat, exp, ...payload } = decoded;
      return generateToken(payload);
    }

    return null; // Token is still valid
  } catch {
    return null;
  }
}

export function validatePassword(password: string): { isValid: boolean; errors: string[] } {
  const errors: string[] = [];
  
  if (password.length < 8) {
    errors.push('Password must be at least 8 characters long');
  }
  
  if (!/[A-Z]/.test(password)) {
    errors.push('Password must contain at least one uppercase letter');
  }
  
  if (!/[a-z]/.test(password)) {
    errors.push('Password must contain at least one lowercase letter');
  }
  
  if (!/\d/.test(password)) {
    errors.push('Password must contain at least one number');
  }
  
  if (!/[!@#$%^&*(),.?":{}|<>]/.test(password)) {
    errors.push('Password must contain at least one special character');
  }
  
  return {
    isValid: errors.length === 0,
    errors
  };
}

export function validateEmail(email: string): boolean {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

export function validatePhone(phone: string): boolean {
  const phoneRegex = /^\+?[\d\s\-\(\)]{10,}$/;
  return phoneRegex.test(phone);
}

export function sanitizeInput(input: string): string {
  return input.trim().replace(/[<>]/g, '');
}

export function generateSecureOTP(length: number = 6): string {
  const digits = '0123456789';
  let otp = '';
  
  for (let i = 0; i < length; i++) {
    otp += digits[Math.floor(Math.random() * digits.length)];
  }
  
  return otp;
}

export function hashPassword(password: string): Promise<string> {
  return import('bcryptjs').then(bcrypt => bcrypt.hash(password, 12));
}

export function comparePassword(password: string, hash: string): Promise<boolean> {
  return import('bcryptjs').then(bcrypt => bcrypt.compare(password, hash));
} 