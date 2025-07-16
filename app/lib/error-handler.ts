import { NextResponse } from 'next/server';
import { AppError, ApiResponse } from '../types';

export class DatabaseError extends Error implements AppError {
  code: string;
  details?: any;

  constructor(message: string, code: string = 'DATABASE_ERROR', details?: any) {
    super(message);
    this.name = 'DatabaseError';
    this.code = code;
    this.details = details;
  }
}

export class ValidationError extends Error implements AppError {
  code: string;
  details?: any;

  constructor(message: string, code: string = 'VALIDATION_ERROR', details?: any) {
    super(message);
    this.name = 'ValidationError';
    this.code = code;
    this.details = details;
  }
}

export class NotFoundError extends Error implements AppError {
  code: string;
  details?: any;

  constructor(message: string, code: string = 'NOT_FOUND', details?: any) {
    super(message);
    this.name = 'NotFoundError';
    this.code = code;
    this.details = details;
  }
}

export class UnauthorizedError extends Error implements AppError {
  code: string;
  details?: any;

  constructor(message: string, code: string = 'UNAUTHORIZED', details?: any) {
    super(message);
    this.name = 'UnauthorizedError';
    this.code = code;
    this.details = details;
  }
}

export class ForbiddenError extends Error implements AppError {
  code: string;
  details?: any;

  constructor(message: string, code: string = 'FORBIDDEN', details?: any) {
    super(message);
    this.name = 'ForbiddenError';
    this.code = code;
    this.details = details;
  }
}

export function logError(error: Error, context?: string): void {
  const timestamp = new Date().toISOString();
  const errorInfo = {
    timestamp,
    name: error.name,
    message: error.message,
    stack: error.stack,
    context,
    ...(error as AppError).code && { code: (error as AppError).code },
    ...(error as AppError).details && { details: (error as AppError).details }
  };

  console.error('🚨 Error Logged:', JSON.stringify(errorInfo, null, 2));
  
  // In production, you might want to send this to a logging service
  if (process.env.NODE_ENV === 'production') {
    // Example: send to external logging service
    // await sendToLoggingService(errorInfo);
  }
}

export function handleApiError(error: unknown, context?: string): NextResponse {
  let appError: AppError;

  if (error instanceof Error) {
    // Check if it's already an AppError
    if ('code' in error) {
      appError = error as AppError;
    } else {
      // Convert generic error to AppError
      appError = {
        message: error.message,
        code: 'INTERNAL_ERROR',
        details: error.stack
      };
    }
  } else {
    // Handle non-Error objects
    appError = {
      message: String(error),
      code: 'UNKNOWN_ERROR',
      details: error
    };
  }

  // Log the error
  logError(error as Error, context);

  // Determine HTTP status code based on error type
  let statusCode = 500;
  switch (appError.code) {
    case 'VALIDATION_ERROR':
      statusCode = 400;
      break;
    case 'UNAUTHORIZED':
      statusCode = 401;
      break;
    case 'FORBIDDEN':
      statusCode = 403;
      break;
    case 'NOT_FOUND':
      statusCode = 404;
      break;
    case 'CONFLICT':
      statusCode = 409;
      break;
    case 'TOKEN_EXPIRED':
      statusCode = 401;
      break;
    case 'INVALID_TOKEN':
      statusCode = 401;
      break;
    default:
      statusCode = 500;
  }

  const response: ApiResponse = {
    success: false,
    error: appError.message,
    ...(process.env.NODE_ENV === 'development' && {
      code: appError.code,
      details: appError.details
    })
  };

  return NextResponse.json(response, { status: statusCode });
}

export function createSuccessResponse<T>(data: T, message?: string): NextResponse {
  const response: ApiResponse<T> = {
    success: true,
    data,
    ...(message && { message })
  };

  return NextResponse.json(response, { status: 200 });
}

export function validateRequiredFields(data: Record<string, any>, requiredFields: string[]): void {
  const missingFields = requiredFields.filter(field => !data[field]);
  
  if (missingFields.length > 0) {
    throw new ValidationError(
      `Missing required fields: ${missingFields.join(', ')}`,
      'VALIDATION_ERROR',
      { missingFields }
    );
  }
}

export function validateFieldType(value: any, expectedType: string, fieldName: string): void {
  const actualType = typeof value;
  
  if (actualType !== expectedType) {
    throw new ValidationError(
      `Field '${fieldName}' must be of type '${expectedType}', got '${actualType}'`,
      'VALIDATION_ERROR',
      { fieldName, expectedType, actualType, value }
    );
  }
}

export function validateStringLength(value: string, minLength: number, maxLength: number, fieldName: string): void {
  if (value.length < minLength || value.length > maxLength) {
    throw new ValidationError(
      `Field '${fieldName}' must be between ${minLength} and ${maxLength} characters`,
      'VALIDATION_ERROR',
      { fieldName, minLength, maxLength, actualLength: value.length }
    );
  }
}

export function validateNumberRange(value: number, min: number, max: number, fieldName: string): void {
  if (value < min || value > max) {
    throw new ValidationError(
      `Field '${fieldName}' must be between ${min} and ${max}`,
      'VALIDATION_ERROR',
      { fieldName, min, max, actualValue: value }
    );
  }
}

export function sanitizeString(value: string): string {
  return value.trim().replace(/[<>]/g, '');
}

export function validateEmail(email: string): void {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    throw new ValidationError('Invalid email format', 'VALIDATION_ERROR', { email });
  }
}

export function validatePhone(phone: string): void {
  const phoneRegex = /^\+?[\d\s\-\(\)]{10,}$/;
  if (!phoneRegex.test(phone)) {
    throw new ValidationError('Invalid phone number format', 'VALIDATION_ERROR', { phone });
  }
}

export function withErrorHandling<T extends any[], R>(
  fn: (...args: T) => Promise<R>,
  context?: string
) {
  return async (...args: T): Promise<R> => {
    try {
      return await fn(...args);
    } catch (error) {
      logError(error as Error, context);
      throw error;
    }
  };
}

export function createAsyncHandler<T extends any[], R>(
  handler: (...args: T) => Promise<R>,
  context?: string
) {
  return async (...args: T): Promise<R> => {
    try {
      return await handler(...args);
    } catch (error) {
      logError(error as Error, context);
      throw error;
    }
  };
} 