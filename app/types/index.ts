// Database Models
export interface User {
  id: number;
  email: string;
  password: string;
  name: string;
  createdAt: Date;
}

export interface Operator {
  id: number;
  phone: string;
  username?: string;
  profileImage?: string;
  createdAt: Date;
}

export interface Job {
  id: number;
  machineId: number;
  productId: number;
  state: string;
  stage: string;
  createdAt: Date;
  machine?: Machine;
  product?: Product;
}

export interface Machine {
  id: number;
  name: string;
  status: string;
  jobs?: Job[];
}

export interface Product {
  id: number;
  name: string;
  jobs?: Job[];
  productCount?: ProductCount;
}

export interface ProductCount {
  id: number;
  productId: number;
  count: number;
  machine: string;
  stage: string;
  state: string;
  createdAt: Date;
  updatedAt: Date;
  product?: Product;
}

export interface OperatorProductUpdate {
  id: number;
  operatorId: number;
  product: string;
  processSteps: Record<string, any>;
  dispatchStatus: string;
  dispatchedCost: number;
  quantity: number;
  createdAt: Date;
  archived: boolean;
  operator?: Operator;
}

export interface Alert {
  id: number;
  senderId?: number;
  message: string;
  icon?: string;
  createdAt: Date;
  recipients?: OperatorAlertStatus[];
}

export interface OperatorAlertStatus {
  id: number;
  operatorId: number;
  alertId: number;
  read: boolean;
  operator?: Operator;
  alert?: Alert;
}

export interface OTP {
  id: number;
  userId: number;
  code: string;
  expiresAt: Date;
  used: boolean;
  user?: User;
}

export interface PasswordResetToken {
  id: number;
  userId: number;
  token: string;
  expiresAt: Date;
  used: boolean;
  user?: User;
}

// API Response Types
export interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
}

export interface JobsResponse extends ApiResponse {
  jobs?: Job[];
}

export interface ProductsResponse extends ApiResponse {
  products?: Product[];
}

export interface MachinesResponse extends ApiResponse {
  machines?: Machine[];
}

export interface AlertsResponse extends ApiResponse {
  alerts?: Alert[];
}

export interface OperatorResponse extends ApiResponse {
  operator?: Operator;
}

// Form Data Types
export interface JobData {
  machine: string;
  product: string;
  state: string;
  stage: string;
  quantity: number;
}

export interface LoginData {
  email?: string;
  phone?: string;
  password?: string;
}

export interface OTPData {
  email?: string;
  phone?: string;
  otp: string;
}

// Component Props
export interface CustomDropdownProps {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
}

export interface SidebarProps {
  isOpen: boolean;
  onClose: () => void;
  username: string | null;
}

// Real-time Event Types
export interface SSEEvent {
  type: string;
  data?: any;
  timestamp?: string;
}

export interface ProductUpdateEvent extends SSEEvent {
  type: 'product_updated' | 'product_moved_to_past' | 'product_added';
  data: {
    product: Product;
    action: string;
  };
}

export interface AlertEvent extends SSEEvent {
  type: 'alert_created' | 'alert_read';
  data: {
    alert: Alert;
    operatorId?: number;
  };
}

// Dashboard Types
export interface DashboardStats {
  totalProducts: number;
  liveProducts: number;
  pastProducts: number;
  dispatchableProducts: number;
  totalMachines: number;
  activeMachines: number;
  inactiveMachines: number;
}

export interface DashboardData {
  stats: DashboardStats;
  liveProducts: Product[];
  pastProducts: Product[];
  machineStatus: Machine[];
  recentActivity: Job[];
  lastUpdated: string;
}

// JWT Token Types
export interface JWTPayload {
  userId?: number;
  email?: string;
  phone?: string;
  iat?: number;
  exp?: number;
}

// Error Types
export interface AppError {
  message: string;
  code?: string;
  details?: any;
}

// Utility Types
export type JobState = 'ON' | 'OFF';
export type JobStage = 'Milling' | 'Cutting' | 'RFD' | 'CNC Finished';
export type DispatchStatus = 'Pending' | 'Dispatched' | 'Completed';
export type AlertStatus = 'unread' | 'read'; 