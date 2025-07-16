import { useState, useEffect, useCallback } from 'react';

interface AlertNotification {
  type: 'unreadCount' | 'error';
  count?: number;
  message?: string;
}

export function useAlertNotifications() {
  const [unreadCount, setUnreadCount] = useState<number>(0);
  const [isConnected, setIsConnected] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);

  // Check if user is authenticated
  const checkAuthentication = useCallback(async () => {
    try {
      const res = await fetch('/api/operator-alerts');
      if (res.status === 401) {
        setIsAuthenticated(false);
        setUnreadCount(0);
        return false;
      } else if (res.ok) {
        setIsAuthenticated(true);
        const data = await res.json();
        setUnreadCount(data.unreadCount || 0);
        return true;
      }
      return false;
    } catch (error) {
      // Silently handle auth check failures
      if (process.env.NODE_ENV === 'development') {
        console.log('Alert notifications: Authentication check failed');
      }
      setIsAuthenticated(false);
      setUnreadCount(0);
      return false;
    }
  }, []);

  const connectToSSE = useCallback(async () => {
    // Only connect if authenticated
    const authenticated = await checkAuthentication();
    if (!authenticated) {
      setError(null); // Don't show error for unauthenticated users
      return null;
    }

    try {
      const eventSource = new EventSource('/api/alerts/notifications');

      eventSource.onopen = () => {
        setIsConnected(true);
        setError(null);
      };

      eventSource.onmessage = (event) => {
        try {
          const data: AlertNotification = JSON.parse(event.data);
          if (data.type === 'unreadCount' && typeof data.count === 'number') {
            setUnreadCount(data.count);
          } else if (data.type === 'error') {
            console.error('SSE error:', data.message);
            setError(data.message || 'Unknown error');
          }
        } catch (error) {
          console.error('Error parsing SSE data:', error);
        }
      };

      eventSource.onerror = (error) => {
        console.error('SSE connection error:', error);
        setIsConnected(false);
        setError('Connection lost');
        eventSource.close();
        
        // Reconnect after 5 seconds only if still authenticated
        setTimeout(async () => {
          const stillAuthenticated = await checkAuthentication();
          if (stillAuthenticated) {
            connectToSSE();
          }
        }, 5000);
      };

      return eventSource;
    } catch (error) {
      console.error('Error creating EventSource:', error);
      setError('Failed to connect');
      return null;
    }
  }, [checkAuthentication]);

  useEffect(() => {
    let eventSource: EventSource | null = null;

    const initializeConnection = async () => {
      eventSource = await connectToSSE();
    };

    initializeConnection();

    return () => {
      if (eventSource) {
        eventSource.close();
      }
    };
  }, [connectToSSE]);

  // Fallback polling when not authenticated or SSE fails
  useEffect(() => {
    if (!isAuthenticated || (!isConnected && error)) {
      const fetchUnreadCount = async () => {
        try {
          const res = await fetch('/api/operator-alerts');
          if (res.ok) {
            const data = await res.json();
            setUnreadCount(data.unreadCount || 0);
            setIsAuthenticated(true);
            setError(null);
          } else if (res.status === 401) {
            setIsAuthenticated(false);
            setUnreadCount(0);
            setError(null); // Don't show error for unauthenticated users
          } else {
            // Silently handle other HTTP errors
            if (process.env.NODE_ENV === 'development') {
              console.log('Alert notifications: HTTP error', res.status);
            }
          }
        } catch (error) {
          // Silently handle fetch errors for unauthenticated users
          if (process.env.NODE_ENV === 'development') {
            console.log('Alert notifications: User not authenticated or network error');
          }
        }
      };

      fetchUnreadCount();
      const interval = setInterval(fetchUnreadCount, 30000); // Poll every 30 seconds
      
      return () => clearInterval(interval);
    }
  }, [isAuthenticated, isConnected, error]);

  return { unreadCount, isConnected, error, isAuthenticated };
} 