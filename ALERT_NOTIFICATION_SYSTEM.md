# Alert Notification System

## Overview

The alert notification system provides real-time notifications for operators in the TrackOPZ application. When new alerts are sent, operators see a red badge with the number of unread alerts on the "See Alerts" button. Once they view the alerts page, the notifications are automatically marked as read and the badge disappears.

## Features

- ✅ Real-time unread alert count
- ✅ Red notification badge on alerts button
- ✅ Automatic marking of alerts as read when viewed
- ✅ Server-Sent Events (SSE) for instant updates
- ✅ JWT-based authentication for operator identification
- ✅ Automatic reconnection on connection loss

## Architecture

### Database Schema

The system uses the existing database schema with these key tables:

- `Operator`: Stores operator information
- `Alert`: Stores alert messages
- `OperatorAlertStatus`: Tracks read/unread status for each operator-alert pair

### API Endpoints

1. **`GET /api/operator-alerts`**: Get unread alert count for current operator
2. **`POST /api/operator-alerts`**: Mark alerts as read for current operator
3. **`GET /api/alerts/notifications`**: SSE endpoint for real-time updates
4. **`GET /api/alerts`**: Get all alerts (existing)
5. **`POST /api/alerts`**: Send new alert (existing)

### Authentication

The system uses JWT tokens stored in HTTP-only cookies to identify operators. The token contains the operator's phone number, which is used to look up the operator ID.

## Implementation Details

### 1. Operator Authentication (`app/lib/operator-auth.ts`)

```typescript
export async function getCurrentOperatorId(): Promise<number | null> {
  const operator = await getCurrentOperator();
  return operator?.id || null;
}
```

This utility function:
- Extracts the JWT token from cookies
- Decodes the token to get the operator's phone number
- Looks up the operator in the database
- Returns the operator ID

### 2. Real-time Notifications (`app/lib/useAlertNotifications.ts`)

```typescript
export function useAlertNotifications() {
  const [unreadCount, setUnreadCount] = useState<number>(0);
  const [isConnected, setIsConnected] = useState<boolean>(false);
  
  // SSE connection logic
  // Automatic reconnection
  // Real-time updates
}
```

This React hook:
- Establishes SSE connection to `/api/alerts/notifications`
- Receives real-time unread count updates
- Automatically reconnects on connection loss
- Returns current unread count and connection status

### 3. SSE Endpoint (`app/api/alerts/notifications/route.ts`)

```typescript
export async function GET(req: NextRequest) {
  const operatorId = await getCurrentOperatorId();
  
  const stream = new ReadableStream({
    start(controller) {
      // Send initial count
      // Set up polling every 5 seconds
      // Handle connection cleanup
    }
  });
}
```

This endpoint:
- Authenticates the operator
- Creates a persistent SSE connection
- Sends initial unread count
- Polls for updates every 5 seconds
- Cleans up on connection close

## Usage

### In React Components

```typescript
import { useAlertNotifications } from '../lib/useAlertNotifications';

function MyComponent() {
  const { unreadCount } = useAlertNotifications();
  
  return (
    <button className="relative">
      <Bell className="w-6 h-6" />
      {unreadCount > 0 && (
        <div className="absolute -top-1 -right-1 w-5 h-5 bg-red-500 rounded-full flex items-center justify-center">
          <span className="text-white text-xs">{unreadCount}</span>
        </div>
      )}
    </button>
  );
}
```

### Marking Alerts as Read

When an operator visits the alerts page, alerts are automatically marked as read:

```typescript
useEffect(() => {
  async function markAlertsAsRead() {
    await fetch('/api/operator-alerts', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
    });
  }
  
  markAlertsAsRead();
}, []);
```

## Testing

Run the test script to verify the system works correctly:

```bash
node scripts/test-alert-notifications.js
```

This script:
1. Creates a test operator
2. Creates a test alert
3. Verifies unread count
4. Marks alerts as read
5. Verifies count is zero
6. Cleans up test data

## Security Considerations

- JWT tokens are stored in HTTP-only cookies
- Operator authentication is required for all alert operations
- SSE connections are authenticated
- Database queries use parameterized queries to prevent injection

## Performance Considerations

- SSE connections are lightweight and efficient
- Unread count is cached and updated every 5 seconds
- Database queries are optimized with proper indexing
- Automatic reconnection prevents stale connections

## Troubleshooting

### Common Issues

1. **Badge not showing**: Check if operator is authenticated
2. **Count not updating**: Verify SSE connection is established
3. **Alerts not marked as read**: Check network connectivity
4. **Authentication errors**: Verify JWT token is valid

### Debug Steps

1. Check browser console for errors
2. Verify SSE connection in Network tab
3. Check database for alert records
4. Verify operator authentication

## Future Enhancements

- Push notifications for mobile devices
- Email notifications for critical alerts
- Alert categories and priority levels
- Custom notification preferences per operator
- Alert history and analytics 