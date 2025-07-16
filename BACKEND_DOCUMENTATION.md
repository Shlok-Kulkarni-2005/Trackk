# TrackOPZ Backend System Documentation

## Overview

This backend system provides comprehensive product lifecycle management and real-time integration between the workpanel and dispatch system. Products automatically move between "live" and "past" states based on their lifecycle events.

## Core Features

### 1. Product Lifecycle Management
- **Live Products**: Products with state 'ON' that are currently in production
- **Past Products**: Products with state 'OFF' that have completed their lifecycle
- **Automatic Transitions**: Products automatically move from live to past when dispatched or manually moved

### 2. Real-time Updates
- Server-Sent Events (SSE) for real-time product status updates
- Automatic dropdown updates in dispatch menu when products move to past
- Live synchronization between workpanel and dispatch system

### 3. Dispatch Integration
- Only live products appear in dispatch dropdown
- Automatic removal from dispatch options when moved to past
- Complete audit trail of product transitions

## API Endpoints

### 1. Live Products (Dispatchable)
**GET** `/api/live-products/dispatchable`
- Returns only products with state 'ON' that can be dispatched
- Used by dispatch page dropdown

**POST** `/api/live-products/dispatchable`
- Moves a product from live to past (state ON → OFF)

### 2. Product Lifecycle Management
**GET** `/api/products/lifecycle?productId={id}`
- Returns complete lifecycle information for a product
- Includes all stages, timestamps, and current state

**POST** `/api/products/lifecycle`
- Updates product lifecycle state
- Actions: `move_to_past`, `reactivate`, `update_state`
- Automatically broadcasts changes to connected clients

### 3. Product Transition
**POST** `/api/products/move-to-past`
- Legacy endpoint for moving products to past
- Now uses the lifecycle API internally

**GET** `/api/products/move-to-past?productId={id}`
- Returns transition history for a product

### 4. Real-time Stream
**GET** `/api/live-products/dispatchable/stream`
- Server-Sent Events endpoint for real-time updates
- Broadcasts product state changes to all connected clients

### 5. Dashboard Status
**GET** `/api/dashboard/status`
- Comprehensive system status
- Includes live/past products, machine status, and statistics

## Database Schema

### Key Tables
- **Job**: Tracks product lifecycle with state ('ON'/'OFF')
- **Product**: Product definitions
- **Machine**: Machine definitions and status
- **OperatorProductUpdate**: Finished product records

### State Management
- **ON**: Product is live and dispatchable
- **OFF**: Product is in past and not dispatchable

## Frontend Integration

### Dispatch Page (`/dispatch`)
- Fetches live products from `/api/live-products/dispatchable`
- Connects to SSE stream for real-time updates
- Automatically removes products from dropdown when moved to past
- Calls lifecycle API when dispatching products

### Workpanel (`/workpanel`)
- Shows live and past products based on job states
- Manual "Move to Past" button in product details
- Uses lifecycle API for state transitions
- Real-time updates via polling

## Real-time Communication

### Server-Sent Events
```javascript
// Connect to real-time stream
const eventSource = new EventSource('/api/live-products/dispatchable/stream');

eventSource.onmessage = (event) => {
  const data = JSON.parse(event.data);
  
  if (data.type === 'product_moved_to_past') {
    // Remove product from live products list
    // Update UI accordingly
  }
};
```

### Broadcast Events
- `product_moved_to_past`: When a product transitions from live to past
- `dispatchable_products_update`: When the list of dispatchable products changes

## Usage Examples

### Moving a Product to Past
```javascript
const response = await fetch('/api/products/lifecycle', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    productId: '123',
    jobId: '456',
    action: 'move_to_past',
    reason: 'dispatched'
  })
});
```

### Getting Product Lifecycle
```javascript
const response = await fetch('/api/products/lifecycle?productId=123');
const { lifecycle } = await response.json();
console.log(lifecycle.currentState); // 'ON' or 'OFF'
console.log(lifecycle.isDispatchable); // true or false
```

### Real-time Updates
```javascript
const eventSource = new EventSource('/api/live-products/dispatchable/stream');

eventSource.onmessage = (event) => {
  const data = JSON.parse(event.data);
  
  switch (data.type) {
    case 'product_moved_to_past':
      // Update dispatch dropdown
      removeProductFromDropdown(data.data.productId);
      break;
    case 'connection_established':
      console.log('Connected to real-time stream');
      break;
  }
};
```

## Error Handling

### Common Error Responses
- `400`: Missing required parameters
- `404`: Product or job not found
- `500`: Server error

### Error Format
```json
{
  "success": false,
  "error": "Error description"
}
```

## Security Considerations

- All endpoints validate input parameters
- Product state transitions are logged for audit purposes
- Real-time connections are properly managed and cleaned up
- Database transactions ensure data consistency

## Performance Optimizations

- Efficient database queries with proper indexing
- Real-time updates reduce polling overhead
- Connection pooling for database operations
- Caching of frequently accessed data

## Monitoring and Logging

- All product transitions are logged with timestamps
- Real-time connection status is monitored
- Database query performance is tracked
- Error rates and response times are monitored

## Future Enhancements

1. **Webhook System**: External integrations for product state changes
2. **Advanced Analytics**: Product lifecycle analytics and reporting
3. **Batch Operations**: Bulk product state updates
4. **Notification System**: Email/SMS alerts for critical state changes
5. **API Rate Limiting**: Protect against abuse
6. **Caching Layer**: Redis for improved performance

## Troubleshooting

### Common Issues

1. **Products not appearing in dispatch dropdown**
   - Check if product state is 'ON'
   - Verify job exists and is latest for product
   - Check API response for errors

2. **Real-time updates not working**
   - Verify SSE connection is established
   - Check browser console for connection errors
   - Ensure server is broadcasting events

3. **State transitions failing**
   - Verify product and job IDs are correct
   - Check if product is in correct state for transition
   - Review server logs for detailed error messages

### Debug Endpoints

- `/api/dashboard/status`: System overview
- `/api/products/lifecycle?productId={id}`: Product details
- `/api/live-products/dispatchable`: Current dispatchable products 