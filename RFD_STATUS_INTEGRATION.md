# RFD Status Integration Documentation

## Overview

The RFD Status Integration automatically decreases product counts in the RFD status tab when products are dispatched from the updatedetailsop tab. This ensures real-time synchronization between the dispatch process and the RFD status display.

## How It Works

### 1. Product Dispatch Process
- When an operator dispatches a product from the updatedetailsop tab
- The system automatically reduces the product count in the RFD status tab
- Real-time updates are broadcast to all connected clients
- Data consistency is maintained between dispatch and RFD status

### 2. RFD Status Count Management
- RFD status counts are based on jobs with RFD machine and OFF state
- When products are dispatched, the oldest RFD jobs are removed (FIFO)
- Counts are updated in real-time across all connected clients
- Products are removed from RFD status when count reaches zero

## API Endpoints

### 1. `/api/rfd-status/update-count`

**POST** - Update RFD status count when products are dispatched
```json
{
  "productId": 123,
  "quantity": 2
}
```

**Response:**
```json
{
  "success": true,
  "message": "Successfully reduced RFD count by 2",
  "remainingCount": 3,
  "productName": "Product A"
}
```

**GET** - Get current RFD count for a specific product
```
/api/rfd-status/update-count?productId=123
```

**Response:**
```json
{
  "success": true,
  "productId": 123,
  "rfdCount": 5
}
```

### 2. `/api/rfd-status/broadcast`

**POST** - Broadcast RFD status updates to connected clients
```json
{
  "type": "rfd_count_updated",
  "productId": 123,
  "productName": "Product A",
  "quantityReduced": 2,
  "remainingCount": 3
}
```

### 3. `/api/rfd-status/stream`

**GET** - Server-Sent Events endpoint for real-time RFD status updates

**Events:**
- `connection_established`: Initial connection message
- `rfd_status_update`: RFD status data update

## Integration Flow

### 1. Dispatch Process
1. Operator selects a product in updatedetailsop tab
2. Operator sets quantity and process steps
3. Operator clicks "Dispatch" button
4. `/api/operator/update` processes the dispatch
5. System automatically calls `/api/rfd-status/update-count`
6. RFD status count is reduced by the dispatched quantity
7. Real-time update is broadcast to all connected clients

### 2. RFD Status Display
1. RFD status tab connects to `/api/rfd-status/stream`
2. When RFD counts are updated, real-time notification is received
3. RFD status tab refreshes to show updated counts
4. Products with zero count are automatically removed from display

## Database Schema

The integration uses the existing database schema:

```sql
-- Jobs table tracks RFD status
model Job {
  id        Int      @id @default(autoincrement())
  machine   Machine  @relation(fields: [machineId], references: [id])
  machineId Int
  product   Product  @relation(fields: [productId], references: [id])
  productId Int
  state     String   -- 'OFF' for RFD status
  stage     String   -- 'RFD' for RFD stage
  createdAt DateTime @default(now())
}

-- OperatorProductUpdate tracks dispatched items
model OperatorProductUpdate {
  id             Int      @id @default(autoincrement())
  operatorId     Int
  product        String
  processSteps   Json
  dispatchStatus String
  quantity       Int      @default(1)
  createdAt      DateTime @default(now())
  archived       Boolean  @default(false)
}
```

## Real-time Updates

### Server-Sent Events
The system uses Server-Sent Events (SSE) for real-time updates:

1. **Connection**: Clients connect to `/api/rfd-status/stream`
2. **Broadcasting**: Updates are broadcast via `/api/rfd-status/broadcast`
3. **Reception**: Clients receive real-time updates and refresh their displays

### Event Types
- `rfd_status_update`: Sent when RFD counts are modified
- `connection_established`: Sent when client connects to stream

## Frontend Integration

### RFD Status Tab (`/procount`)
- Connects to RFD status stream for real-time updates
- Automatically refreshes when RFD counts are updated
- Shows current product counts with real-time accuracy

### Updatedetailsop Tab (`/updatedetailsop`)
- Dispatches products and triggers RFD count reduction
- Maintains data consistency with RFD status
- Provides immediate feedback on dispatch success

## Error Handling

### Insufficient Quantity
- System validates available RFD quantity before dispatch
- Returns error if requested quantity exceeds available count
- Prevents negative counts and data inconsistency

### Network Failures
- Graceful handling of broadcast failures
- Fallback to polling for updates if real-time fails
- Logging of errors for debugging

## Testing

Run the test script to verify functionality:
```bash
node scripts/test-rfd-status-integration.js
```

This will:
1. Check current RFD status
2. Create test products and RFD jobs
3. Test product dispatch process
4. Verify RFD count reduction
5. Test real-time updates
6. Clean up test data

## Configuration

The system automatically:
- Reduces RFD counts when products are dispatched
- Broadcasts real-time updates to all connected clients
- Maintains data consistency between dispatch and RFD status
- Handles edge cases like insufficient quantity

## Benefits

1. **Real-time Synchronization**: RFD status updates immediately when products are dispatched
2. **Data Consistency**: Ensures dispatch and RFD status are always in sync
3. **User Experience**: Operators see immediate feedback on their actions
4. **Scalability**: Server-Sent Events provide efficient real-time updates
5. **Reliability**: Robust error handling and validation

## Troubleshooting

### RFD counts not updating
1. Check if RFD jobs exist for the product
2. Verify dispatch process completed successfully
3. Check browser console for real-time connection errors
4. Ensure RFD machine and OFF state jobs exist

### Real-time updates not working
1. Check if SSE connection is established
2. Verify `/api/rfd-status/stream` endpoint is accessible
3. Check browser console for connection errors
4. Ensure broadcast endpoint is functioning

### Dispatch failures
1. Verify sufficient RFD quantity is available
2. Check product ID and quantity parameters
3. Ensure all required fields are provided
4. Check database connectivity and permissions 