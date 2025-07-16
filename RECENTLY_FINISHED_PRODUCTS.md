# Recently Finished Products Functionality

## Overview

The "Recently Finished Products" dropdown in the operator details page (`/updatedetailsop`) now automatically updates when products are added to or removed from the live products category in the workpanel.

## How It Works

### 1. Product Addition to Live Category
- When a product is added via the `/addjobs` page with state 'ON'
- The product automatically appears in the "Recently Finished Products" dropdown
- Real-time updates are broadcast to all connected clients

### 2. Product Movement to Past Category
- When a product is moved to past (via workpanel "Move to Past" button or dispatch)
- The product is automatically removed from the "Recently Finished Products" dropdown
- Real-time updates are broadcast to all connected clients

## API Endpoints

### `/api/finished-products/dispatchable`
- **GET**: Returns products that should appear in the "Recently Finished Products" dropdown
- Includes both live products (state 'ON') and finished products from OperatorProductUpdate table
- Products are sorted by creation date (most recent first)

### `/api/finished-products/dispatchable/stream`
- **GET**: Server-Sent Events endpoint for real-time updates
- Broadcasts events when products are added or removed

## Real-time Events

### `product_added_to_live`
Triggered when a product is added with state 'ON'
```json
{
  "type": "product_added_to_live",
  "data": {
    "productId": "123",
    "productName": "Product A",
    "machineName": "Cutting MC/1",
    "timestamp": "2025-01-07T10:30:00.000Z"
  }
}
```

### `product_moved_to_past`
Triggered when a product is moved from live to past
```json
{
  "type": "product_moved_to_past",
  "data": {
    "productId": "123",
    "productName": "Product A",
    "timestamp": "2025-01-07T10:35:00.000Z"
  }
}
```

## Frontend Integration

### Updatedetailsop Page (`/updatedetailsop`)
- Connects to real-time stream for automatic updates
- Refreshes dropdown when products are added or removed
- Shows debugging information including product source

### Workpanel Page (`/workpanel`)
- "Move to Past" button triggers product removal from dropdown
- Uses lifecycle API for state transitions
- Automatically refreshes product list

## Testing

Run the test script to verify functionality:
```bash
node scripts/test-finished-products.js
```

This will:
1. Check current state of products
2. Add a new live product
3. Verify it appears in Recently Finished Products
4. Move it to past
5. Verify it's removed from Recently Finished Products
6. Clean up test data

## Database Schema

### Job Table
- Tracks product lifecycle with state ('ON'/'OFF')
- Live products (state 'ON') appear in dropdown
- Past products (state 'OFF') are excluded from dropdown

### OperatorProductUpdate Table
- Stores finished product records
- Products from this table also appear in dropdown

## Configuration

The system automatically:
- Includes live products (state 'ON') in the dropdown
- Excludes past products (state 'OFF') from the dropdown
- Removes products from OperatorProductUpdate when moved to past
- Broadcasts real-time updates to all connected clients

## Troubleshooting

### Products not appearing in dropdown
1. Check if product state is 'ON' in Job table
2. Verify product was created via `/api/jobs` endpoint
3. Check browser console for real-time connection errors

### Products not removed from dropdown
1. Check if product state was changed to 'OFF'
2. Verify lifecycle API call was successful
3. Check if product exists in OperatorProductUpdate table

### Real-time updates not working
1. Check if SSE connection is established
2. Verify `/api/finished-products/dispatchable/stream` endpoint is accessible
3. Check browser console for connection errors 